// DSH AI 面板后端：同步调本机 dsh CLI（headless 单发），独立队列/会话/设置，与 plugin-ai 互不影响
// 路由：POST /api/dsh/chat、GET /api/dsh/chat-progress?reqId=、POST /api/dsh/stop、
//       POST /api/dsh/queue、GET /api/dsh/pending?session=、POST|GET /api/dsh/session、
//       GET|POST /api/dsh/settings、POST /api/dsh/log、POST /api/dsh/context、POST /api/dsh/like
import { readFileSync, writeFileSync, existsSync } from 'node:fs'
import { spawn } from 'node:child_process'
import crypto from 'node:crypto'
import os from 'node:os'
import path from 'node:path'

// === DSH：本机 agent harness（Windows 侧 dsh CLI，headless 单发即走）===
// 协议定案：DSH web 无 REST，唯一稳定入口是官方 CLI headless 应用——
//   dsh --profile headless "<任务>"   回复走 stdout，诊断走 stderr
//   --json                            stdout 改为 NDJSON 事件流（session 事件带会话号，final 事件带回复）
//   --session-id <id>                 续会话；只认已存在的 Session，未知 id 直接报错（新会话号须由 dsh 发放）
//   `-`                               从 stdin 读任务文本
const DSH_TIMEOUT_MS = Number(process.env.DSH_TIMEOUT_MS) || 600000  // headless 单发兜底超时（agent 任务实测可到 5 分钟）
// DSH_CMD：完整命令行模板（按空格分词）。默认直接 `dsh`（win32 生产 / Linux 服务器同款）；
// 只有 WSL 开发机这种「node 在 Linux、dsh 在 Windows」的跨侧环境才需要在 .env 里显式指定 interop 命令
const DSH_CMD = (process.env.DSH_CMD || '').trim() || 'dsh'
// DSH_CWD：dsh 工作目录。默认用户主目录；WSL 开发机须指向 C 盘用户目录
// （避开 UNC 警告，且 Session 按工作目录归档 ~/.dsh/sessions/，续会话须同目录）
const DSH_CWD = (process.env.DSH_CWD || '').trim() || os.homedir()

// 超时杀进程树：win32 直接 taskkill；WSL interop 经 winpid taskkill（信号杀不掉 Windows 实体）；纯 POSIX 组 SIGKILL
function killChildTree(child) {
  try {
    if (process.platform === 'win32') {
      spawn('taskkill', ['/pid', String(child.pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
    } else {
      // WSL interop（cmd.exe/dsh 等 Windows 进程）：POSIX 组信号杀不掉 Windows 侧实体，
      // 须读 /proc/<pid>/winpid 用 Windows taskkill 按真 PID 杀树；拿不到 winpid 再退回组杀
      let winpid = 0
      try { winpid = parseInt(readFileSync(`/proc/${child.pid}/winpid`, 'utf8'), 10) || 0 } catch {}
      if (winpid > 0) {
        spawn('cmd.exe', ['/c', 'taskkill', '/pid', String(winpid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
      } else {
        try { process.kill(-child.pid, 'SIGKILL') } catch { child.kill('SIGKILL') }
      }
    }
  } catch { try { child.kill('SIGKILL') } catch {} }
}

export function register(ctx) {
  const PROTOCOL_PROMPT = ctx.prompt
  const { cors, readBody, pushLog, tryParseActions, docContext } = ctx

  // DSH 面板设置：存 .settings-dsh.json（兼容旧部署：文件仍落 server/ 目录旁），POST 即时生效、优先于 env
  const DSH_SETTINGS_FILE = ctx.dataDir ? path.join(ctx.dataDir, '.settings-dsh.json') : path.join(ctx.serverDir, '.settings-dsh.json')
  let dshSettings = {}
  try { dshSettings = JSON.parse(readFileSync(DSH_SETTINGS_FILE, 'utf8')) } catch {}
  function saveDshSettings() {
    try { writeFileSync(DSH_SETTINGS_FILE, JSON.stringify(dshSettings, null, 2)) } catch {}
  }
  function effectiveDshConfig() {
    const cmd = String(dshSettings.dshCmd || '').trim() || DSH_CMD
    const sec = Number(dshSettings.dshTimeout)
    return { cmd, timeoutMs: sec > 0 ? sec * 1000 : DSH_TIMEOUT_MS }
  }

  // DSH 独立内存队列：与 plugin-ai 的 pendingQueue 互不影响
  const dshQueue = new Map()
  let dshProgress = null // 当前 chat 的思考过程 {reqId, lines[]}，面板轮询读取
  let dshChild = null    // 当前 dsh 子进程，供 /api/dsh/stop 终止
  let dshStopped = false // 本次 chat 是否被主动停止（用于给 reject 一个友好文案）
  function dshQueuePut(session, instruction) {
    if (!dshQueue.has(session)) dshQueue.set(session, [])
    dshQueue.get(session).push({ instruction, ts: Date.now() })
  }
  function dshQueueTake(session) {
    const q = dshQueue.get(session)
    if (!q || q.length === 0) return null
    return q.shift()
  }
  // DSH 活跃会话：持久化（重启服务不丢，否则面板显示「无session」要重聊一次才恢复）
  const DSH_SESSION_FILE = ctx.dataDir ? path.join(ctx.dataDir, 'dsh-active-session.txt') : '/tmp/oo-dsh-active-session'
  let dshActiveSession = ''
  try { dshActiveSession = readFileSync(DSH_SESSION_FILE, 'utf8').trim() } catch {}
  function saveDshSession(s) {
    dshActiveSession = s
    try { writeFileSync(DSH_SESSION_FILE, s) } catch {}
  }

  // 调 DSH CLI：resolve({ reply, sessionId }) / reject(Error，message 带 stderr 摘要)。
  // 统一用 --json 事件流：首行 {"type":"session","sessionId":…} 是 dsh 发放的真实会话号
  // （--session-id 只认已存在会话，故新会话不能自造 id，必须由 dsh 发放），
  // 末行 {"type":"final","text":…} 是回复。消息统一经 stdin（`-` 参数）传入：
  // 绕开 cmd.exe 的引号/特殊字符转义与命令行长度上限，也无注入风险
  function callDSH(message, sessionId, progress) {
    const cfg = effectiveDshConfig()
    const argv = cfg.cmd.split(/\s+/).filter(Boolean)
    argv.push('--profile', 'headless', '--json')
    if (sessionId) argv.push('--session-id', String(sessionId))
    argv.push('-')
    const cwd = existsSync(DSH_CWD) ? DSH_CWD : undefined
    return new Promise((resolve, reject) => {
      let child
      try {
        // 不用 detached：WSL interop 下 detached 会导致子进程退出后 close 事件永不触发（输出已完结但句柄不关）；
        // 进程树清理由 killChildTree 按 winpid taskkill 兜底，无需进程组
        child = spawn(argv[0], argv.slice(1), { cwd, windowsHide: true })
      } catch (e) {
        reject(new Error(`DSH 启动失败（${argv[0]}）: ${e.message}`)); return
      }
      dshChild = child; dshStopped = false  // 登记 /api/dsh/stop 用
      let stdout = '', stderr = '', settled = false, timer = null, dshSession = '', lineBuf = ''
      const t0 = Date.now()
      const recordProgress = (ev) => {
        if (!progress) return
        if (ev.type === 'session' && ev.sessionId) {
          // 连接完成即出提示：区分「CLI 启动/连接耗时」与「模型思考」，面板不再一直停在「连接 DSH…」
          progress.lines.push('✅ 已连接 DSH（' + ((Date.now() - t0) / 1000).toFixed(1) + 's），模型思考中…')
        } else if (ev.type === 'thinking' && typeof ev.text === 'string' && ev.text.trim()) {
          progress.lines.push('💭 ' + ev.text.replace(/\s+/g, ' ').trim().slice(0, 140))
        } else if (ev.type === 'tool_call' && ev.tool) {
          progress.lines.push('🔧 ' + ev.tool + ' ' + JSON.stringify(ev.input ?? '').slice(0, 120))
        }
        if (progress.lines.length > 200) progress.lines.splice(0, progress.lines.length - 200)
      }
      const finish = (fn, arg) => { if (settled) return; settled = true; clearTimeout(timer); dshChild = null; fn(arg) }
      timer = setTimeout(() => {
        killChildTree(child)
        finish(reject, new Error(`DSH 执行超时（${Math.round(cfg.timeoutMs / 1000)}s），进程已终止`))
      }, cfg.timeoutMs)
      // 解析 NDJSON 事件流：type=session 取会话号，type=final 取回复文本
      // 【关键】成功路径不等 close：WSL2 interop 子进程退出后 node 收不到 reap 通知，
      // close 事件可能永不触发（进程实体已死、final 已输出）。解析到 final 即 resolve，
      // 然后尽力清理句柄；close 仅兜底失败路径（非 0 退出码/空回复）
      const parseEvents = () => {
        let reply = ''
        for (const line of stdout.split(/\r?\n/)) {
          if (!line.trim()) continue
          try {
            const ev = JSON.parse(line)
            if (ev.type === 'session' && ev.sessionId) dshSession = String(ev.sessionId)
            if (ev.type === 'final' && typeof ev.text === 'string') reply = ev.text.trim()
          } catch {}
        }
        return reply
      }
      child.on('error', (e) => finish(reject, new Error(`DSH 启动失败（${argv[0]}）: ${e.message}`)))
      child.stdout.on('data', (c) => {
        stdout += c
        // 思考过程：按完整行增量解析 thinking / tool_call 事件喂给进度条
        if (progress) {
          lineBuf += c
          const idx = lineBuf.lastIndexOf('\n')
          if (idx >= 0) {
            for (const line of lineBuf.slice(0, idx).split(/\r?\n/)) {
              if (!line.trim()) continue
              try { recordProgress(JSON.parse(line)) } catch {}
            }
            lineBuf = lineBuf.slice(idx + 1)
          }
        }
        const reply = parseEvents()
        if (reply) {
          finish(resolve, { reply, sessionId: dshSession || null })
          try { child.kill('SIGKILL') } catch {}  // 进程多半已自退，仅清理僵留句柄
        }
      })
      child.stderr.on('data', (c) => (stderr += c))
      child.stdin.on('error', () => {})  // dsh 提前退出时的 EPIPE 不许炸掉服务进程
      child.stdin.end(message)
      child.on('close', (code) => {
        if (settled) return
        const reply = parseEvents()
        if (code === 0 && reply) { finish(resolve, { reply, sessionId: dshSession || null }); return }
        if (dshStopped) { finish(reject, new Error('已停止')); return }
        // 失败/空回复：取 stderr 最后几行作摘要，够定位即可
        const tail = stderr.trim().split(/\r?\n/).filter(Boolean).slice(-3).join(' / ').slice(0, 300)
        finish(reject, new Error(
          code === 0
            ? `DSH 返回空回复${tail ? '（stderr: ' + tail + '）' : ''}`
            : `DSH 退出码 ${code}${tail ? '，stderr: ' + tail + '' : ''}`
        ))
      })
    })
  }

  // === 冷启动预热 ===
  // 实测：dsh 首调 25~60s（Windows 侧 node/dsh 文件缓存与登录态全冷），热调仅约 5s；
  // 且 dsh 是先等 stdin EOF 才 boot，无法挂起预热进程，只能在后台真跑一发极小任务焐热缓存。
  // 服务启动 3s 后预热一次；之后每 DSH_KEEPALIVE_SEC 秒保温（0 = 只预热不保温）
  function warmup() {
    const cfg = effectiveDshConfig()
    const argv = cfg.cmd.split(/\s+/).filter(Boolean)
    argv.push('--profile', 'headless', '--json', '-')
    const cwd = existsSync(DSH_CWD) ? DSH_CWD : undefined
    let child
    try { child = spawn(argv[0], argv.slice(1), { cwd, windowsHide: true }) } catch { return }
    const t0 = Date.now()
    let buf = ''
    child.stdout.on('data', (c) => {
      buf += c
      // 到 final 即预热完成，尽快收尾（WSL interop 下 close 事件不可靠，主动杀树）
      if (buf.includes('"type":"final"')) {
        console.log(`[dsh] warmup done in ${((Date.now() - t0) / 1000).toFixed(1)}s`)
        killChildTree(child)
      }
    })
    child.on('error', () => {})
    child.stdin.on('error', () => {})  // EPIPE 不许炸服务
    child.stdin.end('预热，只回复ok')
  }
  const KEEPALIVE_SEC = Number(process.env.DSH_KEEPALIVE_SEC ?? 300)
  setTimeout(warmup, 3000)
  if (KEEPALIVE_SEC > 0) setInterval(warmup, KEEPALIVE_SEC * 1000)

  // === POST /api/dsh/chat（同步调 DSH CLI，headless 单发，reply= stdout）===
  ctx.post('/api/dsh/chat', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    let payload
    try { payload = JSON.parse(body) }
    catch { res.writeHead(400); res.end(JSON.stringify({ ok: false, error: 'invalid json' })); return }
    const message = String(payload.message || '').trim()
    if (!message) { res.writeHead(400); res.end(JSON.stringify({ ok: false, error: 'empty' })); return }
    // 会话规则：传入 session 则续会话；不传则新开会话，真实会话号由 dsh 发放（--json 事件流回传）。
    // dsh 的 --session-id 只认已存在会话，服务端不自造会话号；万一事件流缺失 session 事件
    // （理论上不该发生），回退生成一个占位号保证响应始终带 session 字段
    // 兼容 docUrl 字段名（插件沿用 plugin-ai 的 body 约定，曾导致 session 永远传不进来、DSH 每次失忆重探）
    const sessionIn = String(payload.session || payload.docUrl || '').trim()
    // 协议：DSH 也吃 actions 协议（不吃协议只会回纯文本，插件降级堆文末，分页/表格全丢）。
    // 新会话才带全量协议（~4.6KB）；续会话时协议已留在 dsh 历史里，只补一行提醒防走样
    const ctxDoc = String(payload.docContext || docContext.map.get(sessionIn) || '').slice(0, 4000)
    // 文件约束用提示而非禁令：用户明确要求读本地文件路径时放行（headless 本就带文件/业务工具）
    const full = (sessionIn
      ? '（提醒：继续遵守本会话首轮约定的 actions JSON 输出协议，只输出一个 JSON 对象）\n'
      : PROTOCOL_PROMPT + '\n')
      + '【任务约束】文档正文已随消息给出（见 [当前文档正文] 段），一般无需读写/搜索本地文件；仅当用户明确给出本地文件路径或要求查询业务数据时才使用工具。\n'
      + '【用户指令】' + message
      + (ctxDoc ? '\n\n[当前文档正文]\n' + ctxDoc : '')
    // 思考过程进度（面板轮询 /api/dsh/chat-progress?reqId= 读取）
    const reqId = String(payload.reqId || '')
    const progress = reqId ? { reqId, lines: [] } : null
    if (progress) dshProgress = progress
    try {
      const { reply, sessionId } = await callDSH(full, sessionIn, progress)
      const session = sessionId || sessionIn || 'dsh-' + crypto.randomBytes(4).toString('hex')
      saveDshSession(session) // 登记给插件轮询兜底（GET /api/dsh/session），重启不丢
      // 与 /api/chat 一致：按 actions 协议解析 DSH 回复，插件按动作精确写入（分页/表格/替换等）
      const parsed = tryParseActions(reply)
      if (parsed) {
        res.end(JSON.stringify({ ok: true, source: 'dsh', message: parsed.summary, actions: parsed.actions, session }))
      } else {
        res.end(JSON.stringify({ ok: true, source: 'dsh', message: reply, session }))
      }
    } catch (e) {
      const msg = String(e.message || e)
      res.writeHead(msg.includes('超时') ? 504 : 502)
      res.end(JSON.stringify({ ok: false, error: msg }))
    }
  })

  // === GET /api/dsh/chat-progress?reqId=（思考过程行，chat 进行中持续更新）===
  ctx.get('/api/dsh/chat-progress', async (req, res, url) => {
    const rid = new URL(url, 'http://x').searchParams.get('reqId') || ''
    cors(res, req)
    const lines = dshProgress && dshProgress.reqId === rid ? dshProgress.lines : []
    res.end(JSON.stringify({ ok: true, lines }))
  })

  // === POST /api/dsh/stop（终止进行中的 DSH 子进程树）===
  ctx.post('/api/dsh/stop', async (req, res) => {
    cors(res, req)
    if (dshChild) { dshStopped = true; killChildTree(dshChild) }
    res.end(JSON.stringify({ ok: true }))
  })

  // === POST /api/dsh/queue（往 DSH 独立队列塞指令，机制同 /api/queue）===
  ctx.post('/api/dsh/queue', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const { session, instruction } = JSON.parse(body)
      if (!session || !instruction) { res.writeHead(400); res.end(JSON.stringify({ error: 'missing' })); return }
      dshQueuePut(session, instruction)
      res.end(JSON.stringify({ ok: true }))
    } catch { res.writeHead(400); res.end(JSON.stringify({ error: 'bad json' })) }
  })

  // === GET /api/dsh/pending?session=（DSH 插件轮询，机制同 /api/pending）===
  ctx.get('/api/dsh/pending', async (req, res, url) => {
    const u = new URL(url, 'http://x')
    const session = u.searchParams.get('session') || ''
    cors(res, req)
    const item = dshQueueTake(session)
    res.end(JSON.stringify({ instruction: item ? item.instruction : null }))
  })

  // === POST /api/dsh/session（登记 DSH 当前会话，持久化防重启丢失）===
  ctx.post('/api/dsh/session', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const s = String(JSON.parse(body).session || '').trim()
      if (s) saveDshSession(s)
    } catch {}
    res.end(JSON.stringify({ ok: true, session: dshActiveSession }))
  })

  // === GET /api/dsh/session（DSH 插件兜底查询）===
  ctx.get('/api/dsh/session', async (req, res) => {
    cors(res, req)
    res.end(JSON.stringify({ session: dshActiveSession }))
  })

  // === GET /api/dsh/settings（回面板已存值 + 实际生效值）===
  ctx.get('/api/dsh/settings', async (req, res) => {
    cors(res, req)
    const cfg = effectiveDshConfig()
    res.end(JSON.stringify({
      dshCmd: dshSettings.dshCmd || '',           // 空 = 用 .env / 默认 dsh
      dshTimeout: dshSettings.dshTimeout || '',    // 空 = 默认 600s
      cmdResolved: cfg.cmd,
      timeoutResolved: Math.round(cfg.timeoutMs / 1000),
    }))
  })

  // === POST /api/dsh/settings（保存即生效；传空恢复 env/默认）===
  ctx.post('/api/dsh/settings', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const p = JSON.parse(body)
      for (const k of ['dshCmd', 'dshTimeout']) {
        if (k in p) {
          const v = typeof p[k] === 'number' ? String(p[k]) : (typeof p[k] === 'string' ? p[k].trim() : '')
          if (v) dshSettings[k] = v
          else delete dshSettings[k]
        }
      }
      saveDshSettings()
      res.end(JSON.stringify({ ok: true }))
    } catch { res.writeHead(400); res.end(JSON.stringify({ error: 'bad json' })) }
  })

  // === POST /api/dsh/log（dsh 插件远程日志，打到服务端 stdout）===
  ctx.post('/api/dsh/log', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const j = JSON.parse(body)
      pushLog(j.source || 'dsh-plugin', String(j.msg || '').slice(0, 500))
    } catch {}
    res.end(JSON.stringify({ ok: true }))
  })

  // === POST /api/dsh/context（dsh 插件回传文档正文，供 chat 选址参考）===
  ctx.post('/api/dsh/context', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const { session, text } = JSON.parse(body)
      const t = String(text || '').slice(0, 8000)
      if (session) docContext.map.set(session, t)
      docContext.last = t
    } catch {}
    res.end(JSON.stringify({ ok: true }))
  })

  // === POST /api/dsh/like（DSH 点赞记忆，独立标记不复用进 plugin-ai 检索）===
  ctx.post('/api/dsh/like', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const p = JSON.parse(body)
      if (!p.message || !p.summary) { res.writeHead(400); res.end(JSON.stringify({ error: 'message/summary required' })); return }
      const dup = ctx.memory.items.find((it) => it.message === p.message && it.summary === p.summary)
      if (!dup) {
        ctx.memory.items.push({ message: String(p.message), summary: String(p.summary), actions: p.actions || null, source: 'dsh', ts: Date.now() })
        ctx.memory.save()
      }
      res.end(JSON.stringify({ ok: true, total: ctx.memory.items.length }))
    } catch { res.writeHead(400); res.end(JSON.stringify({ error: 'bad json' })) }
  })
}
