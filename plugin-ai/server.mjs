// plugin-ai（AI Assistant 面板）后端：直连 OpenAI 兼容 API，回复解析成结构化动作写入文档
// 路由：POST /api/chat、GET /api/chat-progress?reqId=、GET|POST /api/settings、
//       POST /api/like、POST /api/queue、GET /api/pending?session=、POST /api/log、
//       POST /api/context、POST|GET /api/session、GET /api/health
import { readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'

const AI_API_URL = (process.env.AI_API_URL || '').trim()
// 剥掉用户粘贴时可能带上的 /v1 或 /v1beta1 /v2 /openai 这种"版本尾巴"，
// 让拼路径时只拼一处 /v1/chat/completions，避免出现 /v1/v1/chat/completions 双 /v1。
function stripVersionTail(url) {
  return url
    .replace(/\/(v\d+(?:\.\d+)*(?:beta(?:\d+)?)?|openai|api|openapi)\/?$/i, '')
    .replace(/\/$/, '')
}
const AI_API_BASE = stripVersionTail(AI_API_URL)
const AI_API_KEY = process.env.AI_API_KEY || ''
const AI_MODEL = process.env.AI_MODEL || 'gpt-4o-mini'

function maskKey(k) {
  return k ? (k.length > 8 ? k.slice(0, 3) + '***' + k.slice(-4) : '***') : ''
}

export function register(ctx) {
  const PROTOCOL_PROMPT = ctx.prompt
  const { cors, readBody, pushLog, tryParseActions, memory, docContext } = ctx

  // === 运行时设置：插件设置面板保存的覆盖值（.settings.json 不进 git） ===
  // 兼容旧部署：文件仍落 server/ 目录旁（ctx.serverDir），数据目录注入时优先
  const SETTINGS_FILE = ctx.dataDir ? path.join(ctx.dataDir, '.settings.json') : path.join(ctx.serverDir, '.settings.json')
  let runtimeSettings = {}
  try { runtimeSettings = JSON.parse(readFileSync(SETTINGS_FILE, 'utf8')) } catch {}
  function saveSettings() {
    try { writeFileSync(SETTINGS_FILE, JSON.stringify(runtimeSettings, null, 2)) } catch {}
  }
  function effectiveConfig() {
    return {
      apiUrl: stripVersionTail(runtimeSettings.apiUrl || AI_API_URL),
      apiKey: runtimeSettings.apiKey || AI_API_KEY,
      model: runtimeSettings.model || AI_MODEL,
      extraPrompt: runtimeSettings.extraPrompt || '',
    }
  }

  // 系统提示词独立在 prompt.mjs（新增动作先改那边的协议，再到 plugin.js 加写入函数）
  function buildSystemPrompt(docText) {
    const cfg = effectiveConfig()
    let s = PROTOCOL_PROMPT
    if (cfg.extraPrompt) s += '\n【用户附加要求（需遵守）】\n' + cfg.extraPrompt
    if (!docText) return s
    const trimmed = String(docText).slice(0, 4000)
    return s + '\n【文档当前内容（供选址/anchor 参考）】\n' + trimmed
  }

  // 思考过程进度（reqId → {reqId, lines[]}），GET /api/chat-progress 轮询读取；只留最近 20 个
  const aiProgress = new Map()

  async function callLLM(message, ctxDoc, progress) {
    const cfg = effectiveConfig()
    const headers = { 'Content-Type': 'application/json' }
    if (cfg.apiKey) headers['Authorization'] = 'Bearer ' + cfg.apiKey
    const payload = { model: cfg.model, messages: [
        { role: 'system', content: buildSystemPrompt(ctxDoc) },
        { role: 'user', content: message },
      ],
    }
    // 无 progress：老路径一次性拿回复（30s 超时兜底网关半死）
    if (!progress) {
      const r = await fetch(cfg.apiUrl + '/v1/chat/completions', {
        method: 'POST', headers,
        signal: AbortSignal.timeout(30000),
        body: JSON.stringify(payload),
      })
      const j = await r.json()
      return j.choices?.[0]?.message?.content || j.choices?.[0]?.text || j.reply || ''
    }
    // 有 progress：改流式，边收 reasoning_content 边喂思考行（同 DSH 面板思考块）
    // ponytail: 整流 120s 上限（推理模型思考可超 30s），网关挂死仍会被杀
    payload.stream = true
    const r = await fetch(cfg.apiUrl + '/v1/chat/completions', {
      method: 'POST', headers,
      signal: AbortSignal.timeout(120000),
      body: JSON.stringify(payload),
    })
    if (!r.ok || !r.body) {
      const t = await r.text().catch(() => '')
      throw new Error(`LLM ${r.status}${t ? ': ' + t.slice(0, 120) : ''}`)
    }
    const pushLine = (t) => {
      const s = t.replace(/\s+/g, ' ').trim()
      if (s) progress.lines.push('💭 ' + s.slice(0, 140))
      if (progress.lines.length > 200) progress.lines.splice(0, progress.lines.length - 200)
    }
    let answer = '', pend = ''
    const dec = new TextDecoder()
    for await (const chunk of r.body) {
      for (const line of dec.decode(chunk, { stream: true }).split(/\r?\n/)) {
        const s = line.trim()
        if (!s.startsWith('data:')) continue
        const d = s.slice(5).trim()
        if (!d || d === '[DONE]') continue
        try {
          const delta = JSON.parse(d).choices?.[0]?.delta
          if (!delta) continue
          if (typeof delta.reasoning_content === 'string' && delta.reasoning_content) {
            pend += delta.reasoning_content
            let i
            while ((i = pend.indexOf('\n')) >= 0) { pushLine(pend.slice(0, i)); pend = pend.slice(i + 1) }
            if (pend.length >= 80) pushLine(pend), (pend = '')  // 无换行的长思考也持续出行
          }
          if (typeof delta.content === 'string') answer += delta.content
        } catch {}
      }
    }
    if (pend.trim()) pushLine(pend)
    return answer
  }

  async function stubReply(message) {
    return {
      source: 'stub',
      message: `（stub）已收到：${message}。未配置可用的 AI API，请检查 .env 或面板设置。`,
      actions: [],
    }
  }

  async function dispatch(message, ctxDoc, progress) {
    try {
      const reply = (await callLLM(message, ctxDoc, progress)).trim()
      const parsed = tryParseActions(reply)
      if (parsed) {
        return { source: 'ai', message: parsed.summary, reply: parsed.summary, actions: parsed.actions }
      }
      // 模型没按 JSON 输出：整体作为内容插到文末（兜底仍能写入）
      return { source: 'ai', message: reply.slice(0, 120), reply, actions: [{ type: 'insert', target: 'end', content: reply }] }
    } catch (e) {
      console.warn('[assistant] LLM failed, stub fallback:', e.message)
      return await stubReply(message)
    }
  }

  // === 内存队列 ===
  const pendingQueue = new Map()
  function queuePut(session, instruction) {
    if (!pendingQueue.has(session)) pendingQueue.set(session, [])
    pendingQueue.get(session).push({ instruction, ts: Date.now() })
  }
  function queueTake(session) {
    const q = pendingQueue.get(session)
    if (!q || q.length === 0) return null
    return q.shift()
  }

  // 当前活跃会话：插件 iframe 可能丢失 query 参数，用这个兜底
  // 持久化到磁盘，重启后不丢（否则插件要重载页面才能重新拿到 session）
  const SESSION_FILE = ctx.dataDir ? path.join(ctx.dataDir, 'active-session.txt') : '/tmp/oo-active-session'
  let activeSession = ''
  try { activeSession = readFileSync(SESSION_FILE, 'utf8').trim() } catch {}
  function saveActiveSession(s) {
    activeSession = s
    try { writeFileSync(SESSION_FILE, s) } catch {}
  }

  // === POST /api/chat ===
  ctx.post('/api/chat', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    let payload
    try { payload = JSON.parse(body) }
    catch { res.writeHead(400); res.end(JSON.stringify({ error: 'invalid json' })); return }
    const message = (payload.message || '').trim()
    const session = (payload.docUrl || '').trim() || (payload.session || '').trim()
    if (!message) { res.writeHead(400); res.end(JSON.stringify({ error: 'empty' })); return }
    // 点赞记忆命中：相似提问直接复用赞过的回复与写入动作
    const hit = memory.findHit(message)
    if (hit && hit.score >= 0.7 && hit.item.actions) {
      res.end(JSON.stringify({
        source: 'memory',
        message: hit.item.summary,
        actions: hit.item.actions,
        memoryHit: true,
        score: Number(hit.score.toFixed(2)),
      }))
      return
    }
    // 文档上下文：优先用请求自带（插件直聊实时抽取），否则用插件回传的缓存
    let ctxDoc = String(payload.docContext || docContext.map.get(session) || docContext.last || '')
    // 思考过程进度：带 reqId 时走流式（GET /api/chat-progress 轮询读取）
    const reqId = String(payload.reqId || '')
    const progress = reqId ? { reqId, lines: [] } : null
    if (progress) {
      aiProgress.set(reqId, progress)
      if (aiProgress.size > 20) aiProgress.delete(aiProgress.keys().next().value)
    }
    // 次相似命中：把赞过的问答作为风格参考注入
    if (hit && hit.score >= 0.35) {
      ctxDoc += `\n【参考：用户曾点赞的类似提问】问：${hit.item.message} 你当时的回复（风格与选址偏好供参考）：${hit.item.summary}`
    }
    try {
      const r = await dispatch(message, ctxDoc, progress)
      if (hit && hit.score >= 0.35) r.memoryRef = true
      res.end(JSON.stringify(r))
    } catch (e) {
      res.writeHead(500); res.end(JSON.stringify({ source: 'error', message: String(e.message || e) }))
    }
  })

  // === GET /api/chat-progress?reqId=（AI 思考过程行，chat 进行中持续更新）===
  ctx.get('/api/chat-progress', async (req, res, url) => {
    const rid = new URL(url, 'http://x').searchParams.get('reqId') || ''
    cors(res, req)
    const p = aiProgress.get(rid)
    res.end(JSON.stringify({ ok: true, lines: p ? p.lines : [] }))
  })

  // === GET /api/settings（Key 只回掩码）===
  ctx.get('/api/settings', async (req, res) => {
    cors(res, req)
    const cfg = effectiveConfig()
    res.end(JSON.stringify({
      apiUrl: runtimeSettings.apiUrl || '',           // 空 = 用 .env 默认
      apiKeyMasked: maskKey(cfg.apiKey),
      apiKeySet: !!cfg.apiKey,
      model: runtimeSettings.model || '',            // 空 = 用 .env 默认
      extraPrompt: runtimeSettings.extraPrompt || '',
    }))
  })

  // === POST /api/settings（保存覆盖值；传 null/'' 恢复 .env 默认）===
  ctx.post('/api/settings', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const p = JSON.parse(body)
      for (const k of ['apiUrl', 'apiKey', 'model', 'extraPrompt']) {
        if (k in p) {
          const v = typeof p[k] === 'string' ? p[k].trim() : ''
          if (v) runtimeSettings[k] = v
          else delete runtimeSettings[k]
        }
      }
      saveSettings()
      const cfg = effectiveConfig()
      res.end(JSON.stringify({ ok: true, apiKeyMasked: maskKey(cfg.apiKey), apiKeySet: !!cfg.apiKey }))
    } catch { res.writeHead(400); res.end(JSON.stringify({ error: 'bad json' })) }
  })

  // === POST /api/like（点赞记忆）===
  ctx.post('/api/like', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const p = JSON.parse(body)
      if (!p.message || !p.summary) { res.writeHead(400); res.end(JSON.stringify({ error: 'message/summary required' })); return }
      const dup = memory.items.find((it) => it.message === p.message && it.summary === p.summary)
      if (!dup) {
        memory.items.push({ message: String(p.message), summary: String(p.summary), actions: p.actions || null, ts: Date.now() })
        memory.save()
      }
      res.end(JSON.stringify({ ok: true, total: memory.items.length }))
    } catch { res.writeHead(400); res.end(JSON.stringify({ error: 'bad json' })) }
  })

  // === POST /api/queue ===
  ctx.post('/api/queue', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const { session, instruction } = JSON.parse(body)
      if (!session || !instruction) { res.writeHead(400); res.end(JSON.stringify({ error: 'missing' })); return }
      queuePut(session, instruction)
      res.end(JSON.stringify({ ok: true }))
    } catch { res.writeHead(400); res.end(JSON.stringify({ error: 'bad json' })) }
  })

  // === GET /api/pending ===
  ctx.get('/api/pending', async (req, res, url) => {
    const u = new URL(url, 'http://x')
    const session = u.searchParams.get('session') || ''
    cors(res, req)
    const item = queueTake(session)
    res.end(JSON.stringify({ instruction: item ? item.instruction : null }))
  })

  // === POST /api/log（插件远程日志，直接打到服务端 stdout）===
  ctx.post('/api/log', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try {
      const j = JSON.parse(body)
      pushLog(j.source || 'plugin', String(j.msg || '').slice(0, 500))
    } catch {}
    res.end(JSON.stringify({ ok: true }))
  })

  // === POST /api/context（插件回传文档正文，供选址参考）===
  ctx.post('/api/context', async (req, res) => {
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

  // === POST /api/session（登记当前文档会话）===
  ctx.post('/api/session', async (req, res) => {
    const body = await readBody(req)
    cors(res, req)
    try { saveActiveSession((JSON.parse(body).session || '').trim()) } catch {}
    res.end(JSON.stringify({ ok: true, session: activeSession }))
  })

  // === GET /api/session（插件兜底查询）===
  ctx.get('/api/session', async (req, res) => {
    cors(res, req)
    res.end(JSON.stringify({ session: activeSession }))
  })

  // === GET /api/health ===
  ctx.get('/api/health', async (req, res) => {
    cors(res, req)
    res.end(JSON.stringify({ ok: true, apiUrl: AI_API_BASE }))
  })
}
