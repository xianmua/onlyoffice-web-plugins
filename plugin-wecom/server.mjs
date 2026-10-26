// 聊天插件后端：服务器作为聊天中枢（不走 DS 协同通道，不受社区版 License 限制）。
// 消息按文档 id(f) 追加存储：dataDir/chat/<f>.json（dev 兜底 /tmp/oo-chat），
// 每文档内存缓存最近 500 条，写入即落盘。身份复用同目录 wecom-server.mjs 的登录 cookie
//（getSessionUser），未登录一律 401 只读。前端 1s 轮询增量拉取（after=seq）。
// 路由：GET /api/chat/messages?f=&after=（轮询顺带返回登录态，前端免单独探测）、
//       POST /api/chat/send?f=
import fs from 'node:fs'
import path from 'node:path'
import { getSessionUser } from './wecom-server.mjs'

const MAX_MSGS = 500 // 每文档保留上限（超出丢最旧）
const MAX_TEXT = 2000 // 单条消息长度上限

const chats = new Map() // f -> { msgs: [...], seq }
let dir = ''

function safeF(f) {
  return /^[A-Za-z0-9_-]{1,64}$/.test(f || '') ? f : ''
}

function loadDoc(ctx, f) {
  if (chats.has(f)) return chats.get(f)
  let doc = { msgs: [], seq: 0 }
  try {
    const file = path.join(ctx.dataDir ? path.join(ctx.dataDir, 'chat') : '/tmp/oo-chat', f + '.json')
    const arr = JSON.parse(fs.readFileSync(file, 'utf8'))
    if (Array.isArray(arr)) {
      doc.msgs = arr.slice(-MAX_MSGS)
      doc.seq = doc.msgs.length ? doc.msgs[doc.msgs.length - 1].seq : 0
    }
  } catch { /* 首次/损坏：从空开始 */ }
  chats.set(f, doc)
  return doc
}

function persist(ctx, f, doc) {
  try {
    const d = ctx.dataDir ? path.join(ctx.dataDir, 'chat') : '/tmp/oo-chat'
    if (d !== dir) { dir = d; fs.mkdirSync(d, { recursive: true }) }
    fs.writeFileSync(path.join(d, f + '.json'), JSON.stringify(doc.msgs))
  } catch { /* 落盘失败不阻断内存会话 */ }
}

export function register(ctx) {
  const { cors, readBody } = ctx
  const json = (res, code, obj) => {
    res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' })
    res.end(JSON.stringify(obj))
  }

  // 轮询拉取：after=已见最大 seq，返回增量；authed/name 随响应返回（驱动面板三态）
  ctx.get('/api/chat/messages', (req, res) => {
    cors(res, req)
    const u = new URL(req.url, 'http://x')
    const f = safeF(u.searchParams.get('f'))
    if (!f) return json(res, 400, { error: '缺少文档 id' })
    const after = Number(u.searchParams.get('after') || 0) || 0
    const doc = loadDoc(ctx, f)
    const user = getSessionUser(req)
    const from = Math.max(0, Math.min(after, doc.seq))
    json(res, 200, {
      ok: true,
      authed: !!user,
      name: user ? user.name : '',
      seq: doc.seq,
      messages: doc.msgs.filter(m => m.seq > from),
    })
  })

  ctx.post('/api/chat/send', async (req, res) => {
    cors(res, req)
    const u = new URL(req.url, 'http://x')
    const f = safeF(u.searchParams.get('f'))
    if (!f) return json(res, 400, { error: '缺少文档 id' })
    const user = getSessionUser(req)
    if (!user) return json(res, 401, { error: 'not_authed' })
    let body = {}
    try { body = JSON.parse(await readBody(req)) || {} } catch { /* 空体 */ }
    const text = String(body.text || '').trim().slice(0, MAX_TEXT)
    if (!text) return json(res, 400, { error: '消息为空' })
    const doc = loadDoc(ctx, f)
    const msg = { seq: ++doc.seq, userid: user.userid, name: user.name, text, ts: Date.now() }
    doc.msgs.push(msg)
    if (doc.msgs.length > MAX_MSGS) doc.msgs.splice(0, doc.msgs.length - MAX_MSGS)
    persist(ctx, f, doc)
    json(res, 200, { ok: true, message: msg })
  })
}
