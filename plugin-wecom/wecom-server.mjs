// 企微登录后端（归属 plugin-wecom，为聊天/协同提供真实 userid+姓名身份）
// 可选功能，非门禁：WECOM_CORP_ID / WECOM_AGENT_ID / WECOM_SECRET
// 三项齐全才启用；WECOM_MOCK=1 为无凭据本地/e2e 验证开关（不出站，authorize 直接弹回假 code）。
// 授权走独立弹窗：插件 iframe window.open authorize → 企微扫码/静默授权 → 弹回
// /api/auth/wecom/callback-page（本服务直接吐 HTML）→ 弹窗内 POST login 落 cookie → 自动关闭，
// 插件面板轮询 /api/auth/me 感知登录成功。全程不依赖主页面的任何回调处理。
// 路由：GET /api/auth/wecom/config、GET /api/auth/wecom/authorize（302 分流）、
//       GET /api/auth/wecom/callback-page、POST /api/auth/wecom/login、
//       GET /api/auth/me、POST /api/auth/logout
import crypto from 'node:crypto'

const CORP_ID = process.env.WECOM_CORP_ID || ''
const AGENT_ID = process.env.WECOM_AGENT_ID || ''
const SECRET = process.env.WECOM_SECRET || ''
const PUBLIC_ORIGIN = (process.env.WECOM_PUBLIC_ORIGIN || '').replace(/\/+$/, '')
const MOCK = process.env.WECOM_MOCK === '1'
const ENABLED = !!(CORP_ID && AGENT_ID && SECRET) || MOCK
const SESSION_TTL = Number(process.env.WECOM_SESSION_TTL || 7 * 24 * 3600) * 1000
const STATE_TTL = 10 * 60 * 1000 // state 防 CSRF：一次性、10 分钟窗口

const pendingStates = new Map() // state -> 签发时间戳
const sessions = new Map() // sid -> { userid, name, ts }（滑动过期）
let tok = { v: '', exp: 0 } // access_token 缓存（提前 300s 过期）
let tokInflight = null

function parseCookies(req) {
  const out = {}
  const raw = req.headers.cookie || ''
  for (const part of raw.split(';')) {
    const i = part.indexOf('=')
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim())
  }
  return out
}

function getOrigin(req) {
  if (PUBLIC_ORIGIN) return PUBLIC_ORIGIN
  const host = req.headers.host || 'localhost'
  return `http://${host}`
}

function isWxwork(req) {
  return /wxwork/i.test(req.headers['user-agent'] || '')
}

// 统一 qyapi GET：超时 10s + JSON 解析 + errcode 检查
async function qyapiGet(pathname, params) {
  const qs = new URLSearchParams(params).toString()
  const r = await fetch(`https://qyapi.weixin.qq.com${pathname}?${qs}`, { signal: AbortSignal.timeout(10000) })
  const j = await r.json()
  if (j.errcode) throw new Error(`企微接口 ${pathname} 失败: errcode=${j.errcode} errmsg=${j.errmsg || ''}`)
  return j
}

async function getAccessToken(force) {
  if (!force && tok.v && Date.now() < tok.exp) return tok.v
  if (tokInflight) return tokInflight
  tokInflight = (async () => {
    const j = await qyapiGet('/cgi-bin/gettoken', { corpid: CORP_ID, corpsecret: SECRET })
    tok = { v: j.access_token, exp: Date.now() + (j.expires_in - 300) * 1000 }
    return tok.v
  })()
  try { return await tokInflight } finally { tokInflight = null }
}

// 授权弹窗回跳页：拿 URL 里的 code/state 换登录态（cookie 落在本服务 origin），然后自动关窗
const CALLBACK_HTML = `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><title>企微登录</title>
<style>body{font-family:'Segoe UI',Arial,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;color:#333}
#tip{font-size:14px;color:#666}</style></head>
<body><div id="tip">正在登录…</div>
<script>
(async function () {
  var tip = document.getElementById('tip')
  try {
    var q = new URLSearchParams(location.search)
    var r = await fetch('/api/auth/wecom/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ code: q.get('code'), state: q.get('state') })
    })
    var j = await r.json()
    if (j.ok) {
      tip.textContent = '登录成功：' + j.name + '（窗口即将关闭）'
      setTimeout(function () { window.close() }, 400)
      // 关窗失败（非脚本打开）时给出手动关闭提示
      setTimeout(function () { tip.textContent = '登录成功：' + j.name + '，请关闭此窗口回到编辑器' }, 2000)
    } else {
      tip.textContent = '登录失败：' + (j.error || ('HTTP ' + r.status)) + '，请关闭后重试'
    }
  } catch (e) {
    tip.textContent = '登录失败：' + String(e.message || e) + '，请关闭后重试'
  }
})()
</script></body></html>`

// 供其他插件复用：按请求 cookie 取当前登录身份，未登录/过期返回 null
export function getSessionUser(req) {
  const sid = parseCookies(req).oo_sid
  const s = sid && sessions.get(sid)
  if (!s || Date.now() - s.ts > SESSION_TTL) {
    if (s) sessions.delete(sid)
    return null
  }
  s.ts = Date.now()
  return { userid: s.userid, name: s.name }
}

export function register(ctx) {
  const { cors, readBody } = ctx

  // 功能开关：插件面板据此决定显示登录还是未配置提示
  ctx.get('/api/auth/wecom/config', (req, res) => {
    cors(res, req)
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' })
    res.end(JSON.stringify({ enabled: ENABLED, corpId: CORP_ID, agentId: AGENT_ID, mock: MOCK }))
  })

  // 302 跳企业微信授权页：UA 含 wxwork 走企微内静默授权，否则走 PC 扫码登录页。
  // redirect_uri 固定指向本服务的弹窗回跳页（callback-page），登录闭环全程不经过主页面
  ctx.get('/api/auth/wecom/authorize', (req, res) => {
    if (!ENABLED) {
      cors(res, req)
      res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' })
      res.end(JSON.stringify({ error: '企业微信登录未配置' }))
      return
    }
    const redirectUri = `${getOrigin(req)}/api/auth/wecom/callback-page`
    const state = crypto.randomBytes(12).toString('hex')
    pendingStates.set(state, Date.now())
    // 顺手清掉过期 state（惰性清理）
    for (const [s, ts] of pendingStates) if (Date.now() - ts > STATE_TTL) pendingStates.delete(s)

    if (MOCK) {
      res.writeHead(302, { Location: `${redirectUri}?code=mock&state=${state}` })
      res.end()
      return
    }
    const authUrl = isWxwork(req)
      ? `https://open.weixin.qq.com/connect/oauth2/authorize?appid=${CORP_ID}&redirect_uri=${encodeURIComponent(redirectUri)}&response_type=code&scope=snsapi_base&state=${state}&agentid=${AGENT_ID}#wechat_redirect`
      : `https://login.work.weixin.qq.com/wwlogin/sso/login?login_type=CorpApp&appid=${CORP_ID}&agentid=${AGENT_ID}&redirect_uri=${encodeURIComponent(redirectUri)}&state=${state}`
    res.writeHead(302, { Location: authUrl })
    res.end()
  })

  // 授权弹窗回跳页（静态 HTML）
  ctx.get('/api/auth/wecom/callback-page', (req, res) => {
    cors(res, req)
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' })
    res.end(CALLBACK_HTML)
  })

  // code 换身份：state 一次性校验 → getuserinfo 换 userid → user/get 拉姓名（失败兜底 userid）
  ctx.post('/api/auth/wecom/login', async (req, res) => {
    cors(res, req)
    const end = (code, obj) => {
      res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' })
      res.end(JSON.stringify(obj))
    }
    let body = {}
    try { body = JSON.parse(await readBody(req)) || {} } catch { /* 空体 */ }
    if (!ENABLED) return end(404, { error: '企业微信登录未配置' })
    const st = pendingStates.get(body.state)
    if (!st || Date.now() - st > STATE_TTL) return end(401, { error: '登录会话过期，请重新点击登录' })
    pendingStates.delete(body.state) // 一次性，立即失效

    let userid = ''
    let uname = ''
    if (MOCK) {
      userid = 'mock_' + crypto.randomBytes(2).toString('hex')
      uname = body.name || '模拟用户'
    } else {
      try {
        let token = await getAccessToken()
        let j
        try {
          j = await qyapiGet('/cgi-bin/auth/getuserinfo', { access_token: token, code: body.code })
        } catch (e) {
          // token 失效/过期：强制刷新重试一次
          if (!/errcode=(40014|42001)/.test(String(e.message))) throw e
          token = await getAccessToken(true)
          j = await qyapiGet('/cgi-bin/auth/getuserinfo', { access_token: token, code: body.code })
        }
        userid = j.userid || j.UserId || '' // 新旧字段名兼容
        if (!userid) return end(401, { error: '非企业成员，无法登录' })
        try {
          const u = await qyapiGet('/cgi-bin/user/get', { access_token: token, userid })
          uname = u.name || ''
        } catch { /* 无通讯录权限等：拉姓名失败不阻断登录 */ }
        uname = uname || userid
      } catch (e) {
        return end(502, { error: String(e.message || e) })
      }
    }
    const sid = crypto.randomBytes(24).toString('hex')
    sessions.set(sid, { userid, name: uname, ts: Date.now() })
    res.setHeader('Set-Cookie', `oo_sid=${sid}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${Math.floor(SESSION_TTL / 1000)}`)
    end(200, { ok: true, userid, name: uname })
  })

  // 当前登录态（滑动续期 + 惰性清理过期 session）
  ctx.get('/api/auth/me', (req, res) => {
    cors(res, req)
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' })
    const sid = parseCookies(req).oo_sid
    const s = sid && sessions.get(sid)
    if (!s || Date.now() - s.ts > SESSION_TTL) {
      if (s) sessions.delete(sid)
      res.end(JSON.stringify({ authed: false }))
      return
    }
    s.ts = Date.now()
    for (const [k, v] of sessions) if (Date.now() - v.ts > SESSION_TTL) sessions.delete(k)
    res.end(JSON.stringify({ authed: true, userid: s.userid, name: s.name }))
  })

  ctx.post('/api/auth/logout', (req, res) => {
    cors(res, req)
    const sid = parseCookies(req).oo_sid
    if (sid) sessions.delete(sid)
    res.setHeader('Set-Cookie', 'oo_sid=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0')
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' })
    res.end(JSON.stringify({ ok: true }))
  })
}
