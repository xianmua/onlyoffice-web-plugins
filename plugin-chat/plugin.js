// OnlyOffice 聊天插件（服务器中枢模式，兼企微登录身份入口）
// 面板 1s 轮询 GET /api/chat/messages?f=&after=seq 增量拉取；响应顺带返回登录态，
// 驱动三态：已登录（可发，右上角显示姓名+退出）→ 未登录/过期（只读 + 登录条）→ 未配置企微（只读 + 提示）。
// 未登录可看消息（只读）；登录/退出均在本面板完成（共享 auth.js 弹窗授权，cookie 同源共享，
// 主页面 F5 后以真实姓名协同）。
// 文档 id：优先 iframe query 的 session，兜底 GET /api/session（App.vue 登记的当前 f）。
(function () {
  const apiBase = window.location.origin
  const $ = function (id) { return document.getElementById(id) }
  const msgsEl = $('msgs'), emptyEl = $('empty'), whoEl = $('whoName'), logoutBtn = $('logoutBtn')
  const authbarEl = $('authbar'), authTipEl = $('authTip'), loginBtn = $('loginBtn')
  const inputEl = $('input'), sendBtn = $('send')

  let f = '' // 文档 id
  let seq = 0 // 已见最大消息序号（增量拉取）
  let authed = false
  let loginEnabled = false // 后端是否配置了企微登录
  let sending = false

  async function getDocId() {
    const qs = new URLSearchParams(location.search)
    let id = (qs.get('session') || qs.get('f') || '').trim()
    if (!id) {
      try { id = ((await (await fetch(apiBase + '/api/session', { cache: 'no-store' })).json()).session) || '' } catch { /* 兜底不可用 */ }
    }
    return id
  }

  function fmtTime(ts) {
    const d = new Date(ts)
    const p = function (n) { return (n < 10 ? '0' : '') + n }
    return p(d.getHours()) + ':' + p(d.getMinutes())
  }

  function appendMsg(m) {
    if (emptyEl.parentNode) emptyEl.remove()
    const me = authed && whoEl.dataset.userid && whoEl.dataset.userid === m.userid
    const div = document.createElement('div')
    div.className = 'msg' + (me ? ' self' : '')
    const meta = document.createElement('div')
    meta.className = 'meta'
    const b = document.createElement('b')
    b.textContent = m.name
    meta.appendChild(b)
    meta.appendChild(document.createTextNode(' ' + fmtTime(m.ts)))
    const bubble = document.createElement('div')
    bubble.className = 'bubble'
    bubble.textContent = m.text
    div.appendChild(meta); div.appendChild(bubble)
    const nearBottom = msgsEl.scrollTop + msgsEl.clientHeight >= msgsEl.scrollHeight - 24
    msgsEl.appendChild(div)
    if (nearBottom || me) msgsEl.scrollTop = msgsEl.scrollHeight
  }

  // 三态渲染：authed / loginEnabled 决定登录条、输入区可用性
  function renderState(expired) {
    whoEl.textContent = authed ? whoEl.dataset.name || '' : ''
    whoEl.title = authed && whoEl.dataset.userid ? 'userid: ' + whoEl.dataset.userid : ''
    logoutBtn.style.display = authed ? '' : 'none'
    if (authed) {
      authbarEl.classList.remove('show')
      inputEl.disabled = false; sendBtn.disabled = false
    } else {
      inputEl.disabled = true; sendBtn.disabled = true
      authTipEl.textContent = expired ? '登录已过期' : '登录后可参与讨论'
      loginBtn.style.display = loginEnabled ? '' : 'none'
      if (!loginEnabled) authTipEl.textContent = '未配置企业微信登录，仅可查看消息'
      authbarEl.classList.add('show')
    }
  }

  async function poll() {
    try {
      const r = await (await fetch(apiBase + '/api/chat/messages?f=' + encodeURIComponent(f) + '&after=' + seq, { cache: 'no-store' })).json()
      if (!r.ok) return
      const prev = authed
      authed = !!r.authed
      if (r.authed) { whoEl.dataset.userid = whoEl.dataset.userid || ''; whoEl.dataset.name = r.name || '' }
      else if (prev) { whoEl.dataset.name = '' } // 过期/登出：清名字，保留 userid 无意义
      if (authed && !whoEl.dataset.userid) {
        // 首次确认登录态时补 userid（messages 接口不回 userid，发消息后自动补）
      }
      if (prev !== authed) renderState(!authed && prev)
      if (r.messages && r.messages.length) {
        for (const m of r.messages) appendMsg(m)
        seq = r.seq
      }
    } catch { /* 后端不可达：下轮重试 */ }
  }

  async function send() {
    const text = inputEl.value.trim()
    if (!text || !authed || sending) return
    sending = true; sendBtn.disabled = true
    try {
      const r = await (await fetch(apiBase + '/api/chat/send?f=' + encodeURIComponent(f), {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ text }),
      })).json()
      if (r.ok) {
        inputEl.value = ''
        whoEl.dataset.userid = r.message.userid // 发消息即确认自己的 userid
        appendMsg(r.message)
        seq = Math.max(seq, r.message.seq)
      } else if (r.error === 'not_authed') {
        authed = false; renderState(true) // 会话过期：立刻降级
      }
    } catch { /* 网络抖动：输入保留，可重发 */ }
    sending = false; sendBtn.disabled = false
  }

  // 登录按钮：复用共享 auth.js 弹窗授权；成功后下一轮轮询自动解锁
  loginBtn.addEventListener('click', async function () {
    loginBtn.disabled = true
    authTipEl.textContent = '已打开授权窗口，请在新窗口完成登录…'
    await window.WecomAuth.startLogin()
    loginBtn.disabled = false
    const me = await window.WecomAuth.me()
    if (me.authed) {
      whoEl.dataset.name = me.name
      whoEl.dataset.userid = me.userid
      authed = true; renderState(false)
    } else authTipEl.textContent = '登录未完成，可重新点击'
  })

  // 退出登录：清 cookie，下一轮轮询降级只读；主页面 F5 后回到访客协同
  logoutBtn.addEventListener('click', async function () {
    logoutBtn.disabled = true
    await window.WecomAuth.logout()
    whoEl.dataset.name = ''; whoEl.dataset.userid = ''
    authed = false; renderState(false)
    logoutBtn.disabled = false
  })

  sendBtn.addEventListener('click', send)
  inputEl.addEventListener('keydown', function (e) { if (e.key === 'Enter') send() })

  // 跟随 OnlyOffice 编辑器主题（dark/light）
  function applyTheme(theme) {
    const dark = (theme && theme.type === 'dark') ||
      document.documentElement.classList.contains('theme-dark') ||
      document.body.classList.contains('theme-dark')
    document.body.classList.toggle('dark', !!dark)
  }

  let started = false
  async function start() {
    if (started) return
    started = true
    applyTheme()
    if (window.Asc && window.Asc.plugin && window.Asc.plugin.onThemeChanged) {
      window.Asc.plugin.onThemeChanged = function (theme) { applyTheme(theme) }
    }
    try {
      loginEnabled = (await window.WecomAuth.getConfig()).enabled
    } catch { loginEnabled = false }
    f = await getDocId()
    renderState(false)
    poll()
    setInterval(poll, 1000)
  }

  if (window.Asc && window.Asc.plugin) {
    window.Asc.plugin.init = start
    window.Asc.plugin.button = function (id) { if (id === -1) this.executeCommand('close', '') }
  }
  // 看门狗：某些版本插件不回调 init，1.5s 后强启
  setTimeout(start, 1500)
})()
