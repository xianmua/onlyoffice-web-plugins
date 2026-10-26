// 共享企微登录助手：plugin-chat 面板经同源 <script src> 复用（弹窗授权 + 登录态轮询）。
// startLogin()：弹授权窗（企微内静默 / PC 扫码）→ 轮询 /api/auth/me 感知落 cookie →
// 收场返回结果。授权窗被浏览器拦截时返回 blocked，由调用方回落到「按钮点击再试」
//（按钮点击有用户手势，弹窗必成功）。
;(function () {
  window.WecomAuth = {
    apiBase: window.location.origin, // 插件 iframe 与前端/后端同源加载，直接取即可

    async getConfig() {
      return (await fetch(this.apiBase + '/api/auth/wecom/config', { cache: 'no-store' })).json()
    },

    async me() {
      return (await fetch(this.apiBase + '/api/auth/me', { cache: 'no-store' })).json()
    },

    async logout() {
      try { await fetch(this.apiBase + '/api/auth/logout', { method: 'POST' }) } catch { /* 下次轮询自然降级 */ }
    },

    // 弹授权窗并等待结果。返回 { ok:true, userid, name } 或 { ok:false, reason }
    // reason: 'blocked'（弹窗被拦）| 'closed'（未登录即关窗）| 'timeout'
    async startLogin(intervalMs, timeoutMs) {
      let popup = null
      try {
        popup = window.open(this.apiBase + '/api/auth/wecom/authorize', 'wecom-auth', 'width=620,height=680')
      } catch { /* 下面按被拦截处理 */ }
      if (!popup) return { ok: false, reason: 'blocked' }
      intervalMs = intervalMs || 700
      timeoutMs = timeoutMs || 120000
      const t0 = Date.now()
      return new Promise(function (resolve) {
        const timer = setInterval(async function () {
          // 授权窗被手动关闭且尚未登录：收场（留 1.5s 缓冲，防 authorize 302 期间误判）
          if (popup.closed && Date.now() - t0 > 1500) {
            clearInterval(timer)
            resolve({ ok: false, reason: 'closed' })
            return
          }
          if (Date.now() - t0 > timeoutMs) {
            clearInterval(timer)
            try { popup.close() } catch { /* 已关 */ }
            resolve({ ok: false, reason: 'timeout' })
            return
          }
          try {
            const j = await (await fetch(WecomAuth.apiBase + '/api/auth/me', { cache: 'no-store' })).json()
            if (j.authed) {
              clearInterval(timer)
              resolve({ ok: true, userid: j.userid, name: j.name })
            }
          } catch { /* 网络抖动继续轮询 */ }
        }, intervalMs)
      })
    },
  }
})()
