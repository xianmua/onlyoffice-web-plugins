// 部署配置：后端 API 地址（随插件分发的唯一需要改动项）
// 留空 = 与插件同源（本项目内置注入式加载，App.vue pluginsData，无需改动）
// 独立安装到 Document Server（sdkjs-plugins / zip 上传）时，插件 iframe 的源是 DS，
// 需把这里改成后端实际地址，例如：window.OO_API_BASE = 'http://192.168.1.10:8090'
window.OO_API_BASE = ''
