# onlyoffice-web-plugins

ONLYOFFICE Document Server 插件集，符合官方插件规范（guid / version / group / icons2 五档图标）。

## 插件列表

| 插件 | 功能 | 后端依赖 |
|---|---|---|
| plugin-ai | AI 助手：选中文本发送 AI 处理并写回文档 | 有（AI API 代理） |
| plugin-dsh / plugin-wecom | 已从主项目 [onlyoffice-web](https://github.com/xianmua/onlyoffice-web) 迁出，本仓库为其分发备份 | 有（各自的配套后端） |

> 注：三个插件均依赖配套后端提供 API。单独安装后插件 UI 可正常加载，但业务功能需自行对接后端。

## 安装方式

1. **zip 上传**：把插件目录内容打包 zip（**config.json 必须在 zip 根级**），在 ONLYOFFICE「设置 → 插件 → 管理」上传安装
2. **目录直放**：拷贝到 Document Server 的 `sdkjs-plugins/` 目录（docker 容器内为 `/var/www/onlyoffice/documentserver/sdkjs-plugins/`），重启生效
3. **集成方注入**：在编辑器配置中指定：

```js
editorConfig.plugins = {
  autostart: ["asc.{7A1B9C0D-5E2F-4A3B-8C4D-000000000001}"],
  pluginsData: ["https://your-host/plugins/plugin-ai/config.json"]
}
```

## 图标说明

各插件 `resources/light|dark/` 下为五档 DPI 官方命名图标（icon.png / @1.25x / @1.5x / @1.75x / @2x + svg 兜底），可用 [gen-plugin-icons.py](https://github.com/xianmua/onlyoffice-web) 批量再生成。
