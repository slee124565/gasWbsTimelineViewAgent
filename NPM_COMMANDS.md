## NPM Script 指令 (適用於 Clasp CLI 用戶)

如果你使用 `clasp` 命令列工具來管理和部署 Apps Script 專案，`package.json` 中提供了一些便捷的 NPM script：

*   **環境切換** (將 `.clasp-dev.json` 或 `.clasp-ipd.json` 複製為 `.clasp.json`):
    *   `npm run env:dev`: 切換至開發環境配置。
    *   `npm run env:ipd`: 切換至 IPD (內部生產部署) 環境配置。

*   **程式碼推送** (將本地程式碼上傳到 Apps Script 專案，不會創建新的部署版本):
    *   `npm run push:dev`: 登入 clasp 並推送開發環境程式碼。
    *   `npm run push:ipd`: 登入 clasp 並推送 IPD 環境程式碼。

*   **專案部署** (創建一個新的 Apps Script 部署版本):
    *   `npm run deploy:dev`: 部署開發版程式碼，部署描述會自動帶入 `code.js` 中的 `SCRIPT_VERSION`。
    *   `npm run deploy:ipd`: 部署 IPD 版程式碼，部署描述會自動帶入 `code.js` 中的 `SCRIPT_VERSION`。

**使用範例:**

```bash
npm run env:dev      # 切換到開發環境
npm run deploy:dev   # 推送程式碼並部署為開發版
```
