# Max 工作日誌

本專案是 A11 業績總覽與人員管理的靜態網頁。

部署到 GitHub Pages 後，可用手機、個人電腦、公司電腦透過同一個網址開啟。

## 人員資料後端

`apps-script/people-backend.gs` 是人員管理用的 Google Apps Script 後端。
部署為 Web App 後，把部署網址填入 `index.html` 的 `peopleBackendUrl`，人員封存、轉正、離職等狀態就會寫回 Google Sheet。
