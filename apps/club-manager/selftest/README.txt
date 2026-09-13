ClubManager 的 --selftest 不需要額外的測試資料。

它會直接用打包在 web/ 裡的公開網頁（index.html、schools/ 九校版與清單）當
測試對象：起一個隨機埠的本機伺服器，逐頁 GET，檢查狀態碼 200、
Content-Type 為 text/html; charset=utf-8、內容含「社團」與 localStorage
字樣，並驗證 /__clubmanager_ping 與「禁止目錄外存取」。

web/ 內容是公開發布的網站原始檔，不含任何真實學生資料；
系統實際使用時的名單只存在使用者瀏覽器的 localStorage，不會進到這支程式。
