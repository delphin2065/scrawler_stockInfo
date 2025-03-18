# 期貨交易所爬蟲

1. 使用Selenium抓取PCR資料，為了不讓伺服器發現爬蟲程式抓取資料，需加入擬人的行為，
例如把網頁滑至底部，再將其上滑至頂部原點，方可避開伺服器的檢查

2. 使用requests抓取PCR資料，可以找尋payloads，除了設定原先的headers，
方可在requests.get方法中加入params，即可快速查取資料