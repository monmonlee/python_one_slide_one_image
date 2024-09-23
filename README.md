# python_one_slide_one_image
這是一個將已經完成的png逐一貼到ppt上的腳本，每一張圖一個slide
# PowerPoint 生成腳本

## 程式邏輯
本腳本包含三個主要函數：

### a. create_ppt_from_images (路徑, 簡報檔名稱)
- 建立 prs 物件，設定長寬。
- 辨識非為文字雲的圖表，根據每一張符合條件的 .png，並執行 add_image_slide。

### b. add_image_slide(投影片物件，圖片路徑)
- 符合條件的 .png，先建立新的投影片
- 設定圖片大小與位置，添加圖片
- 設定文字框位置與字體，添加投影片標題（標題來自 generate_title）

### c. generate_title(檔案名稱)（已包含在 add_image_slide）
- 到這裡來的檔案已經在 add_image_slide 處理過了，會只剩下檔名（os.path.basename()）
- 剔除擴展名並透過 "_" 切割
- 擷取第一（[0]）和第三（[2]）檔名作為投影片標題

## 使用方法
直接執行 create_ppt_from_images，貼上對應路徑與簡報檔名稱

## 待優化
1. 自動化後，需手動開啟投影片母片（因為 add_image_slide 直接使用佈景主題，圖片位置會不對，需進一步調整參數）

## 解決方法
目前會直接得到空白佈景的完成檔，直接將其複製到有主題的 ppt 中，速度很快，花時間解決投影片母片問題反倒效率較低。
