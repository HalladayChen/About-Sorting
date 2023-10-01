# 多媒體第二次作業
## About Excel files transfer into csv files
### Part1: 新增資料貢獻組別欄位
我們第一個目的是要先在新的column當中，增加一個**資料貢獻組別**的欄位，以便辨識是哪一組貢獻的資料。<br> 第一部分程式碼如下圖：<br><br>
![圖一](https://github.com/HalladayChen/About-Sorting/blob/main/image1.png)<br><br>
![圖二](https://github.com/HalladayChen/About-Sorting/blob/main/image2.png)<br><br>
__1. 初始化__

首先，先import各項函式庫，如pandas、openpyxl，以方便我們對excel資料的處理。<br>
接著初始化各項數值，由於我們的資料具有產業以及研討會兩項，因此先建立兩個dataframe。

__2. 建立新欄位__

接著，我們將本機路徑當中各小組的excel檔案資料，一併做載入。這邊做雙重的for迴圈，可以將各小組的產業以及研討會資料依序載入。<br>
其中，在第一個for迴圈內，先定義new_column_name = '資料貢獻組別'，用意是表示我們所新增的欄位名稱。

__3. 讀取、載入資料__ <br><br>
然後在第二個for迴圈當中，開始正式做資料的讀取以及排序。<br>
首先，sheet.cell(row=1,column=7,value=new_column_name)這一欄為指定'資料貢獻組別'文字所在的欄位。<br>
file_name_without_extention可將excel檔的格式去除，這樣在載入資料貢獻組別的時候就不會顯示格式，僅留下檔名。<br>
接著，在我們所指定的行與列，逐步讀取資料。while迴圈的作用為讀取excel檔資料，當讀取到空白時即停止。<br><br>
*註：曾嘗試使用openxysl函式庫當中的sheet.max_row來讀取資料，但我們發現可能會造成多讀取的狀況，因此採用while迴圈來處理該問題。*<br><br>
最後，在while迴圈當中所載入的資料數，即可作為第三個for迴圈當中的範圍值，將該先前所讀取到的檔名，依序載入到'資料貢獻組別'這個欄位當中，print即可確認檔名是否已被載入。

### Part2:將各組excel資料利用矩陣方式載入
第二部分程式碼如下：

