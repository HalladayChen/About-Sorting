# 多媒體第二次作業
## About Excel files transfer into csv files
### Part1: 新增資料貢獻組別欄位
我們第一個目的是要先在新的column當中，增加一個**資料貢獻組別**的欄位，以便辨識是哪一組貢獻的資料。<br> 第一部分程式碼如下圖：<br><br>
![image](https://github.com/HalladayChen/About-Sorting/blob/main/image1.png)<br><br>
__1. 初始化__

首先，先import各項函式庫，如pandas、openpyxl，以方便我們對excel資料的處理。<br>
接著初始化各項數值，由於我們的資料具有產業以及研討會兩項，因此先建立兩個dataframe。

__2. 建立新欄位__

接著，我們將本機路徑當中各小組的excel檔案資料，一併做載入。這邊做雙重的for迴圈，可以將各小組的產業以及研討會資料依序載入。<br>
其中，在第一個for迴圈內，先定義new_column_name = '資料貢獻組別'，用意是表示我們所新增的欄位名稱。

__3. 讀取、排序資料__
