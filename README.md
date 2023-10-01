# 多媒體第二次作業
## About Excel files transfer into csv files
### Part1: 新增資料貢獻組別欄位
我們第一個目的是要先在新的column當中，增加一個資料貢獻組別的欄位，以便辨識是哪一組貢獻的資料。<br> 第一部分程式碼如下圖：<br><br>
![image](https://github.com/HalladayChen/About-Sorting/blob/main/image1.png)<br><br>
__1. 初始化__

首先，先import各項函式庫，如pandas、openpyxl，以方便我們對excel資料的處理。<br>
接著初始化各項數值，由於我們的資料具有產業以及研討會兩項，因此建立兩個dataframe。
