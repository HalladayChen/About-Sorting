# 多媒體第二次作業
## About Excel files transfer into csv files
### Part1: 新增資料貢獻組別欄位
我們第一個目的是要先在新的column當中，增加一個**資料貢獻組別**的欄位，以便辨識是哪一組貢獻的資料。<br> 第一部分程式碼如下圖：<br><br>
![圖一](https://github.com/HalladayChen/About-Sorting/blob/main/1.jpg)<br><br>
![圖二](https://github.com/HalladayChen/About-Sorting/blob/main/2.jpg)<br><br>
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

### Part2:將各組資料利用矩陣方式載入並轉成excel檔
第二部分程式碼如下：<br><br>
![圖三](https://github.com/HalladayChen/About-Sorting/blob/main/image4.png)<br><br>
![圖四](https://github.com/HalladayChen/About-Sorting/blob/main/image5.png)<br><br>
首先，創建兩個sheet(sheet1、sheet2)，以便將各組的產業及研討會資料作區別。for迴圈將會執行8次，因為我們所讀取進的資料總共有8組。<br>
接著，我們將利用pandas函式庫的內容做讀取與存取。<br>
pd.read_excel會將資料依序讀取至sheet1、2當中；pd.concat則會將讀取到的資料分別存至兩個矩陣當中。<br>
待for迴圈執行完畢後，最後再將兩個sheet矩陣當中的內容，利用sheet_df.to_excel轉成excel檔。

### Part3:將轉出的excel檔再轉為csv檔
第三部分程式碼如下：<br><br>
![圖五](https://github.com/HalladayChen/About-Sorting/blob/main/image6.png)<br><br>
最後一個部分，我們將會把剛才所轉出的資料轉為csv檔。<br>
input_file為剛剛所轉出的excel檔。<br>
output_file則作為檔名為'多媒體相關的研討會'的csv檔。<br>
接著，讀取input_file並將資料存至變數df。<br>
最後，再利用df.to_csv將載入的檔案轉為csv檔，print可顯示檔案已被轉檔成功。

## About Excel files Sorting
有關sorting的部分會包含以下幾種演算法：<br><br>
__1. trimspace__ <br>
當我們在讀取所轉好的csv檔時，資料載入後有時會出現空格。因此我們利用trimspace將空格清出。演算法如下：<br><br>
![圖六](https://github.com/HalladayChen/About-Sorting/blob/main/trimspace.png)<br><br>

__2. 讀取雙引號之情況__ <br>
當轉為csv檔時，各筆子資料會以逗號做分隔。但若子資料當中具有逗號的話，存在於其中的逗號前後會以雙引號標記。因此我們在讀取匯入的csv檔時，會撰寫一演算法來為了避免雙引號干擾到資料的正確讀取。如果遇到雙引號時，會從第一個雙引號，持續讀取資料到第二個雙引號之前。演算法如下：<br><br>
![圖七](https://github.com/HalladayChen/About-Sorting/blob/main/%E9%87%9D%E5%B0%8D%E9%9B%99%E5%BC%95%E8%99%9F%E7%9A%84%E6%BC%94%E7%AE%97%E6%B3%95.png)

__3. CompareString__ <br>
當我們匯入資料後，為了去比較每一筆資料的字首以便排序，我們會利用到ASCII code表，並採用此演算法回傳兩個字串，目的是讓我們可以在以下將會提到的qsort演算法當中使用(比較字首ASCII code大小)。若a > b，strcmp將會回傳負值；a < b，strcmp將會回傳正值；a = b，strcmp將會回傳0。演算法如下：<br><br>
![圖八](https://github.com/HalladayChen/About-Sorting/blob/main/compareString.png)

__4. qsort__ <br>
最關鍵的部分就是這個演算法。我們將透過**CompareString**當中比較後的結果，去做快速排序。原理為選擇一個基準元素(key)，將比基準元素小的元素放在基準元素的左邊，將比基準元素大的元素放在右邊，然後遞歸地對左右兩個子數組進行相同的操作，直到每個子數組都被排出順序為止。演算法如下：<br><br>
![圖九](https://github.com/HalladayChen/About-Sorting/blob/main/qsort.png)

## About Excel files Uniquing
Unique演算法主要用途為比較各筆資料，並將相同的資料做刪除。以下為演算法撰寫：<br>
### 副程式 - compareUnique
演算法如下：<br><br>
![圖十](https://github.com/HalladayChen/About-Sorting/blob/main/CompareUnique.png)<br><br>
![圖十一](https://github.com/HalladayChen/About-Sorting/blob/main/CompareUnique2.png)<br><br>
compareUnique的作用是比較前後兩筆資料。若兩筆資料相同，回傳1；反之則回傳0。由於我們在這中間未處理在單項資料中有額外的","出現的問題，所以每筆資料若使用","分割的子資料數量會不同，因此我們會先以子資料的數量作為判斷依據。<br>
首先創建token，用於指向切割後的子資料。另外創建兩個初始矩陣temp，分別用於依次儲存切割後的子資料。<br>
接著第一與二個while迴圈的作用為讀取資料，count會記錄該矩陣當中存取了多少筆子資料，用於比對子資料數量是否相同。<br>
然後當子資料數量相同，我們將進入第三個while迴圈，作用為從兩筆資料的第一個子資料開始依次比對，其中max_count代表"資訊貢獻者"這個column前面總共有多少個子資料，因此我們只會比對到"資訊貢獻者"前，這邊使用strcmp來比對子資料是否一樣，若這些子資料都相同的話這個function會回傳1，反之則會回傳0。<br><br>
其中，我們只會比對扣除掉倒數兩項的子資料(也就是前五項資料)，因此while迴圈當中的範圍為count-2。

### 主程式
演算法如下：<br><br>
![圖十二](https://github.com/HalladayChen/About-Sorting/blob/main/main_function.png)<br><br>
首先創建兩個暫存矩陣temp，result矩陣儲存unique後(即刪除重複行)的資料；temp1矩陣儲存原本的資料，用於compareUnique當中進行切割，以免原始資料被切割後的資料覆寫。<br>
for迴圈將會逐步比對此筆資料與前五筆子資料是否相同。其中，會利用strcpy去複製原始資料以及欲比對資料至暫存矩陣當中。<br>
最後進入if判斷，如果兩筆資料相同，將會刪除欲比對的資料，並保留原始資料，將其存入resultRow中。為了處理可能有很多筆資料相似的情況，我們使用while迴圈檢查是否下一筆資料依舊是重複的，若有重複，則會繼續判斷下一筆資料，且不保留本筆資料。<br>
其中，k用於判斷resultRow已經寫入至第k行。

## Proportion of Work Division
共分為四個項目：<br><br>
__1. 利用python將excel檔轉為csv檔__ <br>
__2. 利用C做Sorting__ <br>
__3. 利用C做Uniquing__ <br>
__4. 利用Markdown製作報告__ <br>

以下為各組員分工占比：<br>




