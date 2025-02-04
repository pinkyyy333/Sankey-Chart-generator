# Sankey-Chart-generator
Upload an Excel file to generate a Sankey chart.
<br/> Only the employee hours file is provided.
### step1 下載檔案Sankey_Chart_generator.py
將檔案儲存在C槽或是常用的資料夾中，並將路徑複製下來
![image](https://github.com/user-attachments/assets/2fcb67bb-123b-4229-b574-eb328b343c96)

### step2 先將需要的套件下載
在電腦中的搜尋欄位打上"cmd"或"命令提示字元"輸入下面的指令
```python=
pip install streamlit
pip install pandas
pip install pyecharts
pip install streamlit-echarts
```
### step3 執行檔案
在電腦中的搜尋欄位打上"cmd"或"命令提示字元"
1. 輸入路徑，例如：cd C:\sankey
```
cd 你的路徑
```
![image](https://github.com/user-attachments/assets/4e2a8701-4477-4e8f-a53b-cd3edb671f94)

2. 開啟檔案輸入
```
streamlit run Sankey_Chart_generator.py
```

![image](https://github.com/user-attachments/assets/c6075306-c007-4fce-976b-8fbd0b508d62)
<br/>輸入完後會自動跳出網頁，或是在瀏覽器輸入以下網址也可以
```
Local URL: http://localhost:8501
Network URL: http://192.168.16.129:8501 （該網址每個人會不同）
```
<br/>如果要終止該應用程式只需在終端中按 Ctrl + c
![image](https://github.com/user-attachments/assets/fa1d499c-2341-45b5-aa8c-10f64a34209f)

3. 點選browse file並上傳xlsx檔
![image](https://github.com/user-attachments/assets/e71f335f-29cb-4f0f-986d-5c1d37ba03c5)

4. 點取下拉式選單可選取欲觀察的成員，或是也可以按下全選

5. 接著就可以透過sankey chart來觀察每位成員的工時比例了

