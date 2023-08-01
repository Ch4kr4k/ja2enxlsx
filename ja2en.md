## jupyter notebook of japanese to English xlsx converter for single sheet without psswd
E Chakrak <br>
Tm no.896 <br>
[github_link](https://www.github.com/Ch4kr4k/ja23nxlsx)

#### rough proposal of system design
![sys](systemdesign.png)

#### working Concept of japanese to English xlsx converter for single sheet without psswd


```python
# importing libraries
from googletrans import Translator, constants
from pprint import pprint
import pandas as pd
```


```python
_2ja = Translator() 
```


```python
df = pd.read_excel("jap test.xlsx")
```

##### testing mix japanese with english and numeric


```python
print(df.head())
```

      送信側ECUのData Length Code（データ長）を表す。CANフレームの場合DLCは0, 1, 2, 3, 4, 5, 6, 7, 8Byte、 CAN FDフレームの場合DLCは0, 1, 2, 3, 4, 5, 6, 7, 8, 12, 16, 20, 24, 32, 48, 64Byteとなる。\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t  \
    0  送信側ECUのData Length Code（データ長）を表す。CANフレームの場合DLC...                                                                                                                                                     
    
      送信側ECUのData Length Code（データ長）を表す。CANフレームの場合DLCは0, 1, 2, 3, 4, 5, 6, 7, 8Byte、 CAN FDフレームの場合DLCは0, 1, 2, 3, 4, 5, 6, 7, 8, 12, 16, 20, 24, 32, 48, 64Byteとなる。  
    0  送信側ECUのData Length Code（データ長）を表す。CANフレームの場合DLC...                                                                                                            



```python
tmp_text = _2ja.translate(df.iat[0,0], dest="en", src="ja")
df.iat[0,0] = tmp_text.text
```


```python
print(tmp_text.text)
```

    Represents the Data Length Code (data length) of the transmitting ECU. DLC is 0, 1, 2, 3, 4, 5, 6, 7, 8 bytes for CAN frame, and 0, 1, 2, 3, 4, 5, 6, 7, 8, 12 for CAN FD frame , 16, 20, 24, 32, 48, 64 bytes.



```python
for i in range(0,2):  # just visualizing the data
    for j in range(0,2):
        print(f"{i}{j}")
```

    00
    01
    10
    11


Reading the excel sheet that is in japanese


```python
df = pd.read_excel("jap_test2.xlsx", skiprows=0)
df.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>名前はなんですか</th>
      <th>私の名前は</th>
      <th>夢を諦めて死んでください</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>あきらめてはいけない</td>
      <td>天気の子</td>
      <td>すずめ</td>
    </tr>
    <tr>
      <th>1</th>
      <td>君の膵臓を食べたい</td>
      <td>さみしい</td>
      <td>悲しい</td>
    </tr>
    <tr>
      <th>2</th>
      <td>水</td>
      <td>山</td>
      <td>美しい</td>
    </tr>
    <tr>
      <th>3</th>
      <td>女神</td>
      <td>夢</td>
      <td>愛</td>
    </tr>
    <tr>
      <th>4</th>
      <td>心臓</td>
      <td>魂</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
</div>



##### table format
df[cols,rows]

|head|head|
|--|--|
|00|01|
|10|11|


```python
for cols in range (0,5):  # visualization
    for rows in range (0,3):
        print(df.iat[cols, rows])
```

    あきらめてはいけない
    天気の子
    すずめ
    君の膵臓を食べたい
    さみしい
    悲しい
    水
    山
    美しい
    女神
    夢
    愛
    心臓
    魂
    nan


here the below range of rows and cols can be dynamic. by changing it to below
```python
row_range = int(input("row range"))
col_range = int(input("col range"))

```


```python
### function to convert ja to en and save to new excel
def conv2en(df):
    for cols in range (0,5):
        for rows in range (0,3):
            tmp_text = _2ja.translate(df.iat[cols,rows], dest="en", src="ja")
            df.iat[cols,rows] = tmp_text.text
    df.to_excel("test2.xlsx", index=False)
```


```python
conv2en(df)
```


```python
saved_df = pd.read_excel("test2.xlsx", skiprows=0)
saved_df.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>名前はなんですか</th>
      <th>私の名前は</th>
      <th>夢を諦めて死んでください</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>don't give up</td>
      <td>child of the weather</td>
      <td>Sparrow</td>
    </tr>
    <tr>
      <th>1</th>
      <td>i want to eat your pancreas</td>
      <td>lonely</td>
      <td>sad</td>
    </tr>
    <tr>
      <th>2</th>
      <td>water</td>
      <td>Mountain</td>
      <td>beautiful</td>
    </tr>
    <tr>
      <th>3</th>
      <td>Goddess</td>
      <td>dreams</td>
      <td>Love</td>
    </tr>
    <tr>
      <th>4</th>
      <td>heart</td>
      <td>soul</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
</div>




```python

```
