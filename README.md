# Taiwan Stock Exchange Crawler

這是一個去爬 [台灣證券交易所](http://www.twse.com.tw/) 和 [證券櫃檯買賣中心](http://www.tpex.org.tw/) 的爬蟲，
藉由爬下的資料數據做分析。

## Note
### 檔案說明
抓取上市個股成交資料: fetch_trading.py</br>
產生上市買賣壓力分析表: create_trading_pressure_table.py

抓取大盤指數並產生波段動能圖: fetch_twse_idx.py

抓取上市/櫃財報: fetch_earning.py YYYY Q</br>
產生上市櫃財報分析表: create_earning_chart_xlsx.py (xx 年期間 營收, 毛利, 營益)

抓取上市/櫃每月營收: fetch_income.py YYYY MM</br>
產生上市櫃每月營收分析表: create_monthly_chart.py (defaul: 24個月, 4MA)


### 增加`每月營收`擷取
2018/06/7 - add fetch_income.py

抓取上市及上櫃每月營收資料並儲存在 monthly_raw_data</br> 
</br>
step1: 產生表格 (離今天最近的工作日產生) 
```
python fetch_income.py -i
```
step2: 指定的 年份+月份, 若無指定則自動更新至最新
```
python fetch_income.py YYYY MM
ex. pythong fetch_income.py 2018 2
```
### 增加`財報損益表`擷取

2016/07/22 - add fetch_earnings.py

抓取上市及上櫃損益表資料, 分別儲存在 tse_earning_raw_data 及 otc_earning_raw_data</br>
</br>
step1: 產生表格 (離今天最近的工作日產生) 
```
python fetch_earnings.py -i
```
step2: 指定的 年份+季度
```
python fetch_earnings.py YYYY Q
ex. pythong fetch_earnings.py 2013 1
```

### 增加`TWSE波段動能圖`
2017/07/10 - add fetch_twse_idx.py

自動更新台股指數(TWSE_IDX.csv)並產生往前推一年的台股指數波段動能圖(twse_momentum_table.xlsx)

參考文件: documents/TWSE_Momentum_intro (from: [E大](http://ebigmoney.pixnet.net/blog))


### 增加`買賣壓力表`分析 ( [E大](http://ebigmoney.pixnet.net/blog/post/190015890) )
2018/06/12 - add fetch_trading.py:
抓取個股盤後交易資料後(fetch_trading.py), 自動產生買賣壓力圖表(create_trading_pressure_table.py)

raw data 資料夾: "tse_trading_raw_data" & "otc_trading_raw_data"
個股買賣壓力資料夾: trading_pressure_data
表格檔案: trading_pressure_table.xlsm



Step 1.產生空白表格(離今天最近的工作日產生): 
```
python fetch_trading.py -i
```
Step 2.更新資料至今天: 
```
python fetch_trading.py
```
<note>
	
產生圖表
```
python -u fetch_trading.py -t
```
產生 VB script BIN file:
 - excel_vba_tools: extract vba BIN file from excel (.xlsm)
```
python vba_extract.py xxx.xlsm
```

## Contact Info.
taiwanan @ `g9002012@gmail.com`

## Reference
* [CHANGELOG](CHANGELOG.md)
* 原始程式參考: [股票即時資料爬蟲](https://github.com/Asoul/tsrtc)


