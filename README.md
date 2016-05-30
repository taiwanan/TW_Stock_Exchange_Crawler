# Taiwan Stock Exchange Crawler

這是一個去爬 [台灣證券交易所](http://www.twse.com.tw/) 和 [證券櫃檯買賣中心](http://www.tpex.org.tw/) 的爬蟲，
藉由爬下的資料數據做分析。

## Note

### 增加`買賣壓力表`分析 ( [E大](http://ebigmoney.pixnet.net/blog/post/190015890) )

2016/05/30 - Modify from crawl.py to fetch_tse.py:

	* https://github.com/taiwanan/TW_Stock_Exchange_Crawler.git

	* modify crawl.py to fetch_tse.py
		1. initial an uniform table with headers: fetch_tse.py -i
		2. crawling data saved as "tse_raw_data" & "otc_raw_data"
		3. default is updating TSE data only up-to-date
		4. rm_warrants.sh to remove warrants from xxx_raw_data

	* generate buy/sell pressure table in:
		1. one table: create_table_xlsx.py
		2. spread table: create_spread_xlsm.py

	* excel_vba_tools: extract vba bin file from excel (.xlsm)
		vba_extract.py xxx.xlsm

## Usage
###  get RAW data
產生空白表格( 產生 tse_raw_data\ 及 otc_raw_data\ ): 
```
python fetch_tse.py -i
```
更新TSE_RAW_DATA至今天: 
```
python fetch_tse.py
```
###  create buy/sell-pressure table
單一檔案:
```
python create_spreadtable_xlsm.py
```
多個檔案:
```
python create_table_xlsx.py
```

## Contact Info.
taiwanan `g9002012@gmail.com`

## Reference
* [CHANGELOG](CHANGELOG.md)
* 原始程式參考: [股票即時資料爬蟲](https://github.com/Asoul/tsrtc)


