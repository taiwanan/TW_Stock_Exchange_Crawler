echo "delete warrants on both TSE/OTC yearly trading/earning ..."
find ./yearly_trading_raw_data -name "??????.csv" | xargs rm -f 
find ./yearly_trading_raw_data -name "??????.csv" | xargs rm -f 
rm -f ./yearly_trading_raw_data/????A.csv
rm -f ./yearly_trading_raw_data/????B.csv
rm -f ./yearly_trading_raw_data/????E.csv
find ./yearly_earning_raw_data -name "??????.csv" | xargs rm -f 
find ./yearly_earning_raw_data -name "??????.csv" | xargs rm -f 
rm -f ./yearly_earning_raw_data/????A.csv
rm -f ./yearly_earning_raw_data/????B.csv
rm -f ./yearly_earning_raw_data/????E.csv
echo "delete ETF on both TSE/OTC yearlyly trading/earning..."
rm -f ./yearly_trading_raw_data/00??.csv
rm -f ./yearly_trading_raw_data/00???.csv
rm -f ./yearly_trading_raw_data/00???.csv
rm -f ./yearly_trading_raw_data/00???.csv
rm -f ./yearly_earning_raw_data/00??.csv
rm -f ./yearly_earning_raw_data/00???.csv
rm -f ./yearly_earning_raw_data/00???.csv
rm -f ./yearly_earning_raw_data/00???.csv
echo "done!"
