echo "delete warrants on both TSE/OTC earnings..."
find ./tse_earning_raw_data -name "??????.csv" | xargs rm -f 
find ./otc_earning_raw_data -name "??????.csv" | xargs rm -f 
echo "delete banks/exchanges on both TSE/OTC earnings..."
rm -f ./tse_earning_raw_data/1409.csv
rm -f ./tse_earning_raw_data/1718.csv
rm -f ./tse_earning_raw_data/2514.csv
rm -f ./tse_earning_raw_data/5880.csv
rm -f ./tse_earning_raw_data/2905.csv
rm -f ./tse_earning_raw_data/2801.csv
rm -f ./tse_earning_raw_data/2809.csv
rm -f ./tse_earning_raw_data/2812.csv
rm -f ./tse_earning_raw_data/2820.csv
rm -f ./tse_earning_raw_data/2834.csv
rm -f ./tse_earning_raw_data/2836.csv
rm -f ./tse_earning_raw_data/2838.csv
rm -f ./tse_earning_raw_data/2845.csv
rm -f ./tse_earning_raw_data/2849.csv
rm -f ./tse_earning_raw_data/2855.csv
rm -f ./tse_earning_raw_data/2856.csv
rm -f ./tse_earning_raw_data/6005.csv
rm -f ./tse_earning_raw_data/288?.csv
rm -f ./tse_earning_raw_data/2890.csv
rm -f ./tse_earning_raw_data/2891.csv
rm -f ./tse_earning_raw_data/2892.csv
rm -f ./tse_earning_raw_data/????A.csv
rm -f ./otc_earning_raw_data/4147.csv
rm -f ./otc_earning_raw_data/4707.csv
rm -f ./otc_earning_raw_data/5820.csv
rm -f ./otc_earning_raw_data/6015.csv
rm -f ./otc_earning_raw_data/6016.csv
rm -f ./otc_earning_raw_data/6020.csv
rm -f ./otc_earning_raw_data/6021.csv
rm -f ./otc_earning_raw_data/6022.csv
rm -f ./otc_earning_raw_data/6023.csv
rm -f ./otc_earning_raw_data/6024.csv
rm -f ./otc_earning_raw_data/6026.csv
echo "delete ETF on both TSE/OTC earnings..."
rm -f ./tse_earning_raw_data/00??.csv
rm -f ./tse_earning_raw_data/00???.csv
echo "done!"
