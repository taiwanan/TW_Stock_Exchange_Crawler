echo "delete warrants on both TSE/OTC..."
find ./tse_trading_raw_data -name "??????.csv" | xargs rm -f 
find ./otc_trading_raw_data -name "??????.csv" | xargs rm -f 
echo "done!"
