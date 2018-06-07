echo "delete warrants on both TSE/OTC monthly revenue..."
find ./monthly_raw_data -name "??????.csv" | xargs rm -f 
find ./monthly_raw_data -name "??????.csv" | xargs rm -f 
rm -f ./monthly_raw_data/????A.csv
rm -f ./monthly_raw_data/????B.csv
rm -f ./monthly_raw_data/????E.csv
rm -f ./monthly_raw_data/91??.csv		# -DR type stock 
rm -f ./monthly_raw_data/3711.csv 
echo "delete ETF on both TSE/OTC monthly revenue..."
rm -f ./monthly_raw_data/00??.csv
rm -f ./monthly_raw_data/00???.csv
rm -f ./monthly_raw_data/00???.csv
rm -f ./monthly_raw_data/00???.csv
echo "done!"
