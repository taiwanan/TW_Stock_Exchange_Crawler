#!/bin/bash
today=$(date '+%Y/%m')
commit="TSE raw data/TWSE_IDX update to $today" 
python -u fetch_tse.py
python -u fetch_twse_idx.py
git add TWSE_IDX.csv twse_momentum_table.xlsx tse_raw_data/.
git commit -m "$commit"
#/usr/bin/git push
