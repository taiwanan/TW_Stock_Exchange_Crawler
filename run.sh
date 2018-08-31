#!/bin/bash
today=$(date '+%Y/%m')
commit="TSE/OTC raw data & TWSE_IDX update to $today"
python -u fetch_trading.py
python -u fetch_twse_idx.py
git add TWSE_IDX.csv twse_momentum_table.xlsx tse_trading_raw_data/. otc_trading_raw_data/.
git commit -m "$commit"
#/usr/bin/git push
