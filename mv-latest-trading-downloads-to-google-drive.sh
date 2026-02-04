#!/bin/bash
# mv-latest-trading-downloads-to-google-drive.sh
# Stages E*Trade and option price downloads to Google Drive.
# Only processes files modified within the last 24 hours.
#
# 1. DownloadTxnHistory.csv / PortfolioDownload.csv → Investing/Data/Etrade/ (renamed with timestamp)
# 2. *-options-*.csv → Investing/Data/OptionPrices/ (moved as-is)

SRC=~/Downloads
DRIVE=~/"Google Drive/My Drive/Investing/Data"
ETRADE_DEST="$DRIVE/Etrade"
OPTION_DEST="$DRIVE/OptionPrices"

MAX_AGE_HOURS=24
now=$(date +%s)

# Returns 0 (true) if file was modified within MAX_AGE_HOURS
is_recent() {
  local file="$1"
  local mtime=$(stat -f '%m' "$file")
  local age=$(( (now - mtime) / 3600 ))
  [[ $age -lt $MAX_AGE_HOURS ]]
}

# --- E*Trade transaction & portfolio files (rename with creation timestamp) ---
for file in DownloadTxnHistory.csv PortfolioDownload.csv; do
  path="$SRC/$file"

  if [[ ! -f "$path" ]]; then
    echo "Not found: $file (skipping)"
    continue
  fi

  if ! is_recent "$path"; then
    echo "Stale (>24h): $file (skipping)"
    continue
  fi

  ts=$(stat -f '%SB' -t '%Y-%m-%d_%H%M%S' "$path")
  base="${file%.csv}"
  newname="${base}_${ts}.csv"

  echo "$file -> Etrade/$newname"
  mv "$path" "$ETRADE_DEST/$newname"
done

# --- Option price CSVs (move as-is) ---
moved=0
for path in "$SRC"/*-options-*.csv; do
  [[ -f "$path" ]] || continue

  if ! is_recent "$path"; then
    echo "Stale (>24h): $(basename "$path") (skipping)"
    continue
  fi

  file=$(basename "$path")
  echo "$file -> OptionPrices/"
  mv "$path" "$OPTION_DEST/$file"
  ((moved++))
done
if [[ $moved -eq 0 ]]; then
  echo "No recent option price CSVs found"
fi

