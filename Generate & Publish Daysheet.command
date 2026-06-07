#!/bin/bash
cd "/Users/keeganmillar/Library/CloudStorage/OneDrive-TourismNoosa/Schedule Test/files"

echo "NFW 2026 Day Sheet — Generate & Publish"
echo "----------------------------------------"
echo ""

echo "Step 1/2 — Generating daysheet..."
python generate_daysheet.py
echo ""

echo "Step 2/2 — Pushing to GitHub..."
git add -u
if git diff --cached --quiet; then
  echo "No changes to publish."
else
  git -c user.name="Keegan Millar" -c user.email="keeganmillar@gmail.com" \
    commit -m "Update daysheet $(date '+%d %b %Y, %I:%M%p')"
  git push
  echo ""
  echo "Live in ~60 seconds:"
  echo "https://showtechtrader.github.io/nfw-2026-daysheets/daysheet.html"
fi

echo ""
echo "Press any key to close..."
read -n 1
