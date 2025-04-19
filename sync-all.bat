@echo off
echo 🔁 Bắt đầu đồng bộ tất cả project con với GAS và GitHub...

set folders=Library Trader11 Trader21 Web

for %%f in (%folders%) do (
    echo 📂 Xử lý project: %%f
    cd %%f

    echo    ⬇️ Pull từ Google Apps Script
    call clasp pull

    echo    ⬆️ Push lên Google Apps Script
    call clasp push

    cd ..
)

echo 💾 Commit + Push toàn bộ lên GitHub
git add .
git commit -m "🚀 Đồng bộ tất cả GAS project + GitHub"
git push origin master

echo ✅ Đồng bộ hoàn tất!
pause