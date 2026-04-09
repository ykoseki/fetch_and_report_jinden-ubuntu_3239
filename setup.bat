@echo off
chcp 65001 > nul
echo ✨ 業務改善報告書ツール：環境セットアップを開始します ✨
echo.
echo 🐍 Pythonライブラリ (pandas, openpyxl) を最新の状態にします...
echo.
python -m pip install --upgrade pip
pip install pandas openpyxl python-dotenv
echo.
echo ----------------------------------------------------------
echo ✅ セットアップが完了しました！
echo 🚀 これで run.bat を使ってレポートを作成できます。
echo 💡 エクセルテンプレートやJSONファイルはもう不要ですっ！💖
echo ----------------------------------------------------------
echo.
pause
