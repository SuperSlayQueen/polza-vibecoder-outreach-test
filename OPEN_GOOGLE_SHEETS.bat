@echo off
chcp 65001 >nul
echo.
echo === Google Sheet за 60 секунд ===
echo 1) Сейчас откроются: папка с XLSX и пустая Google Таблица
echo 2) В Sheets: Файл - Импорт - Загрузка - Polza_Test_Submission.xlsx
echo 3) Заменить таблицу - Импорт
echo 4) Настройки доступа - Все по ссылке - Читатель
echo 5) Скопируй ссылку в ответ рекрутеру
echo.
explorer "%~dp0data"
start https://sheets.new
start https://github.com/SuperSlayQueen/polza-vibecoder-outreach-test/blob/master/data/Polza_Test_Submission.xlsx
pause
