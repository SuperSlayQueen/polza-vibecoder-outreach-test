# Polza Agency — тестовое (Вайбкодер-аутричер)

**Репозиторий:** https://github.com/SuperSlayQueen/polza-vibecoder-outreach-test

Скрипт персонализации + база 58 РФ B2B (MX ok) + цепочка + Task4 с разбором ловушек.

## Скрипт

```bash
pip install -r requirements.txt
python personalize.py -i data/companies_hardened.csv -o data/out.xlsx
python personalize.py -i data/task4_input.csv -o data/task4.xlsx --strict
```

Файл: [`personalize.py`](./personalize.py)

## Таблица для сдачи

[`data/Polza_Test_Submission.xlsx`](./data/Polza_Test_Submission.xlsx)  
Листы: `База` · `Task4` · `Письма` · `Цепочка` · `Методология` · `Task4_разбор`

### Google Sheet (обязательно по заданию)

1. Открой https://sheets.new  
2. **Файл → Импорт → Загрузка** → `Polza_Test_Submission.xlsx`  
3. «Заменить таблицу» → Импорт  
4. Доступ: «Все, у кого есть ссылка» → Читатель  
5. Ссылку вставь в ответ рекрутеру  

Или запусти `OPEN_GOOGLE_SHEETS.bat` в корне репозитория.

## Задача 5

[`task5_vibe_stack.md`](./task5_vibe_stack.md)
