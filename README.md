# Polza Agency — тестовое (Вайбкодер-аутричер)

Скрипт персонализации + база + цепочка писем + разбор Task4.

## Быстрый старт

```bash
pip install -r requirements.txt
python personalize.py -i data/companies_hardened.csv -o data/out.xlsx
python personalize.py -i data/task4_input.csv -o data/task4.xlsx --strict
```

Главный скрипт: [`personalize.py`](./personalize.py)

Готовая таблица для сдачи: [`data/Polza_Test_Submission.xlsx`](./data/Polza_Test_Submission.xlsx)

Задача 5: [`task5_vibe_stack.md`](./task5_vibe_stack.md)

## Как открыть как Google Sheet (1 минута)

1. Открой [Google Sheets](https://sheets.google.com) → **Пустой файл**
2. **Файл → Импорт → Загрузка** → выбери `data/Polza_Test_Submission.xlsx`
3. «Заменить таблицу» → Импорт
4. **Настройки доступа → Все, у кого есть ссылка → Читатель**
5. Скопируй ссылку в ответ рекрутеру

Листы: `База`, `Task4`, `Письма`, `Цепочка`, `Методология`, `Task4_разбор`.
