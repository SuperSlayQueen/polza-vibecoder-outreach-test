# Polza Agency — тестовое (Вайбкодер-аутричер)

## Запуск

```bash
pip install -r requirements.txt
python personalize.py -i data/companies.csv -o data/out.xlsx
python personalize.py -i data/task4_input.csv -o data/task4_out.xlsx --strict
```

## Состав

| Файл | Назначение |
|------|------------|
| `personalize.py` | Скрипт персонализации (задачи 2 и 4) |
| `data/Polza_Test_Submission.xlsx` | Итоговая таблица: База, Task4, Письма |
| `data/companies.csv` | Входная база для прогона скрипта |
| `data/task4_input.csv` | 15 компаний из задачи 4 |
| `email_sequence.md` | Цепочка из 3 писем |
| `task5_vibe_stack.md` | Вайбкод-стек (задача 5) |
| `TASK4_NOTES.md` | Разбор ловушек в задаче 4 |

