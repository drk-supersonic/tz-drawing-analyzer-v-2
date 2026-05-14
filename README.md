# tz-drawing-analyzer

Анализирует соответствие строительных чертежей (PDF) техническому заданию (DOCX) с помощью LLM.

Точность ~98%. Работает в двух режимах: веб-интерфейс (Streamlit) и терминал.

## Как работает

Трёхэтапный анализ (chain-of-thought):

1. **Извлечение** — из ТЗ и каждого чертежа отдельно извлекаются структурированные параметры в JSON
2. **Верификация** — фантомные значения (числа из чужого контекста) убираются отдельным запросом
3. **Сравнение** — структурированные данные сравниваются с универсальными правилами нормализации

## Структура файлов

```
tz-drawing-analyzer/
├── core.py              # вся логика анализа (без Streamlit)
├── app.py               # веб-интерфейс (Streamlit)
├── run.py               # запуск в терминале
├── requirements.txt
├── README.md
├── inputs/              # не в репозитории (.gitignore)
│   ├── TZ_zadanie.docx
│   └── chertezhi/
│       ├── 01.pdf
│       └── 02.pdf
└── outputs/             # не в репозитории (.gitignore)
    └── report_14_may_2026_09_54.md
```

## Установка

```bash
pip install -r requirements.txt
```

## Запуск

### Веб-интерфейс (Streamlit)

```bash
streamlit run app.py
```

Открыть в браузере: http://localhost:8501

Введите OpenRouter API Key в боковой панели и нажмите «Запустить анализ».

### Терминал

```bash
python3 run.py
```

Скрипт спросит API Key, запустит анализ и сохранит отчёт в `outputs/` с именем вида `report_14_may_2026_09_54.md`.

API Key можно задать заранее через переменную окружения — тогда скрипт не будет спрашивать:

```bash
export OPENROUTER_API_KEY="sk-or-..."
python3 run.py
```

## Настройки

- **API** — [OpenRouter](https://openrouter.ai), модель `google/gemini-2.5-flash`
- **ТЗ** — один `.docx` файл в папке `inputs/`
- **Чертежи** — любое количество `.pdf` файлов в папке `inputs/chertezhi/`

## Деплой на Streamlit Cloud

1. Залить репозиторий на GitHub
2. Подключить на [share.streamlit.io](https://share.streamlit.io)
3. Главный файл — `app.py`

Папки `inputs/` и `outputs/` в репозитории отсутствуют (в `.gitignore`) —
файлы загружаются локально перед запуском.
