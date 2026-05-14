"""
run.py — запуск анализатора чертежей vs ТЗ в терминале (без Streamlit).

Использование:
    python3 run.py

Требования:
    Python 3.9+
    pip3 install pymupdf python-docx requests

Структура папок:
    inputs/
        TZ_zadanie.docx
        chertezhi/
            01.pdf
            02.pdf
    outputs/        ← сюда сохраняются отчёты
    run.py
    app.py
    core.py
"""

import sys
import os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo  # встроено в Python 3.9+

if sys.version_info < (3, 9):
    print(f"Ошибка: нужен Python 3.9+, у вас {sys.version}")
    print("Запустите через: python3 run.py")
    sys.exit(1)

# core.py лежит рядом с run.py — импортируем напрямую
sys.path.insert(0, str(Path(__file__).parent))
from core import docx_to_text, pdf_to_text, run_analysis, generate_report

# ════════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ════════════════════════════════════════════════════════════════

MOSCOW_TZ = ZoneInfo("Europe/Moscow")

MONTH_EN = {
    1: "jan", 2: "feb", 3: "mar", 4: "apr",
    5: "may", 6: "jun", 7: "jul", 8: "aug",
    9: "sep", 10: "oct", 11: "nov", 12: "dec",
}

def report_filename(dt: datetime) -> str:
    """report_14_may_2026_09_16.md"""
    return (
        f"report_{dt.day:02d}_{MONTH_EN[dt.month]}_{dt.year}"
        f"_{dt.hour:02d}_{dt.minute:02d}.md"
    )

def print_progress(pct: int, msg: str):
    bar_len = 30
    filled = int(bar_len * pct / 100)
    bar = "█" * filled + "░" * (bar_len - filled)
    print(f"\r  [{bar}] {pct:3d}%  {msg:<55}", end="", flush=True)
    if pct == 100:
        print()

def ask_api_key() -> str:
    print()
    print("─" * 60)
    print("  Анализатор чертежей vs ТЗ  (терминальный режим)")
    print("─" * 60)
    print()

    # Можно задать заранее: export OPENROUTER_API_KEY="sk-or-..."
    key = os.environ.get("OPENROUTER_API_KEY", "").strip()
    if key:
        print("  API ключ получен из переменной окружения OPENROUTER_API_KEY")
        return key

    print("  Введите OpenRouter API Key:")
    key = input("  > ").strip()
    if not key:
        print("\nОшибка: API ключ не введён")
        sys.exit(1)
    return key

def find_files(inputs_dir: Path, drawings_dir: Path):
    tz_files = sorted(inputs_dir.glob("*.docx")) if inputs_dir.exists() else []
    pdf_files = sorted(drawings_dir.glob("*.pdf")) if drawings_dir.exists() else []

    print()
    print("  Файлы для анализа:")
    print()

    if not tz_files:
        print(f"  ✗ Техническое задание: не найдено в {inputs_dir}/")
        print("    Положите .docx файл в папку inputs/")
        sys.exit(1)
    print(f"  ✓ Техническое задание: {tz_files[0].name}")

    if not pdf_files:
        print(f"  ✗ Чертежи: не найдены в {drawings_dir}/")
        print("    Положите PDF файлы в папку inputs/chertezhi/")
        sys.exit(1)
    for pdf in pdf_files:
        print(f"  ✓ Чертёж: {pdf.name}")
    print()

    return tz_files[0], pdf_files

def print_summary(result: dict):
    s = result.get("summary", {})
    print()
    print("─" * 60)
    print("  Результаты:")
    print(f"    Совпадений:              {s.get('total_matches', 0)}")
    print(f"    Расхождений:             {s.get('total_discrepancies', 0)}")
    print(f"    Отсутствует в чертежах:  {s.get('total_missing', 0)}")
    print(f"    Доп. данные из чертежей: {s.get('total_extra', 0)}")
    print()

    discrepancies = result.get("discrepancies", [])
    if discrepancies:
        print("  Расхождения:")
        for d in discrepancies:
            crit = d.get("criticality", "")
            icon = {"КРИТИЧНО": "✗✗", "ВАЖНО": "✗ ", "НЕЗНАЧИТЕЛЬНО": "~ "}.get(crit, "  ")
            print(f"    {icon} [{crit}] {d.get('parameter', '')}")
            print(f"         ТЗ:     {d.get('requirement', '')}")
            print(f"         Чертёж: {d.get('drawing', '')}")
        print()

    recs = s.get("recommendations", [])
    if recs:
        print("  Рекомендации:")
        for r in recs:
            print(f"    • {r}")
        print()

# ════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════

def main():
    base_dir = Path(__file__).parent
    inputs_dir = base_dir / "inputs"
    drawings_dir = inputs_dir / "chertezhi"
    outputs_dir = base_dir / "outputs"
    outputs_dir.mkdir(exist_ok=True)

    api_key = ask_api_key()
    tz_path, pdf_files = find_files(inputs_dir, drawings_dir)

    # Чтение файлов
    print("  Читаю файлы...")
    tz_text = docx_to_text(str(tz_path))
    print(f"    ТЗ: {len(tz_text):,} символов")

    drawings = {}
    for pdf_file in pdf_files:
        text = pdf_to_text(str(pdf_file))
        drawings[pdf_file.name] = text
        print(f"    {pdf_file.name}: {len(text):,} символов")

    # Анализ
    print()
    print("  Запускаю анализ (3 запроса к API)...")
    print()

    try:
        result = run_analysis(api_key, tz_text, drawings, progress_callback=print_progress)
    except Exception as e:
        print(f"\n\n  Ошибка при анализе: {e}")
        sys.exit(1)

    print_summary(result)

    # Сохранение отчёта
    now = datetime.now(MOSCOW_TZ)
    filename = report_filename(now)
    report_path = outputs_dir / filename

    report_text = generate_report(result, tz_path.name, [p.name for p in pdf_files])
    report_path.write_text(report_text, encoding="utf-8")

    print("─" * 60)
    print(f"  Отчёт сохранён: outputs/{filename}")
    print("─" * 60)
    print()


if __name__ == "__main__":
    main()
