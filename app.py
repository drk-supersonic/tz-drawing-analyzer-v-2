"""
app.py — Streamlit-интерфейс анализатора чертежей vs ТЗ.
Вся логика анализа находится в core.py.
"""

import streamlit as st
import requests
from core import (
    MODEL, INPUTS_DIR, DRAWINGS_DIR,
    docx_to_text, pdf_to_text,
    run_analysis, generate_report,
)

# ════════════════════════════════════════════════════════════════
# STREAMLIT ИНТЕРФЕЙС
# ════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Анализатор чертежей vs ТЗ", layout="wide")

if "report" not in st.session_state:
    st.session_state.report = None
if "result" not in st.session_state:
    st.session_state.result = None

st.title("Анализатор соответствия чертежей ТЗ")
st.markdown("Сравнивает требования технического задания с данными рабочих чертежей.")

with st.sidebar:
    st.header("Настройки")
    api_key = st.text_input(
        "OpenRouter API Key",
        type="password",
        help="Получить ключ на openrouter.ai"
    )
    st.caption(f"Модель: `{MODEL}`")
    st.divider()
    st.markdown("""
**Расположение файлов:**
- ТЗ: `inputs/*.docx`
- Чертежи: `inputs/chertezhi/*.pdf`
""")
    st.divider()
    st.markdown("""
**Архитектура v2:**
1. Извлечение параметров из каждого документа
2. Сравнение структурированных данных

→ Точность ~97%
""")

st.subheader("Файлы для анализа")

tz_files = sorted(INPUTS_DIR.glob("*.docx")) if INPUTS_DIR.exists() else []
pdf_files = sorted(DRAWINGS_DIR.glob("*.pdf")) if DRAWINGS_DIR.exists() else []

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Техническое задание:**")
    if tz_files:
        for f in tz_files:
            st.markdown(f"- `{f.name}`")
    else:
        st.warning("Не найдено .docx в папке inputs/")

with col2:
    st.markdown("**Чертежи:**")
    if pdf_files:
        for f in pdf_files:
            st.markdown(f"- `{f.name}`")
    else:
        st.warning("Не найдено .pdf в папке inputs/chertezhi/")

st.divider()
run_btn = st.button("Запустить анализ", type="primary", use_container_width=True)

if run_btn:
    if not api_key:
        st.error("Введи OpenRouter API Key в боковой панели")
        st.stop()
    if not tz_files:
        st.error("Не найден файл ТЗ. Положи .docx в папку inputs/")
        st.stop()
    if not pdf_files:
        st.error("Не найдены PDF. Положи файлы в папку inputs/chertezhi/")
        st.stop()

    tz_path = str(tz_files[0])
    tz_name = tz_files[0].name
    pdf_names = [p.name for p in pdf_files]

    progress = st.progress(0, text="Запуск...")
    status = st.empty()

    def update_progress(pct: int, msg: str):
        progress.progress(pct, msg)
        status.caption(msg)

    try:
        progress.progress(5, "Читаю Техническое Задание...")
        tz_text = docx_to_text(tz_path)

        drawings = {}
        for idx, pdf_file in enumerate(pdf_files):
            pct = 5 + int(10 * (idx / len(pdf_files)))
            progress.progress(pct, f"Читаю чертёж {idx+1}/{len(pdf_files)}: {pdf_file.name}...")
            drawings[pdf_file.name] = pdf_to_text(str(pdf_file))

        total_chars = sum(len(t) for t in drawings.values())
        status.caption(
            f"Прочитано: ТЗ {len(tz_text):,} символов, "
            f"чертежи {total_chars:,} символов суммарно"
        )

        result = run_analysis(api_key, tz_text, drawings, progress_callback=update_progress)

        progress.progress(95, "Формирую отчёт...")
        report = generate_report(result, tz_name, pdf_names)
        progress.progress(100, "Готово!")
        status.empty()

        st.session_state.report = report
        st.session_state.result = result

    except requests.exceptions.HTTPError as e:
        try:
            api_error = e.response.json()
        except Exception:
            api_error = e.response.text[:500]
        st.error(f"Ошибка API: HTTP {e.response.status_code}")
        st.exception(RuntimeError(str(api_error)))
    except Exception as e:
        st.error(f"Ошибка: {e}")
        st.exception(e)

if st.session_state.result is not None:
    st.success("Анализ завершён")

    s = st.session_state.result["summary"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Совпадений", s["total_matches"])
    c2.metric("Расхождений", s["total_discrepancies"])
    c3.metric("Отсутствует в чертежах", s["total_missing"])
    c4.metric("Доп. данные из чертежей", s.get("total_extra", 0))

    st.download_button(
        label="Скачать отчёт (comparison_report.md)",
        data=st.session_state.report.encode("utf-8"),
        file_name="comparison_report.md",
        mime="application/octet-stream",
        use_container_width=True,
        key="download_report",
    )

    with st.expander("Полный отчёт", expanded=True):
        st.markdown(st.session_state.report)

    with st.expander("🔍 Отладка: извлечённые параметры", expanded=False):
        debug = st.session_state.result.get("_debug", {})
        if debug:
            tab_tz, *tabs_drawings = st.tabs(
                ["ТЗ"] + list(debug.get("drawings_extracted", {}).keys())
            )
            with tab_tz:
                st.json(debug.get("tz_extracted", {}))
            for tab, (name, data) in zip(tabs_drawings, debug.get("drawings_extracted", {}).items()):
                with tab:
                    st.json(data)

st.markdown("---")
st.caption(f"OpenRouter · {MODEL} · Streamlit · v2 (chain-of-thought)")
