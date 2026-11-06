import os
import sys
import subprocess
from pathlib import Path

REQUIRED_PACKAGES = {
    "streamlit": "streamlit",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd"
}

def install_and_import(package_import_name, package_pip_name):
    try:
        __import__(package_import_name)
    except ImportError:
        print(f"Пакет '{package_pip_name}' не найден. Установка...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_pip_name])

for import_name, pip_name in REQUIRED_PACKAGES.items():
    install_and_import(import_name, pip_name)

if not os.environ.get("STREAMLIT_RUNNING"):
    os.environ["STREAMLIT_RUNNING"] = "1"
    try:
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())
    except Exception:
        subprocess.run([sys.executable, "-m", "streamlit", "run", sys.argv[0]])
        sys.exit()

import datetime
import io
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.exceptions import InvalidFileException
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Сборка бонусов Meest", layout="centered")
st.title("📦 Расчет бонусов Meest Экспресс")

STORE_TAB_COLUMN = "Табельный номер директора"
BRANCH_TO_TAB = {
    18301: 3639, 18302: 13224, 18304: 47456, 18283: 45928, 18284: 76343,
    18269: 38262, 18268: 11970, 18270: 12797, 22482: 38016, 18282: 68694,
    22483: 76218, 22478: 66552, 18274: 56769, 18273: 71971, 22480: 75339,
    18275: 8642, 18280: 13729, 18279: 46564, 22481: 42120, 18307: 54521,
    22479: 72304, 18300: 15371, 19936: 64801, 19950: 46771, 19935: 47251,
    18286: 8634, 19939: 16508, 18308: 72362
}

def parse_register_file(file):
    df = pd.read_excel(file, header=None)
    
    hdr_r = None
    for r in range(min(10, df.shape[0])):
        row_str = [str(x) for x in df.iloc[r].values if pd.notna(x)]
        if 'Підрозділ' in row_str or '№ відділення' in row_str:
            hdr_r = r
            break
            
    if hdr_r is None:
        return pd.DataFrame()
        
    hdr_main_raw = df.iloc[hdr_r]
    hdr_sub_raw = df.iloc[hdr_r + 1]
    
    current_srv = None
    hdr_main_ffill = []
    for c in range(df.shape[1]):
        val = hdr_main_raw[c]
        if pd.notna(val) and str(val).strip() not in ['Підрозділ', '№ відділення', 'Виконавець', 'ЕДРПОУВиконавця']:
            current_srv = str(val).strip()
        elif pd.notna(val) and str(val).strip() in ['Підрозділ', '№ відділення', 'Виконавець', 'ЕДРПОУВиконавця']:
            current_srv = None
        hdr_main_ffill.append(current_srv)
        
    branch_col = None
    store_name_col = None
    for c in range(df.shape[1]):
        if pd.notna(hdr_main_raw[c]):
            header = str(hdr_main_raw[c]).strip()
            if header == '№ відділення':
                branch_col = c
            elif header == 'Підрозділ':
                store_name_col = c
            
    rows = []
    for r in range(hdr_r + 2, df.shape[0]):
        b_val = df.iloc[r, branch_col] if branch_col is not None else None
        if pd.notna(b_val):
            try:
                b_num = int(float(b_val))
            except ValueError:
                continue
                
            row_data = {
                'Відділення': b_num,
                'Название из реестра': (
                    str(df.iloc[r, store_name_col]).strip()
                    if store_name_col is not None and pd.notna(df.iloc[r, store_name_col])
                    else None
                ),
                'sum_1kg': 0.0,
                'sum_2kg': 0.0,
                'sum_10kg': 0.0,
                'sum_over_10kg': 0.0,
                'qty_1kg': 0,
                'qty_2kg': 0,
                'qty_10kg': 0,
                'qty_over_10kg': 0,
            }
            
            for c in range(df.shape[1]):
                srv = hdr_main_ffill[c]
                sub = str(hdr_sub_raw[c]).strip() if pd.notna(hdr_sub_raw[c]) else ""
                val = df.iloc[r, c]
                if pd.isna(val) or srv is None:
                    continue
                    
                if 'Видача відправлення отримувачу до 1 кг' in srv:
                    if sub == 'Сума': row_data['sum_1kg'] = float(val)
                    elif sub == 'К-сть': row_data['qty_1kg'] = int(val)
                elif 'Видача відправлення отримувачу до 2 кг' in srv:
                    if sub == 'Сума': row_data['sum_2kg'] = float(val)
                    elif sub == 'К-сть': row_data['qty_2kg'] = int(val)
                elif 'Видача відправлення отримувачу до 10 кг' in srv:
                    if sub == 'Сума': row_data['sum_10kg'] = float(val)
                    elif sub == 'К-сть': row_data['qty_10kg'] = int(val)
                elif 'Видача відправлення отримувачу більше 10 кг' in srv:
                    if sub == 'Сума': row_data['sum_over_10kg'] = float(val)
                    elif sub == 'К-сть': row_data['qty_over_10kg'] = int(val)
                    
            rows.append(row_data)
            
    return pd.DataFrame(rows)

st.subheader("1. Динамические файлы (Реестры)")
uploaded_files = st.file_uploader("Загрузите 3 файла реестров (.xls / .xlsx)", type=["xls", "xlsx"], accept_multiple_files=True)

st.subheader("2. Статичные файлы (Справочники)")

app_directory = Path(__file__).resolve().parent
excel_files = sorted(
    (path for path in app_directory.iterdir() if path.is_file() and path.suffix.lower() in {".xlsx", ".xls"}),
    key=lambda path: path.name.lower()
)

def select_reference_file(label, candidates, preferred_name):
    if not candidates:
        st.warning(f"В папке приложения не найден файл для выбора: {label}.")
        selected_file = None
    else:
        preferred_name = preferred_name.casefold()
        default_index = next(
            (index for index, path in enumerate(candidates) if path.name.casefold() == preferred_name),
            0
        )
        selected_file = st.selectbox(
            label,
            options=candidates,
            index=default_index,
            format_func=lambda path: path.name
        )

    return selected_file

def reference_table(source_file, columns):
    if hasattr(source_file, "seek"):
        source_file.seek(0)
    dataframe = pd.read_excel(source_file, sheet_name=0)
    for column in columns:
        if column not in dataframe.columns:
            if column != STORE_TAB_COLUMN:
                raise ValueError(f"В файле отсутствует столбец: {column}")
            if "Відділення" in dataframe.columns:
                dataframe[column] = pd.to_numeric(dataframe["Відділення"], errors="coerce").map(BRANCH_TO_TAB)
            else:
                dataframe[column] = pd.NA
        elif column == STORE_TAB_COLUMN and "Відділення" in dataframe.columns:
            fallback_mapping = pd.to_numeric(dataframe["Відділення"], errors="coerce").map(BRANCH_TO_TAB)
            dataframe[column] = dataframe[column].fillna(fallback_mapping)
    return dataframe.loc[:, columns]

def validated_reference_rows(dataframe, columns, numeric_columns, unique_column):
    rows = []
    seen_keys = set()

    for row_number, values in enumerate(dataframe.loc[:, columns].itertuples(index=False, name=None), start=2):
        if all(pd.isna(value) or str(value).strip() == "" for value in values):
            continue

        row = {}
        for column, value in zip(columns, values):
            if pd.isna(value) or str(value).strip() == "":
                raise ValueError(f"Заполните все поля в строке {row_number}.")

            if column in numeric_columns:
                try:
                    number = float(value)
                except (TypeError, ValueError) as error:
                    raise ValueError(f"В строке {row_number} в поле «{column}» должно быть целое число.") from error
                if not number.is_integer() or number <= 0:
                    raise ValueError(f"В строке {row_number} в поле «{column}» должно быть положительное целое число.")
                row[column] = int(number)
            else:
                row[column] = str(value).strip()

        key = row[unique_column]
        if key in seen_keys:
            raise ValueError(f"Значение «{key}» в столбце «{unique_column}» повторяется.")
        seen_keys.add(key)
        rows.append(row)

    return rows

def save_reference_workbook(source_file, rows, columns):
    source_name = source_file.name
    if Path(source_name).suffix.lower() == ".xls":
        workbook = Workbook()
    elif isinstance(source_file, Path):
        workbook = load_workbook(source_file)
    else:
        source_file.seek(0)
        workbook = load_workbook(source_file)

    worksheet = workbook.worksheets[0]
    last_row = max(worksheet.max_row, len(rows) + 1)

    for column_number, column_name in enumerate(columns, start=1):
        worksheet.cell(row=1, column=column_number).value = column_name
        for row_number in range(2, last_row + 1):
            worksheet.cell(row=row_number, column=column_number).value = None

    for row_number, row in enumerate(rows, start=2):
        for column_number, column_name in enumerate(columns, start=1):
            cell = worksheet.cell(row=row_number, column=column_number)
            value = row[column_name]
            cell.value = value
            if isinstance(value, str):
                cell.data_type = "s"

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()

def render_reference_editor(source_file, columns, numeric_columns, unique_column, editor_key):
    try:
        dataframe = reference_table(source_file, columns)
    except (OSError, ValueError, InvalidFileException) as error:
        st.error(f"Не удалось прочитать справочник: {error}")
        return

    with st.form(f"{editor_key}_form"):
        edited_dataframe = st.data_editor(
            dataframe,
            num_rows="dynamic",
            hide_index=True,
            key=f"{editor_key}_{source_file.name}"
        )
        save_clicked = st.form_submit_button("Сохранить изменения", type="primary")

    if not save_clicked:
        return

    try:
        rows = validated_reference_rows(edited_dataframe, columns, numeric_columns, unique_column)
        workbook_data = save_reference_workbook(source_file, rows, columns)
        if isinstance(source_file, Path) and source_file.suffix.lower() != ".xls":
            source_file.write_bytes(workbook_data)
            st.success(f"Изменения сохранены в файле «{source_file.name}».")
        else:
            download_filename = Path(source_file.name).with_suffix(".xlsx").name
            st.download_button(
                "Скачать обновленный справочник",
                data=workbook_data,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{editor_key}_download"
            )
            if isinstance(source_file, Path):
                st.info("Формат .xls нельзя перезаписать напрямую — скачайте обновленный файл в формате .xlsx.")
            else:
                st.info("Загруженный файл не изменяется на компьютере автоматически — скачайте обновленный вариант.")
    except (OSError, ValueError, InvalidFileException) as error:
        st.error(f"Не удалось сохранить справочник: {error}")

stores_candidates = [
    path for path in excel_files
    if "магазин" in path.name.casefold()
    and "директор" not in path.name.casefold()
    and "деректор" not in path.name.casefold()
]
directors_candidates = [
    path for path in excel_files
    if "директор" in path.name.casefold() or "деректор" in path.name.casefold()
]

stores_ref_file = select_reference_file(
    "Файл со справочником магазинов",
    stores_candidates,
    "магазины мист.xlsx"
)
directors_ref_file = select_reference_file(
    "Файл со справочником директоров",
    directors_candidates,
    "деректора магазинов.xlsx"
)

with st.expander("Заменить справочник файлом с компьютера"):
    stores_ref_upload = st.file_uploader(
        "Другой файл магазинов (необязательно)",
        type=["xlsx", "xls"],
        key="stores_ref_upload"
    )
    directors_ref_upload = st.file_uploader(
        "Другой файл директоров (необязательно)",
        type=["xlsx", "xls"],
        key="directors_ref_upload"
    )

if stores_ref_upload is not None:
    stores_ref_file = stores_ref_upload
if directors_ref_upload is not None:
    directors_ref_file = directors_ref_upload

with st.expander("Редактировать справочники магазинов и директоров"):
    st.caption("Измените нужные ячейки, добавьте строки внизу таблицы или удалите ненужные строки, затем сохраните изменения.")
    stores_tab, directors_tab = st.tabs(["Магазины", "Директора"])
    with stores_tab:
        if stores_ref_file is not None:
            render_reference_editor(
                stores_ref_file,
                ["Магазин", "Відділення", STORE_TAB_COLUMN],
                {"Відділення", STORE_TAB_COLUMN},
                "Відділення",
                "stores_reference"
            )
    with directors_tab:
        if directors_ref_file is not None:
            render_reference_editor(
                directors_ref_file,
                ["Магазин", "Табельний номер", "ПІБ"],
                {"Табельний номер"},
                "Табельний номер",
                "directors_reference"
            )

generate_clicked = st.button("Сформировать отчет")
if generate_clicked:
    st.session_state["generate_report"] = True

if st.session_state.get("generate_report"):
    if len(uploaded_files) != 3:
        st.error("Пожалуйста, загрузите ровно 3 файла реестров!")
        st.session_state["generate_report"] = False
    elif not stores_ref_file or not directors_ref_file:
        st.error("Пожалуйста, загрузите оба статичных справочника!")
        st.session_state["generate_report"] = False
    else:
        # Парсинг 3 реестров
        list_dfs = [parse_register_file(f) for f in uploaded_files]
        df_registers = pd.concat(list_dfs, ignore_index=True)

        register_names = (
            df_registers.dropna(subset=["Название из реестра"])
            .groupby("Відділення")["Название из реестра"]
            .first()
        )
        df_aggregated = df_registers.groupby('Відділення', as_index=False).sum(numeric_only=True)
        df_aggregated["Название из реестра"] = df_aggregated["Відділення"].map(register_names)
        df_aggregated['Кіль-ть видач'] = (
            df_aggregated['qty_1kg'] + df_aggregated['qty_2kg']
            + df_aggregated['qty_10kg'] + df_aggregated['qty_over_10kg']
        )
        df_aggregated['Бонусы_суммарно'] = (
            df_aggregated['sum_1kg'] + df_aggregated['sum_2kg']
            + df_aggregated['sum_10kg'] + df_aggregated['sum_over_10kg']
        )
        
        store_columns = ["Магазин", "Відділення", STORE_TAB_COLUMN]
        director_columns = ["Магазин", "Табельний номер", "ПІБ"]
        df_stores_ref = reference_table(stores_ref_file, store_columns)
        df_directors_ref = reference_table(directors_ref_file, director_columns)
        df_stores_ref["Відділення"] = pd.to_numeric(df_stores_ref["Відділення"], errors="coerce")
        df_aggregated["Відділення"] = pd.to_numeric(df_aggregated["Відділення"], errors="coerce")
        known_branches = set(df_stores_ref["Відділення"].dropna().astype(int))
        missing_branches = sorted(
            set(df_aggregated["Відділення"].dropna().astype(int)) - known_branches
        )

        if missing_branches:
            st.warning(
                "В отчет попали отделения, которых нет в справочнике. "
                "Заполните данные для каждого магазина — номер отделения уже указан."
            )
            st.subheader("Добавить магазины в справочники")
            st.caption(
                "Новый магазин будет добавлен в справочник магазинов. "
                "Выбранный из списка директор уже есть в справочнике; новый директор будет добавлен туда автоматически."
            )

            director_by_tab = {}
            for _, director in df_directors_ref.iterrows():
                try:
                    tab_number = int(float(director["Табельний номер"]))
                except (TypeError, ValueError):
                    continue
                director_name = str(director["ПІБ"]).strip()
                if director_name and director_name.casefold() != "nan":
                    director_by_tab[tab_number] = director_name

            director_options = ["__new__", None, *sorted(director_by_tab)]

            def format_director_option(value):
                if value is None:
                    return "Выберите директора"
                if value == "__new__":
                    return "Добавить нового директора"
                return f"{value} — {director_by_tab[value]}"

            entered_stores = {}
            for branch in missing_branches:
                st.markdown(f"**Отделение {branch}**")
                suggested_shop_name = df_aggregated.loc[
                    df_aggregated["Відділення"] == branch,
                    "Название из реестра"
                ].iloc[0]
                suggested_shop_name = (
                    suggested_shop_name
                    if isinstance(suggested_shop_name, str) and suggested_shop_name.strip()
                    else ""
                )
                shop_name = st.text_input(
                    "Название магазина",
                    value=suggested_shop_name,
                    key=f"missing_shop_name_{branch}"
                )
                director_choice = st.selectbox(
                    "Директор",
                    options=director_options,
                    format_func=format_director_option,
                    key=f"missing_shop_director_{branch}"
                )
                new_tab_number = None
                new_director_name = None
                if director_choice == "__new__":
                    new_tab_number = st.number_input(
                        "Табельный номер нового директора",
                        min_value=1,
                        step=1,
                        key=f"missing_shop_tab_number_{branch}"
                    )
                    new_director_name = st.text_input(
                        "ФИО нового директора",
                        key=f"missing_shop_director_name_{branch}"
                    )
                entered_stores[branch] = {
                    "Магазин": shop_name.strip(),
                    "Директор": director_choice,
                    "Новый табельный номер": new_tab_number,
                    "Новое ФИО": new_director_name
                }

            save_missing_stores = st.button(
                "Сохранить магазины и продолжить отчет",
                type="primary",
                key="save_missing_stores"
            )
            if not save_missing_stores:
                st.stop()

            new_store_rows = []
            new_directors_by_tab = {}
            try:
                for branch, values in entered_stores.items():
                    if not values["Магазин"]:
                        raise ValueError(f"Укажите название магазина для отделения {branch}.")
                    director_choice = values["Директор"]
                    if director_choice is None:
                        raise ValueError(f"Выберите директора для отделения {branch}.")
                    if director_choice == "__new__":
                        tab_number = int(values["Новый табельный номер"])
                        director_name = (values["Новое ФИО"] or "").strip()
                        if not director_name:
                            raise ValueError(f"Укажите ФИО директора для отделения {branch}.")
                        existing_new_director = new_directors_by_tab.get(tab_number)
                        if existing_new_director and existing_new_director["ПІБ"] != director_name:
                            raise ValueError(
                                f"Для табельного номера {tab_number} указаны разные ФИО."
                            )
                        new_directors_by_tab.setdefault(tab_number, {
                            "Магазин": values["Магазин"],
                            "Табельний номер": tab_number,
                            "ПІБ": director_name
                        })
                    else:
                        tab_number = director_choice
                    new_store_rows.append({
                        "Магазин": values["Магазин"],
                        "Відділення": branch,
                        STORE_TAB_COLUMN: tab_number
                    })

                new_director_rows = list(new_directors_by_tab.values())
                updated_stores = pd.concat(
                    [df_stores_ref, pd.DataFrame(new_store_rows, columns=store_columns)],
                    ignore_index=True
                )
                store_rows = validated_reference_rows(
                    updated_stores,
                    store_columns,
                    {"Відділення", STORE_TAB_COLUMN},
                    "Відділення"
                )
                store_workbook_data = save_reference_workbook(stores_ref_file, store_rows, store_columns)

                updated_directors = df_directors_ref
                director_workbook_data = None
                if new_director_rows:
                    updated_directors = pd.concat(
                        [df_directors_ref, pd.DataFrame(new_director_rows, columns=director_columns)],
                        ignore_index=True
                    )
                    director_rows = validated_reference_rows(
                        updated_directors,
                        director_columns,
                        {"Табельний номер"},
                        "Табельний номер"
                    )
                    director_workbook_data = save_reference_workbook(
                        directors_ref_file,
                        director_rows,
                        director_columns
                    )

                if isinstance(stores_ref_file, Path) and stores_ref_file.suffix.lower() != ".xls":
                    stores_ref_file.write_bytes(store_workbook_data)
                    st.success(f"Справочник магазинов обновлен: {stores_ref_file.name}.")
                else:
                    st.download_button(
                        "Скачать обновленный справочник магазинов",
                        data=store_workbook_data,
                        file_name=Path(stores_ref_file.name).with_suffix(".xlsx").name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_updated_stores"
                    )

                if director_workbook_data is not None:
                    if isinstance(directors_ref_file, Path) and directors_ref_file.suffix.lower() != ".xls":
                        directors_ref_file.write_bytes(director_workbook_data)
                        st.success(f"Справочник директоров обновлен: {directors_ref_file.name}.")
                    else:
                        st.download_button(
                            "Скачать обновленный справочник директоров",
                            data=director_workbook_data,
                            file_name=Path(directors_ref_file.name).with_suffix(".xlsx").name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_updated_directors"
                        )

                df_stores_ref = updated_stores
                df_directors_ref = updated_directors
            except (OSError, ValueError, InvalidFileException) as error:
                st.error(f"Не удалось сохранить новые записи: {error}")
                st.stop()

        st.session_state["generate_report"] = False
        merged = df_stores_ref.merge(df_aggregated, on="Відділення", how="left")
        metric_columns = [
            "sum_1kg", "sum_2kg", "sum_10kg", "sum_over_10kg",
            "qty_1kg", "qty_2kg", "qty_10kg", "qty_over_10kg",
            "Кіль-ть видач", "Бонусы_суммарно"
        ]
        merged[metric_columns] = merged[metric_columns].fillna(0)
        merged["Табельный номер"] = pd.to_numeric(merged[STORE_TAB_COLUMN], errors="coerce")
        mapped_tabs = merged["Відділення"].map(BRANCH_TO_TAB)
        merged["Табельный номер"] = merged["Табельный номер"].fillna(mapped_tabs)
        
        # Подтягивание ФИО директоров по табельному номеру из справочника директоров
        df_directors_ref_clean = df_directors_ref.dropna(subset=['Табельний номер']).copy()
        df_directors_ref_clean['Табельный номер'] = df_directors_ref_clean['Табельний номер'].astype(int)
        
        final_df = merged.merge(df_directors_ref_clean[['Табельный номер', 'ПІБ']], on='Табельный номер', how='left')
        
        # Определение наименования прошлого месяца на украинском
        ukr_months_upper = [
            "СІЧЕНЬ", "ЛЮТИЙ", "БЕРЕЗЕНЬ", "КВІТЕНЬ", "ТРАВЕНЬ", "ЧЕРВЕНЬ",
            "ЛИПЕНЬ", "СЕРПЕНЬ", "ВЕРЕСЕНЬ", "ЖОВТЕНЬ", "ЛИСТОПАД", "ГРУДЕНЬ"
        ]
        ukr_months_lower = [
            "січень", "лютий", "березень", "квітень", "травень", "червень",
            "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"
        ]
        
        today = datetime.date.today()
        first_day_this_month = today.replace(day=1)
        last_month_date = first_day_this_month - datetime.timedelta(days=1)
        prev_month_idx = last_month_date.month - 1
        
        bonus_col_name = f"Бонуси за {ukr_months_upper[prev_month_idx]}, грн"
        output_filename = f"бонусы {ukr_months_lower[prev_month_idx]}.xlsx"
        
        # Формирование таблицы в точности как в образце
        final_report = pd.DataFrame({
            '№': range(1, len(final_df) + 1),
            'Магазин': final_df['Магазин'],
            'Відділення': final_df['Відділення'],
            'Кіль-ть видач': final_df['Кіль-ть видач'].astype(int),
            bonus_col_name: final_df['Бонусы_суммарно'].astype(int),
            'ПІБ директора': final_df['ПІБ'].fillna('—'),
            'Таб.  номер': final_df['Табельный номер'].fillna('—')
        })
        
        # Добавление строки "Загалом"
        total_row = pd.DataFrame([{
            '№': None,
            'Магазин': 'Загалом',
            'Відділення': None,
            'Кіль-ть видач': final_report['Кіль-ть видач'].sum(),
            bonus_col_name: final_report[bonus_col_name].sum(),
            'ПІБ директора': None,
            'Таб.  номер': None
        }])
        
        final_report_with_total = pd.concat([final_report, total_row], ignore_index=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sheet_name = f'MeestБонуси{ukr_months_upper[prev_month_idx]}'
            final_report_with_total.to_excel(writer, index=False, sheet_name=sheet_name)

            worksheet = writer.sheets[sheet_name]
            bonus_col_idx = final_report_with_total.columns.get_loc(bonus_col_name) + 1
            for column_name in ('Магазин', 'ПІБ директора'):
                column_idx = final_report_with_total.columns.get_loc(column_name) + 1
                for row_idx in range(1, worksheet.max_row + 1):
                    worksheet.cell(row=row_idx, column=column_idx).alignment = Alignment(horizontal='left')

            worksheet.cell(row=1, column=bonus_col_idx).font = Font(bold=True)
            for row_idx in range(2, worksheet.max_row + 1):
                worksheet.cell(row=row_idx, column=bonus_col_idx).number_format = '# ##0'

            total_row_idx = worksheet.max_row
            total_fill = PatternFill(fill_type='solid', fgColor='FFF2CC')
            for cell in worksheet[total_row_idx]:
                cell.fill = total_fill
                font = copy(cell.font)
                font.sz = 12
                font.bold = cell.column in (2, bonus_col_idx)
                cell.font = font
        
        st.success("Отчет сформирован 1-в-1 с образцом!")
        st.dataframe(final_report_with_total)
        
        st.download_button(
            label=f"📥 Скачать {output_filename}",
            data=output.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )