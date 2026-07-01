import streamlit as st
import pandas as pd
import numpy as np
import re
import math

# 1. Page Configuration & Styling
st.set_page_config(
    page_title="Панель Управления Рекламоносителями | MVP",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header { font-family: 'Inter', sans-serif; font-weight: 700; color: #1e293b; margin-bottom: 5px; }
    .sub-header { font-family: 'Inter', sans-serif; color: #64748b; margin-bottom: 25px; }
    .metric-card { background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 15px; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-header">💼 Кабинет Менеджера: База Рекламоносителей</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Единый реестр конструкций (Фаза 1 MVP)</p>', unsafe_allow_html=True)

# Helper Column Mappings
COLUMN_ALIASES = {
    'id': ['gid', 'id', 'идентификатор', 'номер', 'код', 'код конструкции', 'espar id'],
    'address': ['адрес поверхности', 'адрес', 'местоположение', 'address', 'расположение'],
    'format': ['тип конструкции', 'формат', 'вид конструкции', 'конструкция', 'format', 'тип'],
    'material': ['материал  постера', 'материал постера', 'материал', 'тип материала'],
    'screens_outputs': ['экраны, выходы', 'выходов в блоке', 'выходов в сутки', 'показы в сутки', 'экраны'],
    'district': ['город/район', 'район', 'округ', 'district', 'муниципальный округ'],
    'direction': ['направление', 'сторона', 'direction', 'ориентация'],
    'lat': ['широта', 'lat', 'latitude', 'координата y', 'y', 'широта_wgs84'],
    'lon': ['долгота', 'lon', 'lng', 'longitude', 'координата x', 'x', 'долгота_wgs84'],
    'price_russ': ['цена', 'прайс', 'стоимость'],
    'price_elvis': ['прайс, в мес, в т.ч. 5% ндс, руб', 'прайс, в мес, без ндс, руб', 'цена без ндс', 'прайс',
                    'стоимость'],

    # ИСПРАВЛЕНИЕ: "доп. монтаж" вынесли в самое начало, чтобы скрипт искал сначала его!
    'installation_price': ['доп. монтаж, c ндс (22 %)', 'доп. монтаж', 'доп монтаж', 'монтаж, c ндс (22 %)',
                           '1й монтаж', 'монтаж'],

    'june': ['июнь', 'june', '06'],
    'july': ['июль', 'july', '07'],
    'august': ['август', 'august', '08'],
    'september': ['сентябрь', 'september', '09'],
    'october': ['октябрь', 'october', '10'],
    'november': ['ноябрь', 'november', '11'],
    'december': ['декабрь', 'december', '12']
}


def load_excel_smart(file):
    try:
        df_temp = pd.read_excel(file, header=None, nrows=30)
        best_row_idx = 0
        max_score = 0
        keywords = ['адрес', 'gid', 'формат', 'тип', 'широта', 'долгота', 'июнь', 'июль', 'цена', 'прайс', 'монтаж']
        for idx, row in df_temp.iterrows():
            row_vals = [str(val).lower() for val in row.values]
            score = sum(1 for val in row_vals for kw in keywords if kw in val)
            if score > max_score:
                max_score = score
                best_row_idx = idx
        file.seek(0)
        return pd.read_excel(file, header=best_row_idx)
    except Exception as e:
        st.error(f"Ошибка при поиске заголовков: {e}")
        return pd.DataFrame()


def extract_coord(series):
    if series is None or series.empty:
        return pd.Series(dtype=float)
    s_str = series.astype(str).str.replace(',', '.', regex=False).str.replace(';', ' ', regex=False).str.strip()

    def extract_float(val):
        if val is None: return None
        if isinstance(val, (int, float)): return None if math.isnan(val) else float(val)
        val_str = str(val).strip()
        if not val_str or val_str.lower() in ['nan', 'none', '']: return None
        match = re.search(r'[-+]?\d*\.\d+|\d+', val_str)
        if match:
            try:
                return float(match.group())
            except ValueError:
                pass
        return None

    return s_str.apply(extract_float)


def get_column_data(df, target_key, aliases, default_val=None, is_numeric=False, is_coord=False):
    cols_normalized = {str(col).strip().lower(): col for col in df.columns}
    matched_col = None
    for alias in aliases:
        if alias.strip().lower() in cols_normalized:
            matched_col = cols_normalized[alias.strip().lower()]
            break
    if not matched_col:
        for alias in aliases:
            for k, original_col in cols_normalized.items():
                if alias.strip().lower() in k or k in alias.strip().lower():
                    if target_key in ['price_russ', 'price_elvis'] and ('монтаж' in k): continue
                    matched_col = original_col
                    break
            if matched_col: break

    if matched_col is not None:
        series = df[matched_col]
        if is_coord:
            return extract_coord(series)
        elif is_numeric:
            series_clean = series.fillna(default_val) if default_val is not None else series.fillna(0)

            def clean_numeric(val):
                if pd.isna(val) or val is None: return default_val
                if isinstance(val, (int, float)): return float(val)
                val_str = str(val).strip().replace(' ', '').replace(' ', '').replace(',', '.')
                match = re.search(r'[-+]?\d*\.\d+|\d+', val_str)
                if match:
                    try:
                        return float(match.group())
                    except:
                        pass
                return default_val

            return series_clean.apply(clean_numeric)
        else:
            return series
    else:
        return pd.Series([default_val] * len(df))


def process_subheaders(df):
    addr_col = next((c for c in df.columns if str(c).strip().lower() in [a.lower() for a in COLUMN_ALIASES['address']]),
                    None)
    format_col = next(
        (c for c in df.columns if str(c).strip().lower() in [a.lower() for a in COLUMN_ALIASES['format']]), None)

    if addr_col and format_col:
        is_subheader = df[format_col].isna() | (df[format_col].astype(str).str.strip() == '') | (
                    df[format_col].astype(str).str.lower() == 'nan')
        df['route_group'] = np.where(is_subheader, df[addr_col], np.nan)
        df['route_group'] = df['route_group'].ffill().fillna('Основной список')
        df = df[~is_subheader].copy()
    else:
        df['route_group'] = 'Основной список'
    return df


@st.cache_data(show_spinner="Загрузка и нормализация данных...")
def load_and_clean_data(file_russ, file_elvis):
    all_dfs = []

    # --- RUSS DATA ---
    if file_russ is not None:
        try:
            df_r = load_excel_smart(file_russ)
            df_r = process_subheaders(df_r)
            df_r = df_r.reset_index(drop=True)

            russ_mapped = pd.DataFrame()
            russ_mapped['provider'] = ['Russ'] * len(df_r)
            russ_mapped['id'] = get_column_data(df_r, 'id', COLUMN_ALIASES['id'], 'unknown').astype(str)
            russ_mapped['address'] = get_column_data(df_r, 'address', COLUMN_ALIASES['address'], 'Не указан').fillna(
                'Не указан')
            russ_mapped['format'] = get_column_data(df_r, 'format', COLUMN_ALIASES['format'], 'Щит 3х6').fillna(
                'Щит 3х6').astype(str)
            russ_mapped['material'] = get_column_data(df_r, 'material', COLUMN_ALIASES['material'], 'Не указан').fillna(
                'Не указан').astype(str)

            russ_district = get_column_data(df_r, 'district', COLUMN_ALIASES['district'], 'Не указан').fillna(
                'Не указан')
            russ_mapped['route_group'] = russ_district

            # Маска для цифровых экранов
            digital_mask = pd.Series(False, index=df_r.index)
            for col in df_r.columns:
                col_name = str(col).lower()
                if any(k in col_name for k in ['способ', 'вид', 'тип', 'формат', 'подтип']):
                    digital_mask = digital_mask | df_r[col].astype(str).str.lower().str.contains(
                        'экран|видео|digital|цифр', regex=True)

            russ_mapped['screens_outputs'] = '-'
            sec_col = next((c for c in df_r.columns if 'выход, сек' in str(c).lower()), None)
            out_col = next((c for c in df_r.columns if 'выходов в сутки' in str(c).lower()), None)

            if sec_col and out_col:
                def build_screen_str(row):
                    s = str(row[sec_col]).replace('nan', '').replace('None', '').strip()
                    o = str(row[out_col]).replace('nan', '').replace('None', '').strip()
                    if not s and not o: return '-'
                    res = []
                    if s: res.append(f"{s} сек")
                    if o: res.append(f"{o} вых/сут.")
                    return " / ".join(res)

                russ_mapped['screens_outputs'] = df_r.apply(build_screen_str, axis=1)

            russ_mapped['direction'] = get_column_data(df_r, 'direction', COLUMN_ALIASES['direction'], '-').fillna('-')

            prices_from_months = []
            for m in ['june', 'july', 'august', 'september', 'october', 'november', 'december']:
                raw_month = get_column_data(df_r, m, COLUMN_ALIASES[m], 'Занято')
                status_list = []
                price_list = []
                for val in raw_month:
                    val_str = str(val).strip().lower()
                    if val_str in ['nan', 'none', '', '-', 'нет данных', '0.0', '0']:
                        status_list.append('Занято')
                        price_list.append(0.0)
                        continue
                    if 'продан' in val_str or 'занят' in val_str:
                        status_list.append('Занято')
                        price_list.append(0.0)
                        continue
                    if 'бронь' in val_str or 'резерв' in val_str:
                        status_list.append('Забронировано')
                        price_list.append(0.0)
                        continue
                    if 'свобод' in val_str:
                        status_list.append('Свободно')
                        price_list.append(0.0)
                        continue

                    val_num = val_str.replace(' ', '').replace(' ', '').replace(',', '.')
                    match = re.search(r'[-+]?\d*\.\d+|\d+', val_num)
                    if match and float(match.group()) > 0:
                        status_list.append('Свободно')
                        price_list.append(float(match.group()))
                    else:
                        status_list.append('Занято')
                        price_list.append(0.0)

                russ_mapped[f'status_{m}'] = status_list
                prices_from_months.append(price_list)

            explicit_price = get_column_data(df_r, 'price', COLUMN_ALIASES['price_russ'], 0, is_numeric=True)
            max_monthly_price = np.max(prices_from_months, axis=0)
            russ_mapped['price'] = np.where(explicit_price > 0, explicit_price, max_monthly_price)

            # МОНТАЖ
            russ_mapped['installation_price'] = get_column_data(df_r, 'installation_price',
                                                                COLUMN_ALIASES['installation_price'], 0,
                                                                is_numeric=True)
            # ИСПРАВЛЕНИЕ: Обнуляем монтаж для цифровых экранов
            russ_mapped.loc[digital_mask, 'installation_price'] = 0

            russ_mapped['lat'] = get_column_data(df_r, 'lat', COLUMN_ALIASES['lat'], None, is_coord=True)
            russ_mapped['lon'] = get_column_data(df_r, 'lon', COLUMN_ALIASES['lon'], None, is_coord=True)
            all_dfs.append(russ_mapped)
        except Exception as e:
            st.error(f"Ошибка чтения Russ: {e}")

    # --- ELVIS DATA ---
    if file_elvis is not None:
        try:
            df_e = load_excel_smart(file_elvis)
            df_e = process_subheaders(df_e)

            id_col_e = next(
                (c for c in df_e.columns if str(c).strip().lower() in [a.lower() for a in COLUMN_ALIASES['id']]), None)
            if id_col_e is not None:
                df_e = df_e.dropna(subset=[id_col_e])
                df_e = df_e[df_e[id_col_e].astype(str).str.strip() != '']

            df_e = df_e.reset_index(drop=True)

            elvis_mapped = pd.DataFrame()
            elvis_mapped['provider'] = ['Элвис'] * len(df_e)
            elvis_mapped['route_group'] = df_e['route_group']

            elvis_mapped['id'] = get_column_data(df_e, 'id', COLUMN_ALIASES['id'], 'unknown').astype(str)
            elvis_mapped['address'] = get_column_data(df_e, 'address', COLUMN_ALIASES['address'], 'Не указан').fillna(
                'Не указан')
            elvis_mapped['format'] = get_column_data(df_e, 'format', COLUMN_ALIASES['format'], 'Щит 3х6').fillna(
                'Щит 3х6').astype(str)
            elvis_mapped['material'] = get_column_data(df_e, 'material', COLUMN_ALIASES['material'],
                                                       'Не указан').fillna('Не указан').astype(str)

            screens_e = get_column_data(df_e, 'screens_outputs', COLUMN_ALIASES['screens_outputs'], '-')
            is_digital_e = elvis_mapped['format'].str.lower().str.contains('экран|видео|digital|цифр', regex=True)
            elvis_mapped['screens_outputs'] = '-'
            elvis_mapped.loc[is_digital_e, 'screens_outputs'] = screens_e.loc[is_digital_e].fillna('-').astype(str)

            elvis_mapped['direction'] = get_column_data(df_e, 'direction', COLUMN_ALIASES['direction'], '-').fillna('-')

            for m in ['june', 'july', 'august', 'september', 'october', 'november', 'december']:
                raw_vals = get_column_data(df_e, m, COLUMN_ALIASES[m], 'Свободно').fillna('Свободно')

                def clean_elvis_status(v):
                    v_str = str(v).strip().lower()
                    if v_str in ['nan', 'none', '', '-', 'нет данных', '0.0']: return 'Свободно'
                    if 'свобод' in v_str: return 'Свободно'
                    if 'бронь' in v_str or 'резерв' in v_str: return 'Забронировано'
                    if 'продан' in v_str or 'занят' in v_str: return 'Занято'
                    return 'Свободно'

                elvis_mapped[f'status_{m}'] = raw_vals.apply(clean_elvis_status)

            elvis_mapped['price'] = get_column_data(df_e, 'price', COLUMN_ALIASES['price_elvis'], 0, is_numeric=True)
            elvis_mapped['installation_price'] = get_column_data(df_e, 'installation_price',
                                                                 COLUMN_ALIASES['installation_price'], 0,
                                                                 is_numeric=True)
            # Для Элвиса тоже обнуляем монтаж у экранов на всякий случай
            elvis_mapped.loc[is_digital_e, 'installation_price'] = 0

            elvis_mapped['lat'] = get_column_data(df_e, 'lat', COLUMN_ALIASES['lat'], None, is_coord=True)
            elvis_mapped['lon'] = get_column_data(df_e, 'lon', COLUMN_ALIASES['lon'], None, is_coord=True)

            all_dfs.append(elvis_mapped)
        except Exception as e:
            st.error(f"Ошибка чтения Элвис: {e}")

    if all_dfs:
        merged_df = pd.concat(all_dfs, ignore_index=True)
        is_spb = (
                (merged_df['lat'].isna()) | (merged_df['lon'].isna()) |
                ((merged_df['lat'] > 59.0) & (merged_df['lat'] < 61.0) &
                 (merged_df['lon'] > 28.0) & (merged_df['lon'] < 32.0))
        )
        return merged_df[is_spb]
    return pd.DataFrame()


# 3. File Upload panel
st.sidebar.markdown("### 📥 Загрузка файлов")
uploaded_russ = st.sidebar.file_uploader("Реестр Russ (.xlsx)", type=["xlsx"])
uploaded_elvis = st.sidebar.file_uploader("Реестр Элвис (.xlsx)", type=["xlsx"])

df = load_and_clean_data(uploaded_russ, uploaded_elvis)

if df.empty:
    st.info("📥 Пожалуйста, загрузите реестр Russ и/или реестр Элвис во вкладке слева для начала работы.")
    filtered_df = pd.DataFrame()
else:
    # 4. Filters Panel
    st.sidebar.markdown("### 🔍 Фильтры Поиска")

    prov_list = ['Все'] + sorted(list(df['provider'].astype(str).unique()))
    selected_prov = st.sidebar.selectbox("Поставщик", prov_list)

    route_list = ['Все'] + sorted(list(df['route_group'].astype(str).unique()))
    selected_route = st.sidebar.selectbox("Район / Направление (Группа)", route_list)

    format_list = ['Все'] + sorted(list(df['format'].astype(str).unique()))
    selected_format = st.sidebar.selectbox("Формат щита", format_list)

    selected_month = st.sidebar.selectbox("Месяц для анализа",
                                          ["Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"])
    month_key_map = {
        "Июнь": "status_june", "Июль": "status_july", "Август": "status_august",
        "Сентябрь": "status_september", "Октябрь": "status_october",
        "Ноябрь": "status_november", "Декабрь": "status_december"
    }
    month_col = month_key_map[selected_month]

    status_list = ['Все', 'Свободно', 'Забронировано', 'Занято']
    selected_status = st.sidebar.selectbox(f"Статус на {selected_month}", status_list)

    # Apply filters
    filtered_df = df.copy()
    if selected_prov != 'Все':
        filtered_df = filtered_df[filtered_df['provider'] == selected_prov]
    if selected_route != 'Все':
        filtered_df = filtered_df[filtered_df['route_group'] == selected_route]
    if selected_format != 'Все':
        filtered_df = filtered_df[filtered_df['format'] == selected_format]

    if selected_status != 'Все' and month_col in filtered_df.columns:
        status_norm = selected_status.strip().lower()
        if status_norm == 'свободно':
            filtered_df = filtered_df[
                filtered_df[month_col].astype(str).str.strip().str.lower().isin(['свободно', 'свободен', 'своб'])]
        elif status_norm in ['забронировано']:
            filtered_df = filtered_df[
                filtered_df[month_col].astype(str).str.strip().str.lower().isin(['забронировано', 'бронь', 'резерв'])]
        elif status_norm in ['занято']:
            filtered_df = filtered_df[filtered_df[month_col].astype(str).str.strip().str.lower().isin(
                ['занято', 'продано', 'продан', 'занят'])]

    # 5. Dashboard
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Всего найдено", len(filtered_df))
    if len(filtered_df) > 0 and month_col in filtered_df.columns:
        free_count = len(filtered_df[filtered_df[month_col].astype(str).str.strip().str.lower().isin(
            ['свободно', 'свободен', 'своб'])])
        avg_price = int(filtered_df['price'].mean())
    else:
        free_count = 0
        avg_price = 0

    m2.metric(f"Свободно ({selected_month})", free_count)
    m3.metric("Средняя аренда", f"{avg_price:,.0f} ₽".replace(",", " "))
    m4.metric(f"Занято ({selected_month})", len(filtered_df) - free_count)

    st.markdown("---")

    # 6. Map
    st.subheader("📍 Карта щитов")
    if 'lat' in filtered_df.columns and not filtered_df.empty:
        map_subset = filtered_df.dropna(subset=['lat', 'lon'])
        if not map_subset.empty:
            st.map(map_subset, latitude='lat', longitude='lon', color='#0f766e', size=25)
        else:
            st.info("У отфильтрованных точек нет координат.")

    st.markdown("---")

    # 7. Table
    st.subheader("📋 Реестр")
    display_cols = ['provider', 'route_group', 'id', 'address', 'format', 'material', 'screens_outputs', 'price',
                    'installation_price',
                    'status_june', 'status_july', 'status_august', 'status_september', 'status_october',
                    'status_november', 'status_december']
    actual_display_cols = [col for col in display_cols if col in filtered_df.columns]

    st.dataframe(
        filtered_df[actual_display_cols],
        use_container_width=True,
        column_config={
            "route_group": st.column_config.TextColumn("Район / Направление"),
            "price": st.column_config.NumberColumn("Аренда/мес, ₽", format="%d ₽"),
            "installation_price": st.column_config.NumberColumn("Доп. Монтаж, ₽", format="%d ₽"),
        }
    )