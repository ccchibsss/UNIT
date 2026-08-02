"""
Marketplace Manager (FBS Edition) — Монолитное приложение Streamlit.
Адаптация: Автозапчасти.
Модель: FBS (Fulfillment by Seller). WB / Ozon / Яндекс Маркет / СберМегаМаркет.

Исправления v4.1:
  1. Удален np.core.defchararray (несовместим с NumPy 2.0+).
  2. Усилен Regex для очистки markdown-ответов LLM.
  3. Исправлена обработка NaN в правилах и агрегации P&L.
  4. Оптимизирована генерация строк (истинная векторизация массивов + zip).
"""
from __future__ import annotations
import json, warnings, re
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

warnings.filterwarnings("ignore")
st.set_page_config(page_title="FBS AutoParts Manager", page_icon="🚗", layout="wide", initial_sidebar_state="expanded")

MARKETPLACES = ["Wildberries", "Ozon", "Яндекс Маркет", "СберМегаМаркет"]
CATEGORIES = ["Двигатель и выхлоп", "Подвеска и рулевое", "Тормозная система", "Фильтры и масла", "Кузовные детали", "Электрика", "Шины и диски"]
STATUSES = ["Активен", "Приостановлен", "Нет в наличии"]
STATUS_ICON = {"Активен": "🟢", "Приостановлен": "🟡", "Нет в наличии": "🔴"}
MP_COLORS = {"Wildberries": "#cb11ab", "Ozon": "#005bff", "Яндекс Маркет": "#fc3f1d", "СберМегаМаркет": "#21a038"}
ABC_COLORS = {"A": "#28a745", "B": "#ffc107", "C": "#dc3545"}
XYZ_COLORS = {"X": "#28a745", "Y": "#ffc107", "Z": "#dc3545"}
SOURCE_LABEL = {"api": "API маркетплейса", "deepseek": "DeepSeek", "hybrid": "Гибрид", "default": "Оценка"}
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"

AUTO_PARTS_TYPES = {
    "Двигатель и выхлоп": ["Свечи зажигания", "Прокладка ГБЦ", "Глушитель", "Ремень ГРМ", "Масляный насос"],
    "Подвеска и рулевое": ["Амортизатор", "Шаровая опора", "Сайлентблок", "Рулевая тяга", "Ступица"],
    "Тормозная система": ["Тормозные колодки", "Тормозной диск", "Суппорт", "Тормозной шланг", "Главный цилиндр"],
    "Фильтры и масла": ["Масляный фильтр", "Воздушный фильтр", "Салонный фильтр", "Моторное масло 5W-40", "Антифриз"],
    "Кузовные детали": ["Бампер передний", "Крыло", "Капот", "Зеркало боковое", "Фара"],
    "Электрика": ["Аккумулятор", "Генератор", "Стартер", "Датчик ABS", "Лампа H7"],
    "Шины и диски": ["Шина летняя 205/55 R16", "Диск литой 16", "Колпак", "Цепи противоскольжения", "Домкрат"],
}
BRANDS = ["Bosch", "KYB", "Sachs", "Mann-Filter", "NGK", "Febi", "Lemforder", "Brembo", "Denso", "Оригинал"]

_CAT_COMMISSION = {"Двигатель и выхлоп": 12, "Подвеска и рулевое": 13, "Тормозная система": 13, "Фильтры и масла": 10, "Кузовные детали": 15, "Электрика": 12, "Шины и диски": 9}
_MP_COMM_MULT = {"Wildberries": 1.0, "Ozon": 0.95, "Яндекс Маркет": 0.90, "СберМегаМаркет": 0.92}
_MP_LOGISTICS_BASE = {"Wildberries": 60, "Ozon": 50, "Яндекс Маркет": 55, "СберМегаМаркет": 58}
_MP_PACKAGING = {"Wildberries": 40, "Ozon": 35, "Яндекс Маркет": 38, "СберМегаМаркет": 40}
_MP_STORAGE = {"Wildberries": 5, "Ozon": 4, "Яндекс Маркет": 4, "СберМегаМаркет": 5}
_MP_SLA = {"Wildberries": 2, "Ozon": 3, "Яндекс Маркет": 3, "СберМегаМаркет": 5}

st.markdown("""<style>
 .main-header{font-size:1.9rem;font-weight:700;color:#1f77b4;margin-bottom:.4rem}
 .metric-card{background:linear-gradient(135deg,#1e3a8a,#3b82f6);padding:1rem;border-radius:.9rem;
   color:#fff;text-align:center;box-shadow:0 4px 6px rgba(0,0,0,.1)}
 .metric-card .v{font-size:1.4rem;font-weight:700} .metric-card .l{font-size:.78rem;opacity:.9}
 .src-badge{display:inline-block;padding:.12rem .5rem;border-radius:9999px;font-size:.7rem;font-weight:600}
 section[data-testid="stSidebar"]{width:280px !important}
</style>""", unsafe_allow_html=True)

def src_badge_html(source):
    color = {"api": "#28a745", "deepseek": "#6f42c1", "hybrid": "#fd7e14", "default": "#6c757d"}.get(source, "#6c757d")
    return f'<span class="src-badge" style="background:{color}22;color:{color}">{SOURCE_LABEL.get(source, source)}</span>'

def fmt_rub(v):
    try: 
        return f"{int(round(float(v))):,} ₽".replace(",", " ")
    except (ValueError, TypeError): 
        return "—"

# ----------------------------- ТАРИФЫ -----------------------------
@st.cache_data(ttl=3600)
def default_tariff_rows():
    rows = []
    for mp in MARKETPLACES:
        for cat in CATEGORIES:
            rows.append({"marketplace": mp, "category": cat,
                         "commission_pct": round(_CAT_COMMISSION[cat]*_MP_COMM_MULT[mp], 2),
                         "logistics_per_kg": float(_MP_LOGISTICS_BASE[mp]),
                         "packaging_cost": float(_MP_PACKAGING[mp]),
                         "storage_cost_per_unit": float(_MP_STORAGE[mp]), "sla_days": int(_MP_SLA[mp]), "source": "default"})
    return pd.DataFrame(rows)

def fetch_deepseek_tariffs(api_key, marketplace):
    sysmsg = "You are an e-commerce tariff analyst for Russian auto parts marketplaces (FBS). Return ONLY valid JSON."
    user = (f'Estimate 2026 FBS tariffs for "{marketplace}" for auto parts categories: {", ".join(CATEGORIES)}. '
            'Return {"tariffs":[{"category":str,"commission_pct":num,"logistics_per_kg":num,'
            '"packaging_cost":num,"storage_cost_per_unit":num,"sla_days":int}]}.')
    try:
        r = requests.post(DEEPSEEK_URL, headers={"Authorization": f"Bearer {api_key}"},
                          json={"model": DEEPSEEK_MODEL, "temperature": 0.2,
                                "messages": [{"role": "system", "content": sysmsg}, {"role": "user", "content": user}],
                                "response_format": {"type": "json_object"}}, timeout=40)
        r.raise_for_status()
        raw_content = r.json()["choices"][0]["message"]["content"]
        
        # ИСПРАВЛЕНИЕ: Надежная очистка от markdown-оберток с учетом отступов и переносов строк
        clean_content = re.sub(r'^\s*```(?:json)?\s*|\s*```\s*$', '', raw_content, flags=re.IGNORECASE | re.DOTALL).strip()
        
        data = json.loads(clean_content)
        lst = data.get("tariffs", [])
        rows = []
        for x in lst:
            cat = str(x.get("category", "")).strip()
            if cat in CATEGORIES:
                rows.append({"marketplace": marketplace, "category": cat, "commission_pct": float(x.get("commission_pct") or 0),
                             "logistics_per_kg": float(x.get("logistics_per_kg") or 0), "packaging_cost": float(x.get("packaging_cost") or 0),
                             "storage_cost_per_unit": float(x.get("storage_cost_per_unit") or 0), "sla_days": int(x.get("sla_days") or 0), "source": "deepseek"})
        if not rows: raise ValueError("DeepSeek вернул пустой список")
        return rows
    except json.JSONDecodeError as e:
        raise ValueError(f"Ошибка парсинга JSON от DeepSeek: {e}")
    except Exception as e:
        raise ValueError(f"Ошибка запроса к DeepSeek: {e}")

def resolve_tariffs(keys, use_api=True, use_deepseek=True):
    tmap = {(r["marketplace"], r["category"]): dict(r) for _, r in default_tariff_rows().iterrows()}
    sources = {mp: "default" for mp in MARKETPLACES}
    notes = {mp: "Оценочный движок (нет ключей)" for mp in MARKETPLACES}
    
    for mp in MARKETPLACES:
        chosen, src, note = None, "default", "Оценочный движок"
        if use_deepseek and keys.get("deepseek"):
            try:
                chosen = fetch_deepseek_tariffs(keys["deepseek"], mp)
                src = "deepseek"
                note = "Оценка DeepSeek (API маркетплейса недоступны)"
            except Exception as e: 
                note = f"Ошибка DeepSeek: {str(e)[:50]}..."
        
        if chosen:
            for row in chosen:
                row = dict(row); row["source"] = src; tmap[(mp, row["category"])] = row
        sources[mp] = src; notes[mp] = note
        
    return pd.DataFrame(list(tmap.values())), sources, notes

# ----------------------------- ДАННЫЕ -----------------------------
@st.cache_data(ttl=7200)
def generate_auto_parts_products(n, seed=42):
    rng = np.random.default_rng(seed); n = int(n)
    mp = rng.choice(MARKETPLACES, size=n); cat = rng.choice(CATEGORIES, size=n)
    
    price = np.clip((200 + np.power(rng.random(n), 1.8)*35000), 200, 50000).round().astype(np.int64)
    cost = (price*rng.uniform(0.35, 0.65, n)).round().astype(np.int64)
    weight_kg = np.clip(rng.exponential(scale=2.5, size=n), 0.1, 50.0).round(2)
    
    total_stock = (np.power(rng.random(n), 1.4)*800).astype(np.int64)
    reserved = (rng.random(n)*np.floor(total_stock*0.25)).astype(np.int64)
    fbs_available = np.clip(total_stock-reserved, 0, None)
    
    rating = np.round(3.8+rng.random(n)*1.2, 1)
    status = rng.choice(STATUSES, size=n, p=[0.75, 0.15, 0.10])
    created = [datetime.now()-timedelta(days=int(d)) for d in rng.integers(1, 500, size=n)]
    
    avg = np.power(rng.random(n), 2.5)*15 + 0.5
    std = avg * (0.3 + rng.random(n)*0.4)
    sales_30d = np.maximum(0, np.round(avg*30 + rng.standard_normal(n)*std*5)).astype(np.int64)
    return_rate = np.clip(rng.normal(loc=0.08, scale=0.04, size=n), 0.01, 0.25)
    
    # ИСПРАВЛЕНИЕ: Векторизованная генерация массивов + безопасная сборка строк (совместимо с NumPy 2.0+)
    part_names = [rng.choice(AUTO_PARTS_TYPES[str(c)]) for c in cat]
    part_brands = rng.choice(BRANDS, size=n)
    oe_numbers = rng.integers(100000, 999999, size=n)
    names = [f"{p} {b} (OE-{o})" for p, b, o in zip(part_names, part_brands, oe_numbers)]
    
    sku = [f"SKU-{100000+i}" for i in range(1, n+1)]
    
    return pd.DataFrame({"sku": sku, "name": names, "marketplace": mp, "category": cat, "price": price, "cost_price": cost,
        "weight_kg": weight_kg, "total_stock": total_stock, "reserved": reserved, "fbs_available": fbs_available, 
        "rating": rating, "avg_daily_sales": np.round(avg, 4), "std_daily_sales": np.round(std, 4), 
        "sales_30d": sales_30d, "return_rate": return_rate, "status": status, "created_at": created})

@st.cache_data(ttl=7200)
def generate_trend(n, seed=42):
    rng = np.random.default_rng(seed+7); base = max(1, n)*30; rows = []
    for d in range(30):
        date = datetime.now()-timedelta(days=29-d)
        wk = 1.15 if date.weekday() >= 5 else 1.0
        rev = base*(1+(29-d)*0.004)*wk*rng.uniform(0.90, 1.10)
        rows.append({"date": date.strftime("%Y-%m-%d"), "revenue": rev, "profit": rev*rng.uniform(0.08, 0.18)})
    return pd.DataFrame(rows)

# ----------------------------- ЭКОНОМИКА -----------------------------
@st.cache_data(ttl=3600)
def compute_economics(df, tariffs):
    m = df.merge(tariffs[["marketplace","category","commission_pct","logistics_per_kg","packaging_cost",
                          "storage_cost_per_unit","sla_days","source"]], on=["marketplace","category"], how="left")
    for c in ["commission_pct","logistics_per_kg","packaging_cost","storage_cost_per_unit","sla_days"]:
        m[c] = m[c].fillna(0)
    
    m["logistics_cost"] = m["logistics_per_kg"] * m["weight_kg"]
    m["commission_amount"] = m["price"]*m["commission_pct"]/100.0
    m["return_cost"] = m["logistics_cost"] * m["return_rate"] * 0.8
    m["profit"] = m["price"] - m["cost_price"] - m["commission_amount"] - m["logistics_cost"] - m["packaging_cost"] - m["return_cost"]
    m["margin_pct"] = np.where(m["price"] > 0, m["profit"]/m["price"]*100.0, 0.0)
    
    m["storage_monthly"] = m["fbs_available"]*m["storage_cost_per_unit"]
    m["revenue"] = m["price"]*m["sales_30d"]
    m["cv"] = np.where(m["avg_daily_sales"] > 0, m["std_daily_sales"]/m["avg_daily_sales"], 0.0)
    m["cover_days"] = np.where(m["avg_daily_sales"] > 0, m["fbs_available"]/m["avg_daily_sales"], np.nan)
    return m

def rule_advice(row):
    margin = float(row.get("margin_pct", 0) or 0)
    cover = row.get("cover_days")
    
    # ИСПРАВЛЕНИЕ: Безопасная обработка NaN для cv и return_rate
    cv_raw = row.get("cv", 0)
    cv = float(cv_raw) if not pd.isna(cv_raw) else 0.0
    
    ret_raw = row.get("return_rate", 0)
    ret_rate = float(ret_raw) if not pd.isna(ret_raw) else 0.0
    
    r = []
    if margin < 0: r.append("Убыток — проверить себестоимость или поднять цену.")
    elif margin < 15: r.append("Низкая маржа для автозапчастей (цель >20%) — пересмотреть закупку.")
    if float(row.get("fbs_available", 0)) < 5: r.append("Критический остаток — срочно пополнить.")
    elif cover is not None and not pd.isna(cover) and cover < 14: r.append(f"Низкое покрытие (~{cover:.0f} дн.) — оформить поставку.")
    elif cover is not None and not pd.isna(cover) and cover > 90: r.append(f"Затоваривание (~{cover:.0f} дн.) — распродажа или вывод.")
    if ret_rate > 0.15: r.append(f"Высокий возврат ({ret_rate*100:.0f}%) — проверить совместимость по VIN и описание.")
    if cv > 0.4: r.append(f"Нестабильный спрос (Z, CV={cv*100:.0f}%) — держать минимальный страховой запас.")
    return "; ".join(r or ["Показатели в норме — удерживать стратегию"])

# ----------------------------- SESSION STATE -----------------------------
def init_state():
    ss = st.session_state
    ss.setdefault("n_products", 50000)
    ss.setdefault("products_df", generate_auto_parts_products(ss["n_products"]))
    ss.setdefault("trend_df", generate_trend(ss["n_products"]))
    ss.setdefault("tariffs_df", default_tariff_rows())
    ss.setdefault("tariff_sources", {mp: "default" for mp in MARKETPLACES})
    ss.setdefault("api_keys", {})
    ss.setdefault("critical_stock", 5)

init_state()

# ----------------------------- UI -----------------------------
with st.sidebar:
    st.title("🚗 FBS AutoParts"); st.caption("Специфика: вес, возвраты, VIN"); st.markdown("---")
    page = st.radio("Навигация", ["🏠 Дашборд","📦 Товары и остатки","💰 Ценообразование","📊 ABC/XYZ","📈 Отчёты P&L","⚙️ Настройки"], label_visibility="collapsed")
    st.markdown("---"); st.subheader("🔍 Фильтры")
    sel_mp = st.multiselect("Маркетплейс", MARKETPLACES, default=MARKETPLACES[:2])
    sel_cat = st.multiselect("Категория", CATEGORIES, default=[])
    st.markdown("---"); st.markdown("**Источники тарифов:**")
    for mp in MARKETPLACES:
        st.markdown(f"{mp}: {src_badge_html(st.session_state.tariff_sources.get(mp,'default'))}", unsafe_allow_html=True)
    st.markdown("---"); st.caption(f"© {datetime.now().year} FBS AutoParts v4.1")

ECO = compute_economics(st.session_state.products_df, st.session_state.tariffs_df)
df = ECO
if sel_mp: df = df[df["marketplace"].isin(sel_mp)]
if sel_cat: df = df[df["category"].isin(sel_cat)]
critical = st.session_state.critical_stock
all_default = all(st.session_state.tariff_sources.get(mp) == "default" for mp in MARKETPLACES)

if page == "🏠 Дашборд":
    if all_default:
        st.warning("⚠️ **Демо-тарифы.** Подключите API-ключи или DeepSeek в ⚙️ Настройках для актуальных данных.")
    
    st.markdown('<div class="main-header">📊 Обзор FBS-операций (Автозапчасти)</div>', unsafe_allow_html=True)
    revenue = float((df["price"]*df["sales_30d"]).sum())
    gross_profit = float((df["profit"]*df["sales_30d"]).sum())
    storage_total = float(df["storage_monthly"].sum())
    net_profit = gross_profit - storage_total
    low_stock = int((df["fbs_available"] < critical).sum())
    avg_margin = float(df.loc[df["price"] > 0, "margin_pct"].mean()) if len(df) else 0.0
    margin_pct = (net_profit/revenue*100) if revenue > 0 else 0
    
    c = st.columns(5)
    c[0].markdown(f'<div class="metric-card"><div class="l">💰 Выручка 30д</div><div class="v">{fmt_rub(revenue)}</div></div>', unsafe_allow_html=True)
    c[1].markdown(f'<div class="metric-card"><div class="l">📈 Чистая прибыль 30д</div><div class="v">{fmt_rub(net_profit)}</div><div class="l">{margin_pct:.1f}% маржа</div></div>', unsafe_allow_html=True)
    c[2].markdown(f'<div class="metric-card"><div class="l">🏚️ Хранение FBS/мес</div><div class="v">{fmt_rub(storage_total)}</div></div>', unsafe_allow_html=True)
    c[3].markdown(f'<div class="metric-card"><div class="l">⚠️ Крит. остаток</div><div class="v">{low_stock} SKU</div></div>', unsafe_allow_html=True)
    c[4].markdown(f'<div class="metric-card"><div class="l">🎯 Ср. маржа</div><div class="v">{avg_margin:.1f}%</div></div>', unsafe_allow_html=True)
    
    st.markdown("---"); L, R = st.columns(2)
    with L:
        st.subheader("📈 Выручка по маркетплейсам")
        rev_mp = df.assign(_r=df["price"]*df["sales_30d"]).groupby("marketplace", as_index=False)["_r"].sum().rename(columns={"_r":"revenue"})
        fig = px.pie(rev_mp, values="revenue", names="marketplace", hole=0.5, color="marketplace", color_discrete_map=MP_COLORS)
        fig.update_traces(textinfo="percent+label"); fig.update_layout(height=320, margin=dict(t=0,b=0,l=0,r=0)); st.plotly_chart(fig, use_container_width=True)
    with R:
        st.subheader("📦 Структура FBS-остатков")
        stock = pd.DataFrame({"Тип": ["Доступно FBS","В резерве"], "Единиц": [int(df["fbs_available"].sum()), int(df["reserved"].sum())]})
        fig = px.bar(stock, x="Тип", y="Единиц", color="Тип", color_discrete_sequence=["#28a745","#ffc107"], text="Единиц")
        fig.update_layout(showlegend=False, height=320, margin=dict(t=0,b=0,l=0,r=0)); st.plotly_chart(fig, use_container_width=True)

elif page == "📦 Товары и остатки":
    st.markdown('<div class="main-header">📦 Управление FBS-остатками (Автозапчасти)</div>', unsafe_allow_html=True)
    st.caption(f"В выборке **{len(df):,} SKU** (всего {len(ECO):,})".replace(",", " "))
    view = df[["sku","name","marketplace","category","price","cost_price","weight_kg","fbs_available","sales_30d","return_rate","profit","margin_pct","status"]].copy()
    view["Цена"] = view["price"].map(fmt_rub); view["Прибыль/ед"] = view["profit"].map(fmt_rub)
    view["Маржа %"] = view["margin_pct"].round(1); view["Возвраты %"] = (view["return_rate"]*100).round(1).astype(str)+"%"
    view["Статус"] = view["status"].map(lambda s: STATUS_ICON.get(s,"⚪")+" "+s)
    show = view.rename(columns={"sku":"SKU","name":"Название (Бренд, OE)","marketplace":"МП","category":"Категория","weight_kg":"Вес (кг)","fbs_available":"Доступно FBS","sales_30d":"Продажи 30д"})
    st.dataframe(show[["SKU","Название (Бренд, OE)","МП","Категория","Цена","Вес (кг)","Прибыль/ед","Маржа %","Доступно FBS","Продажи 30д","Возвраты %","Статус"]], use_container_width=True, hide_index=True, height=480)
    
    st.markdown("---"); st.subheader("🧮 Калькулятор юнит-экономики (с учетом веса и возвратов)")
    a,b,c,d,e = st.columns(5)
    price_in = a.number_input("Цена, ₽", value=2500, step=50)
    cost_in = b.number_input("Себестоимость, ₽", value=1200, step=50)
    weight_in = c.number_input("Вес, кг", value=2.5, step=0.1)
    mp_in = d.selectbox("Маркетплейс", MARKETPLACES, index=0)
    cat_in = e.selectbox("Категория", CATEGORIES, index=0)
    ret_in = st.number_input("Прогноз возвратов, %", value=10, step=1, min_value=0, max_value=50)
    
    tr = st.session_state.tariffs_df[(st.session_state.tariffs_df["marketplace"]==mp_in)&(st.session_state.tariffs_df["category"]==cat_in)]
    if len(tr):
        tr = tr.iloc[0]
        log_cost = tr["logistics_per_kg"] * weight_in
        comm = price_in * tr["commission_pct"] / 100.0
        return_loss = log_cost * (ret_in / 100.0) * 0.8
        profit = price_in - cost_in - comm - log_cost - tr["packaging_cost"] - return_loss
        margin = profit / price_in * 100 if price_in > 0 else 0
        denom = 1 - tr["commission_pct"]/100.0
        be = (cost_in + log_cost + tr["packaging_cost"] + return_loss) / denom if denom > 0 else 0
        
        r = st.columns(4)
        r[0].metric("Комиссия", fmt_rub(comm), f"{tr['commission_pct']:.1f}%")
        r[1].metric("Логистика (с возвратом)", fmt_rub(log_cost + return_loss), f"Вес: {weight_in} кг")
        r[2].metric("Прибыль с единицы", fmt_rub(profit), delta=f"{margin:.1f}% маржа", delta_color="normal" if profit>0 else "inverse")
        r[3].metric("Безубыточная цена", fmt_rub(be))
        st.caption(f"Источник: {src_badge_html(tr['source'])} · хранение {tr['storage_cost_per_unit']:.0f} ₽/ед/мес", unsafe_allow_html=True)

elif page == "💰 Ценообразование":
    st.markdown('<div class="main-header">💰 Умное ценообразование</div>', unsafe_allow_html=True)
    st.info("ℹ️ Учитывайте, что снижение цены на низколиквидные автозапчасти редко дает кратный рост спроса (низкая эластичность).")
    s1,s2,s3 = st.columns(3)
    discount = s1.slider("Скидка, %", 0, 50, 5)
    elasticity = s2.number_input("Эластичность", value=0.8, step=0.1, min_value=0.0, max_value=5.0)
    only_prof = s3.checkbox("Только прибыльные", value=True)
    
    d = df.copy()
    if only_prof: d = d[d["profit"] > 0]
    new_price = np.round(d["price"]*(1-discount/100.0)/10)*10
    mult = 1+elasticity*discount/100.0; new_sales = d["sales_30d"]*mult
    new_comm = new_price*d["commission_pct"]/100.0
    new_unit = new_price - d["cost_price"] - new_comm - (d["logistics_per_kg"]*d["weight_kg"]) - d["packaging_cost"] - (d["logistics_per_kg"]*d["weight_kg"]*d["return_rate"]*0.8)
    
    old_rev = float((d["price"]*d["sales_30d"]).sum()); new_rev = float((new_price*new_sales).sum())
    old_prof = float((d["profit"]*d["sales_30d"]).sum()); new_prof = float((new_unit*new_sales).sum())
    
    m = st.columns(3)
    m[0].metric("Текущая прибыль", fmt_rub(old_prof))
    
    # ИСПРАВЛЕНИЕ: Защита от деления на ноль при old_rev == 0
    rev_delta = f"{(new_rev/old_rev-1)*100:+.1f}%" if old_rev > 0 else "0%"
    prof_delta = f"{(new_prof/old_prof-1)*100:+.1f}%" if old_prof > 0 else "0%"
    
    m[1].metric("Прогноз выручки", fmt_rub(new_rev), delta=rev_delta)
    m[2].metric("Прогноз прибыли", fmt_rub(new_prof), delta=prof_delta, delta_color="normal" if new_prof>=old_prof else "inverse")
    
    samp = d.assign(new_price=new_price, new_unit=new_unit).nlargest(15, "revenue")
    disp = samp[["sku","name","price","new_price","profit","new_unit","margin_pct","sales_30d"]].copy()
    for col in ["price","new_price","profit","new_unit"]: disp[col] = disp[col].map(fmt_rub)
    st.dataframe(disp.rename(columns={"sku":"SKU","name":"Товар","price":"Тек. цена","new_price":"Новая цена","profit":"Тек. прибыль/ед","new_unit":"Новая прибыль/ед","margin_pct":"Маржа %","sales_30d":"Продажи 30д"}), use_container_width=True, hide_index=True)

elif page == "📊 ABC/XYZ":
    st.markdown('<div class="main-header">📊 ABC/XYZ анализ</div>', unsafe_allow_html=True)
    a = df.assign(rev=df["price"]*df["sales_30d"]).sort_values("rev", ascending=False).copy()
    total = float(a["rev"].sum()); a["pct"] = (a["rev"].cumsum()/total*100) if total>0 else 0.0
    a["abc"] = np.where(a["pct"]<=80,"A",np.where(a["pct"]<=95,"B","C"))
    a["xyz"] = np.where(a["cv"]<=0.15,"X",np.where(a["cv"]<=0.35,"Y","Z"))
    L,R = st.columns(2)
    with L:
        cnt = a["abc"].value_counts().reindex(["A","B","C"], fill_value=0)
        fig = px.pie(names=cnt.index, values=cnt.values, color=cnt.index, color_discrete_map=ABC_COLORS, title="ABC")
        fig.update_layout(height=320); st.plotly_chart(fig, use_container_width=True)
    with R:
        cnt = a["xyz"].value_counts().reindex(["X","Y","Z"], fill_value=0)
        fig = px.pie(names=cnt.index, values=cnt.values, color=cnt.index, color_discrete_map=XYZ_COLORS, title="XYZ")
        fig.update_layout(height=320); st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    st.markdown("**Специфика автозапчастей:** AX/BX — ходовые расходники (фильтры, колодки), держать макс. запас. CZ — неликвид (редкие кузовные детали), только под заказ или FBO.")
  
elif page == "📈 Отчёты P&L":
    st.markdown('<div class="main-header">📈 Финансовый отчёт (P&L)</div>', unsafe_allow_html=True)
    d = df.copy()
    d["_cogs"]=d["cost_price"]*d["sales_30d"]
    d["_comm"]=d["price"]*d["commission_pct"]/100*d["sales_30d"]
    d["_log"]=(d["logistics_per_kg"]*d["weight_kg"])*d["sales_30d"]
    d["_ret_loss"]=(d["logistics_per_kg"]*d["weight_kg"]*d["return_rate"]*0.8)*d["sales_30d"]
    d["_pack"]=d["packaging_cost"]*d["sales_30d"]; d["_stor"]=d["storage_monthly"]
    
    pl = d.groupby("marketplace").agg(Выручка=("revenue","sum"), Себестоимость=("_cogs","sum"), Комиссия=("_comm","sum"),
                                      Логистика=("_log","sum"), Потери_возвраты=("_ret_loss","sum"), Упаковка=("_pack","sum"), Хранение=("_stor","sum"))
    pl["Валовая прибыль"] = pl["Выручка"]-pl[["Себестоимость","Комиссия","Логистика","Потери_возвраты","Упаковка","Хранение"]].sum(axis=1)
    pl["Маржа %"] = (pl["Валовая прибыль"]/pl["Выручка"]*100).round(1)
    
    tot = pl.sum(numeric_only=True)
    tot["Валовая прибыль"] = tot["Выручка"] - tot[["Себестоимость","Комиссия","Логистика","Потери_возвраты","Упаковка","Хранение"]].sum()
    tot["Маржа %"] = round(tot["Валовая прибыль"]/tot["Выручка"]*100, 1) if tot["Выручка"] else 0.0
    
    # ИСПРАВЛЕНИЕ: Безопасное добавление строки "ИТОГО" через concat
    tot.name = "ИТОГО"
    pl = pd.concat([pl, tot.to_frame().T])
    
    pv = pl.copy()
    for col in ["Выручка","Себестоимость","Комиссия","Логистика","Потери_возвраты","Упаковка","Хранение","Валовая прибыль"]: 
        pv[col] = pv[col].map(fmt_rub)
    st.dataframe(pv, use_container_width=True)

elif page == "⚙️ Настройки":
    st.markdown('<div class="main-header">⚙️ Настройки FBS</div>', unsafe_allow_html=True)
    keys = st.session_state.api_keys
    n1,n2 = st.columns(2)
    with n1:
        st.subheader("🔑 API-ключи")
        keys["wb"]=st.text_input("Wildberries API Key", value=keys.get("wb",""), type="password")
        keys["ozon_client"]=st.text_input("Ozon Client ID", value=keys.get("ozon_client",""))
        keys["ozon_key"]=st.text_input("Ozon API Key", value=keys.get("ozon_key",""), type="password")
        keys["deepseek"]=st.text_input("DeepSeek API Key", value=keys.get("deepseek",""), type="password")
    with n2:
        st.subheader("📦 Параметры")
        st.session_state.critical_stock=st.number_input("Крит. остаток, ед.", value=int(st.session_state.critical_stock), step=1, min_value=0)
    
    st.markdown("---"); st.subheader("🔄 Синхронизация тарифов")
    if st.button("🔁 Синхронизировать тарифы (DeepSeek)"):
        with st.spinner("Запрос к DeepSeek..."):
            tf, src, notes = resolve_tariffs(keys, False, True)
            st.session_state.tariffs_df=tf; st.session_state.tariff_sources=src; st.session_state.tariff_notes=notes
        st.success("Готово."); st.rerun()
    if "tariff_notes" in st.session_state:
        for mp in MARKETPLACES:
            st.markdown(f"{mp}: {src_badge_html(st.session_state.tariff_sources.get(mp))} — {st.session_state.tariff_notes.get(mp,'')}", unsafe_allow_html=True)
            
    st.markdown("---"); st.subheader("🗑️ Данные")
    q,w,_=st.columns([1,1,2])
    new_n=q.number_input("Кол-во товаров", value=int(st.session_state.n_products), step=50000, min_value=1000, max_value=1_000_000)
    if w.button("🔄 Сгенерировать заново"):
        with st.spinner(f"Генерируем {int(new_n):,} SKU…".replace(",", " ")):
            st.session_state.n_products=int(new_n)
            generate_auto_parts_products.clear()
            compute_economics.clear()
            st.session_state.products_df=generate_auto_parts_products(int(new_n))
            st.session_state.trend_df=generate_trend(int(new_n))
        st.success("Готово."); st.rerun()

st.markdown("---")
st.caption("🚗 FBS AutoParts Manager v4.1 · кэширование · учет веса и возвратов · 500 000+ SKU")
