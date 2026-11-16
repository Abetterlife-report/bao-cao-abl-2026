import os
import sys
import subprocess
import importlib.util

# 0. Tự động kiểm tra & cài đặt các thư viện cần thiết cho ĐÚNG bản Python đang chạy
REQUIRED_PACKAGES = {
    "streamlit": "streamlit",
    "pandas": "pandas",
    "plotly": "plotly",
    "openpyxl": "openpyxl",
    "xlsxwriter": "xlsxwriter",
    "pptx": "python-pptx",
    "matplotlib": "matplotlib",
}

missing_pkgs = [
    pip_name for mod_name, pip_name in REQUIRED_PACKAGES.items()
    if importlib.util.find_spec(mod_name) is None
]

if missing_pkgs:
    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_pkgs])

import io
import re
import glob
import hashlib
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    HAS_MPL = True
except ImportError:
    HAS_MPL = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# 1. Cấu hình trang
st.set_page_config(page_title="Báo Cáo Bán Hàng ABL 2026 (v17)", page_icon="📊", layout="wide")
st.title("📊 Dashboard Phân Tích Bán Hàng & Cơ Cấu Nhóm Hàng - ABL")

FOLDER_PATH = r"D:\ABL\03_Báo cáo\01_Doanh số\Data\Data_nhóm hàng"

excel_files = sorted([
    f for f in glob.glob(os.path.join(FOLDER_PATH, "*.xls*"))
    if not os.path.basename(f).startswith("~$")
])

if not excel_files:
    st.error(f"Không tìm thấy file Excel nào trong thư mục: `{FOLDER_PATH}`")
    st.stop()

NEEDED_COLS = {
    "Ngày", "Mã đơn hàng eShop", "Số hóa đơn", "Mã hàng hóa", "Tên hàng hóa",
    "Đơn vị tính", "Số lượng", "Đơn giá trước thuế", "Doanh thu",
    "Thuế GTGT", "Giá xuất kho", "Lợi nhuận", "Nhóm1", "Nhóm2", "Nhóm 1", "Nhóm 2"
}

NUM_COLS = [
    "Số lượng", "Đơn giá trước thuế", "Doanh thu",
    "Thuế GTGT", "Giá xuất kho", "Lợi nhuận"
]

PALETTE_COLORS = [
    "#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD",
    "#8C564B", "#E377C2", "#17BECF", "#BCBD22", "#393B79",
    "#E6550D", "#31A354", "#756BB1", "#636363", "#AD494A"
]

def parse_vn_date(series, fallback_date=pd.NaT):
    if pd.api.types.is_datetime64_any_dtype(series):
        res = pd.to_datetime(series).dt.normalize()
    else:
        s_str = series.astype(str).str.strip()
        res = pd.to_datetime(s_str, format="%d/%m/%Y", errors="coerce")
        mask_na = res.isna() & series.notna()
        if mask_na.any():
            res.loc[mask_na] = pd.to_datetime(s_str[mask_na], dayfirst=True, errors="coerce")
        res = res.dt.normalize()
    if pd.notna(fallback_date):
        res = res.fillna(fallback_date)
    return res

def build_complete_daily_series(df_source, master_days_df, group_col, selected_items, value_col):
    """
    Ép dữ liệu theo đúng khung thời gian chuẩn master_days_df của Dữ liệu đầu vào.
    Ngày nào không phát sinh bán sẽ tự động điền 0 để mọi biểu đồ khớp 100% trục thời gian.
    """
    if master_days_df.empty or not selected_items:
        return pd.DataFrame(columns=["Ngày_dt", "Ngày hiển thị", group_col, value_col])
    sub = df_source[df_source[group_col].isin(selected_items)]
    grouped = sub.groupby(["Ngày_dt", "Ngày hiển thị", group_col], as_index=False)[value_col].sum()
    grid = master_days_df.assign(key=1).merge(
        pd.DataFrame({group_col: selected_items, "key": 1}), on="key"
    ).drop("key", axis=1)
    complete = (
        grid.merge(grouped, on=["Ngày_dt", "Ngày hiển thị", group_col], how="left")
        .fillna({value_col: 0})
        .sort_values(["Ngày_dt", group_col])
    )
    return complete

def build_line_chart_with_sum(df_line, x_col, y_col, group_col, ordered_items, master_ordered_days, title, legend_title, line_width=2.5, marker_size=7):
    """
    Vẽ biểu đồ đường đồng bộ trục thời gian chuẩn master_ordered_days,
    có dòng '∑ TỔNG NGÀY' trong hộp Hover và tổng kỳ trên Legend.
    """
    val_fmt = ",.2f" if y_col == "Số lượng" else ",.0f"
    fig = go.Figure()

    df_sum_daily = (
        df_line.groupby(["Ngày_dt", x_col], as_index=False)[y_col]
        .sum()
        .sort_values("Ngày_dt")
    )
    grand_total = df_sum_daily[y_col].sum()

    fig.add_trace(go.Scatter(
        x=df_sum_daily[x_col],
        y=df_sum_daily[y_col],
        mode="lines+markers",
        name=f"∑ TỔNG ({grand_total:{val_fmt}})",
        line=dict(color="#111827", width=2.2, dash="dot"),
        marker=dict(size=6, symbol="diamond", color="#111827"),
        hovertemplate=f"<b>∑ TỔNG NGÀY: %{{y:{val_fmt}}}</b><extra></extra>"
    ))

    for idx, item in enumerate(ordered_items):
        sub = df_line[df_line[group_col] == item].sort_values("Ngày_dt")
        if sub.empty:
            continue
        item_total = sub[y_col].sum()
        color = PALETTE_COLORS[idx % len(PALETTE_COLORS)]
        short_name = item if len(str(item)) <= 32 else str(item)[:30] + "…"
        fig.add_trace(go.Scatter(
            x=sub[x_col],
            y=sub[y_col],
            mode="lines+markers",
            name=f"{short_name} [{item_total:{val_fmt}}]",
            line=dict(color=color, width=line_width),
            marker=dict(size=marker_size, color=color),
            hovertemplate=f"{short_name}: <b>%{{y:{val_fmt}}}</b><extra></extra>"
        ))

    fig.update_layout(
        title=title,
        height=430,
        xaxis_title="Ngày bán",
        yaxis_title=y_col,
        legend_title=f"{legend_title} [Tổng kỳ]",
        hovermode="x unified",
        margin=dict(t=55, l=45, r=20, b=45)
    )
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=master_ordered_days)
    return fig

@st.cache_data(show_spinner=False)
def load_single_file(file_path, file_mtime):
    file_name = os.path.basename(file_path)
    try:
        try:
            df_temp = pd.read_excel(
                file_path,
                sheet_name="SỔ CHI TIẾT BÁN HÀNG",
                skiprows=3,
                usecols=lambda c: str(c).strip() in NEEDED_COLS
            )
        except ValueError:
            df_temp = pd.read_excel(
                file_path,
                sheet_name=0,
                skiprows=3,
                usecols=lambda c: str(c).strip() in NEEDED_COLS
            )

        df_temp.columns = [str(c).strip() for c in df_temp.columns]
        if "Nhóm 1" in df_temp.columns and "Nhóm1" not in df_temp.columns:
            df_temp.rename(columns={"Nhóm 1": "Nhóm1"}, inplace=True)
        if "Nhóm 2" in df_temp.columns and "Nhóm2" not in df_temp.columns:
            df_temp.rename(columns={"Nhóm 2": "Nhóm2"}, inplace=True)

        if "Mã hàng hóa" not in df_temp.columns or "Doanh thu" not in df_temp.columns:
            return None, file_name

        match_date = re.search(r"(\d{2})[./-](\d{2})[./-](\d{4})", file_name)
        file_date = (
            pd.to_datetime(f"{match_date.group(1)}/{match_date.group(2)}/{match_date.group(3)}", format="%d/%m/%Y", errors="coerce")
            if match_date else pd.NaT
        )

        df_temp = df_temp[df_temp["Mã hàng hóa"].notna()].copy()
        df_temp = df_temp[~df_temp["Mã hàng hóa"].astype(str).str.contains("Tổng|Total", case=False, na=False)]

        df_temp["Ngày_dt"] = parse_vn_date(df_temp["Ngày"], fallback_date=file_date)
        df_temp["Ngày hiển thị"] = df_temp["Ngày_dt"].dt.strftime("%d/%m/%Y").fillna("Không rõ ngày")

        df_temp["Đơn vị tính"] = df_temp["Đơn vị tính"].fillna("(Chưa có ĐVT)").astype(str).str.strip()
        df_temp.loc[df_temp["Đơn vị tính"] == "", "Đơn vị tính"] = "(Chưa có ĐVT)"

        for grp_col, default_lbl in [("Nhóm2", "(Chưa phân Nhóm 2)"), ("Nhóm1", "(Chưa phân Nhóm 1)")]:
            if grp_col not in df_temp.columns:
                df_temp[grp_col] = default_lbl
            else:
                df_temp[grp_col] = df_temp[grp_col].fillna(default_lbl).astype(str).str.strip()
                df_temp.loc[df_temp[grp_col].isin(["", "nan", "None", "0"]), grp_col] = default_lbl

        for str_col in ["Mã đơn hàng eShop", "Số hóa đơn", "Mã hàng hóa", "Tên hàng hóa"]:
            if str_col in df_temp.columns:
                df_temp[str_col] = (
                    df_temp[str_col]
                    .fillna("")
                    .astype(str)
                    .str.replace(r"\.0$", "", regex=True)
                    .str.strip()
                )

        df_temp["SKU_Label"] = "[" + df_temp["Mã hàng hóa"] + "] " + df_temp["Tên hàng hóa"]

        for col in NUM_COLS:
            if col in df_temp.columns:
                df_temp[col] = pd.to_numeric(df_temp[col], errors="coerce").fillna(0)

        df_temp["Nguồn file"] = os.path.splitext(file_name)[0]
        return df_temp, None
    except Exception as e:
        return None, f"{file_name} ({e})"

# --- THANH BỘ LỌC BÊN TRÁI (SIDEBAR) ---
st.sidebar.header("📂 Dữ liệu đầu vào")
selected_files = st.sidebar.multiselect(
    "Chọn các file báo cáo ngày:",
    options=excel_files,
    default=excel_files,
    format_func=lambda x: os.path.basename(x)
)

if not selected_files:
    st.warning("⚠️ Vui lòng chọn ít nhất 1 file báo cáo ở thanh bên trái.")
    st.stop()

all_dfs = []
skipped = []
for f_path in selected_files:
    mtime = os.path.getmtime(f_path)
    df_part, err = load_single_file(f_path, mtime)
    if df_part is not None:
        all_dfs.append(df_part)
    elif err:
        skipped.append(err)

if skipped:
    st.sidebar.info(f"Đã bỏ qua {len(skipped)} file khác form: {', '.join(skipped)}")

if not all_dfs:
    st.error("Không tìm thấy dữ liệu hợp lệ theo form Sổ Chi Tiết Bán Hàng.")
    st.stop()

df_raw_input = pd.concat(all_dfs, ignore_index=True)

st.sidebar.markdown("---")
st.sidebar.header("🔍 Bộ lọc phân tích")

# 1. Lọc theo khoảng ngày — tự động đồng bộ lại khi danh sách file đầu vào thay đổi
files_sig = hashlib.md5("|".join(sorted(selected_files)).encode()).hexdigest()[:10]
valid_dates = df_raw_input["Ngày_dt"].dropna()
df = df_raw_input.copy()

if not valid_dates.empty:
    min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
    if min_date < max_date:
        date_range = st.sidebar.date_input(
            "Chọn khoảng ngày:",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key=f"date_range_{files_sig}"
        )
        if isinstance(date_range, tuple) and len(date_range) == 2:
            df = df[(df["Ngày_dt"].dt.date >= date_range[0]) & (df["Ngày_dt"].dt.date <= date_range[1])]
    else:
        st.sidebar.info(f"📅 Dữ liệu 1 ngày: **{min_date.strftime('%d/%m/%Y')}**")

# Tạo KHUNG THỜI GIAN CHUẨN CHUNG (MASTER_DAYS_DF) từ Dữ liệu đầu vào + Khoảng ngày đã chọn
MASTER_DAYS_DF = (
    df[["Ngày_dt", "Ngày hiển thị"]]
    .dropna(subset=["Ngày_dt"])
    .drop_duplicates()
    .sort_values("Ngày_dt")
    .reset_index(drop=True)
)
MASTER_ORDERED_DAYS = MASTER_DAYS_DF["Ngày hiển thị"].tolist()

if MASTER_ORDERED_DAYS:
    st.sidebar.caption(f"🕒 Mốc thời gian đồng bộ: **{MASTER_ORDERED_DAYS[0]}** → **{MASTER_ORDERED_DAYS[-1]}** ({len(MASTER_ORDERED_DAYS)} ngày)")

# 2. Bộ lọc phân cấp Nhóm 2 (Cha) -> Nhóm 1 (Con)
all_nhom2 = sorted(df["Nhóm2"].unique().tolist())
selected_nhom2 = st.sidebar.multiselect(
    "🗂️ Nhóm 2 (Ngành hàng cha):",
    options=all_nhom2,
    default=all_nhom2
)
df_after_n2 = df[df["Nhóm2"].isin(selected_nhom2)]

available_nhom1 = sorted(df_after_n2["Nhóm1"].unique().tolist())
selected_nhom1 = st.sidebar.multiselect(
    "📁 Nhóm 1 (Nhóm hàng con):",
    options=available_nhom1,
    default=available_nhom1
)

# 3. Lọc ĐVT & Từ khóa
units = sorted(df_after_n2["Đơn vị tính"].unique().tolist())
selected_units = st.sidebar.multiselect("Đơn vị tính:", options=units, default=units)
keyword = st.sidebar.text_input("Tìm tên hàng / mã hàng / mã đơn:", "")
top_n = st.sidebar.slider("Số lượng Top hiển thị:", min_value=5, max_value=30, value=10, step=5)

df_filtered = df_after_n2[
    df_after_n2["Nhóm1"].isin(selected_nhom1) &
    df_after_n2["Đơn vị tính"].isin(selected_units)
].copy()

if keyword.strip():
    kw = keyword.strip()
    df_filtered = df_filtered[
        df_filtered["Tên hàng hóa"].str.contains(kw, case=False, na=False) |
        df_filtered["Mã hàng hóa"].str.contains(kw, case=False, na=False) |
        df_filtered["Mã đơn hàng eShop"].str.contains(kw, case=False, na=False)
    ]

if df_filtered.empty:
    st.warning("Không có giao dịch nào khớp với bộ lọc hiện tại.")
    st.stop()

# --- TÍNH TOÁN CHỈ SỐ & BẢNG TỔNG HỢP ---
total_rev = df_filtered["Doanh thu"].sum()
total_profit = df_filtered["Lợi nhuận"].sum()
margin_pct = (total_profit / total_rev * 100) if total_rev > 0 else 0
total_orders = df_filtered["Mã đơn hàng eShop"].nunique()
aov = total_rev / total_orders if total_orders > 0 else 0
total_qty = df_filtered["Số lượng"].sum()

# Đồng bộ df_daily với MASTER_DAYS_DF để không bị thiếu ngày nào
df_daily_raw = df_filtered.groupby(["Ngày_dt", "Ngày hiển thị"], as_index=False).agg(
    Doanh_thu=("Doanh thu", "sum"),
    Loi_nhuan=("Lợi nhuận", "sum"),
    Gia_von=("Giá xuất kho", "sum"),
    So_don=("Mã đơn hàng eShop", "nunique"),
    So_luong=("Số lượng", "sum"),
    So_dong=("Mã hàng hóa", "count")
)
df_daily = (
    MASTER_DAYS_DF.merge(df_daily_raw, on=["Ngày_dt", "Ngày hiển thị"], how="left")
    .fillna({"Doanh_thu": 0, "Loi_nhuan": 0, "Gia_von": 0, "So_don": 0, "So_luong": 0, "So_dong": 0})
    .sort_values("Ngày_dt")
)
df_daily["So_don"] = df_daily["So_don"].astype(int)
df_daily["So_dong"] = df_daily["So_dong"].astype(int)

df_prod = df_filtered.groupby(["Nhóm2", "Nhóm1", "Mã hàng hóa", "Tên hàng hóa", "Đơn vị tính"], as_index=False).agg(
    So_luong=("Số lượng", "sum"),
    Doanh_thu=("Doanh thu", "sum"),
    Gia_von=("Giá xuất kho", "sum"),
    Loi_nhuan=("Lợi nhuận", "sum"),
    So_don=("Mã đơn hàng eShop", "nunique")
)

df_loss = df_filtered[df_filtered["Lợi nhuận"] < 0].sort_values("Lợi nhuận")
df_zero_cost = df_filtered[(df_filtered["Giá xuất kho"] == 0) & (df_filtered["Doanh thu"] > 0)]

# --- PHẦN 1: 5 THẺ CHỈ SỐ KPI ---
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("💰 Doanh thu (trước thuế)", f"{total_rev:,.0f} đ")
k2.metric("📈 Lợi nhuận gộp", f"{total_profit:,.0f} đ", delta=f"Biên LN: {margin_pct:.1f}%")
k3.metric("🧾 Tổng số đơn hàng", f"{total_orders:,} đơn")
k4.metric("🛒 Trung bình / Đơn (AOV)", f"{aov:,.0f} đ")
k5.metric("📦 Tổng SL bán ra", f"{total_qty:,.2f}")

st.markdown("---")

# --- PHẦN 2: 7 TAB NGHIỆP VỤ ---
tab_grp, tab_n1_sku, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🗂️ Phân tích nhóm hàng (N2 → N1)",
    "📦 Phân tích nhóm hàng (N1 → SKU)",
    "📅 Xu hướng theo Ngày",
    "📈 Biến động từng SKU",
    "🏆 Top Sản Phẩm",
    "⚠️ Cảnh báo Lỗ & Giá vốn",
    "📋 Sổ Chi Tiết"
])

# ==============================================================================
# TAB 1: PHÂN TÍCH PHÂN CẤP NHÓM 2 (CHA) -> NHÓM 1 (CON) + 2 ĐỒ THỊ ĐƯỜNG
# ==============================================================================
with tab_grp:
    st.subheader("🗂️ Cơ Cấu & Xu Hướng Bán Hàng: Nhóm 2 (Cha) → Nhóm 1 (Con)")

    c_grp1, c_grp2 = st.columns([1.2, 1.5])
    with c_grp1:
        grp_metric_lbl = st.radio(
            "Chọn chỉ số phân tích nhóm hàng:",
            ["Doanh thu", "Lợi nhuận", "Số lượng"],
            horizontal=True,
            key="grp_metric"
        )
    with c_grp2:
        hide_old_unassigned = st.checkbox(
            "Ẩn dữ liệu '(Chưa phân Nhóm)' từ các file cũ để xem rõ Nhóm 2 → Nhóm 1",
            value=True,
            key="hide_old_n2"
        )

    grp_col_map = {"Doanh thu": "Doanh_thu", "Lợi nhuận": "Loi_nhuan", "Số lượng": "So_luong"}
    grp_col = grp_col_map[grp_metric_lbl]

    df_grp_src = df_filtered.copy()
    if hide_old_unassigned:
        df_grp_src = df_grp_src[
            (~df_grp_src["Nhóm2"].str.contains("Chưa phân", na=False)) &
            (~df_grp_src["Nhóm1"].str.contains("Chưa phân", na=False))
        ]

    if df_grp_src.empty:
        st.warning("⚠️ Các file đang chọn chưa có dữ liệu Nhóm 1 / Nhóm 2. Hãy chọn file mới hoặc bỏ tích ô 'Ẩn dữ liệu Chưa phân Nhóm'.")
        df_nhom2_sum = pd.DataFrame(columns=["Nhóm2", "So_SKU", "So_luong", "Doanh_thu", "Gia_von", "Loi_nhuan", "Tỷ trọng DT (%)", "Biên LN (%)"])
        df_hierarchy = pd.DataFrame(columns=["Nhóm2", "Nhóm1", "So_SKU", "So_luong", "Doanh_thu", "Gia_von", "Loi_nhuan", "Tỷ trọng DT (%)", "Biên LN (%)"])
        df_sun = pd.DataFrame()
        df_n2_line = pd.DataFrame()
        df_n1_line = pd.DataFrame()
    else:
        grp_rev_total = df_grp_src["Doanh thu"].sum()

        df_nhom2_sum = df_grp_src.groupby("Nhóm2", as_index=False).agg(
            So_SKU=("Mã hàng hóa", "nunique"),
            So_luong=("Số lượng", "sum"),
            Doanh_thu=("Doanh thu", "sum"),
            Gia_von=("Giá xuất kho", "sum"),
            Loi_nhuan=("Lợi nhuận", "sum")
        ).sort_values(grp_col, ascending=False)
        df_nhom2_sum["Tỷ trọng DT (%)"] = (df_nhom2_sum["Doanh_thu"] / grp_rev_total * 100) if grp_rev_total > 0 else 0
        df_nhom2_sum["Biên LN (%)"] = (df_nhom2_sum["Loi_nhuan"] / df_nhom2_sum["Doanh_thu"] * 100).where(df_nhom2_sum["Doanh_thu"] > 0, 0)

        df_hierarchy = df_grp_src.groupby(["Nhóm2", "Nhóm1"], as_index=False).agg(
            So_SKU=("Mã hàng hóa", "nunique"),
            So_luong=("Số lượng", "sum"),
            Doanh_thu=("Doanh thu", "sum"),
            Gia_von=("Giá xuất kho", "sum"),
            Loi_nhuan=("Lợi nhuận", "sum")
        ).sort_values(["Nhóm2", grp_col], ascending=[True, False])
        df_hierarchy["Tỷ trọng DT (%)"] = (df_hierarchy["Doanh_thu"] / grp_rev_total * 100) if grp_rev_total > 0 else 0
        df_hierarchy["Biên LN (%)"] = (df_hierarchy["Loi_nhuan"] / df_hierarchy["Doanh_thu"] * 100).where(df_hierarchy["Doanh_thu"] > 0, 0)

        g_col1, g_col2 = st.columns([1.1, 1.2])
        with g_col1:
            df_sun = df_hierarchy[df_hierarchy[grp_col] > 0].copy()
            if not df_sun.empty:
                fig_sun = px.sunburst(
                    df_sun,
                    path=["Nhóm2", "Nhóm1"],
                    values=grp_col,
                    color="Nhóm2",
                    color_discrete_sequence=px.colors.qualitative.Bold,
                    title=f"Cơ Cấu Tỷ Trọng {grp_metric_lbl}: Nhóm 2 (Vòng trong) → Nhóm 1 (Vòng ngoài)"
                )
                fig_sun.update_traces(textinfo="label+percent parent")
                fig_sun.update_layout(height=470, margin=dict(t=50, l=10, r=10, b=10))
                st.plotly_chart(fig_sun, width="stretch")
            else:
                st.info("Không có giá trị dương để vẽ biểu đồ tỷ trọng.")

        with g_col2:
            df_n1_rank = df_hierarchy.sort_values(grp_col, ascending=True).tail(15)
            fig_n1 = px.bar(
                df_n1_rank,
                x=grp_col,
                y="Nhóm1",
                color="Nhóm2",
                color_discrete_sequence=px.colors.qualitative.Bold,
                orientation="h",
                text=grp_col,
                title=f"Top 15 Nhóm 1 (Con) Cao Nhất Theo {grp_metric_lbl}",
                hover_data=["Nhóm2", "So_SKU", "So_luong", "Doanh_thu", "Loi_nhuan"]
            )
            txt_fmt = "%{text:,.2f}" if grp_col == "So_luong" else "%{text:,.0f}"
            fig_n1.update_traces(texttemplate=txt_fmt, textposition="outside")
            fig_n1.update_layout(height=470, yaxis_title="Nhóm 1 (Con)", xaxis_title=grp_metric_lbl, legend_title="Nhóm 2 (Cha)")
            st.plotly_chart(fig_n1, width="stretch")

        st.markdown("---")
        st.markdown(f"#### 📈 Xu Hướng Tăng/Giảm Theo Ngày Của Cấp Cha (Nhóm 2) & Cấp Con (Nhóm 1) — Theo {grp_metric_lbl}")

        df_grp_src["N1_Unique"] = df_grp_src["Nhóm1"] + " (" + df_grp_src["Nhóm2"].str[:1] + ")"
        n1_line_rank = (
            df_grp_src.groupby("N1_Unique")[grp_metric_lbl]
            .sum()
            .sort_values(ascending=False)
            .index.tolist()
        )
        picked_n1_lines = st.multiselect(
            "Chọn các Nhóm 1 (Con) hiển thị trên đồ thị đường bên phải (mặc định Top 8 Nhóm 1 cao nhất):",
            options=n1_line_rank,
            default=n1_line_rank[:8],
            key="picked_n1_lines"
        )

        line_col1, line_col2 = st.columns(2)
        with line_col1:
            n2_list = sorted(df_grp_src["Nhóm2"].unique().tolist())
            df_n2_line = build_complete_daily_series(df_grp_src, MASTER_DAYS_DF, "Nhóm2", n2_list, grp_metric_lbl)
            fig_line_n2 = build_line_chart_with_sum(
                df_n2_line,
                x_col="Ngày hiển thị",
                y_col=grp_metric_lbl,
                group_col="Nhóm2",
                ordered_items=n2_list,
                master_ordered_days=MASTER_ORDERED_DAYS,
                title=f"1. Đồ Thị Đường Cấp Cha (Nhóm 2) — {grp_metric_lbl}",
                legend_title="Nhóm 2 (Cha)",
                line_width=2.8,
                marker_size=7
            )
            st.plotly_chart(fig_line_n2, width="stretch")

        with line_col2:
            active_n1_lines = picked_n1_lines if picked_n1_lines else n1_line_rank[:8]
            df_n1_line = build_complete_daily_series(df_grp_src, MASTER_DAYS_DF, "N1_Unique", active_n1_lines, grp_metric_lbl)
            fig_line_n1 = build_line_chart_with_sum(
                df_n1_line,
                x_col="Ngày hiển thị",
                y_col=grp_metric_lbl,
                group_col="N1_Unique",
                ordered_items=active_n1_lines,
                master_ordered_days=MASTER_ORDERED_DAYS,
                title=f"2. Đồ Thị Đường Cấp Con (Nhóm 1) — {grp_metric_lbl}",
                legend_title="Nhóm 1 (Con)",
                line_width=2.5,
                marker_size=7
            )
            st.plotly_chart(fig_line_n1, width="stretch")

        st.markdown("---")
        b_col1, b_col2 = st.columns([1, 1.3])
        with b_col1:
            st.markdown("#### 📌 Tổng hợp theo Nhóm 2 (Ngành hàng cha)")
            st.dataframe(
                df_nhom2_sum.style.format({
                    "So_luong": "{:,.2f}", "Doanh_thu": "{:,.0f}",
                    "Gia_von": "{:,.0f}", "Loi_nhuan": "{:,.0f}",
                    "Tỷ trọng DT (%)": "{:.1f}%", "Biên LN (%)": "{:.1f}%"
                }),
                width="stretch",
                height=320
            )
        with b_col2:
            st.markdown("#### 📋 Chi tiết phân cấp Nhóm 2 → Nhóm 1 (Nhóm con)")
            st.dataframe(
                df_hierarchy.style.format({
                    "So_luong": "{:,.2f}", "Doanh_thu": "{:,.0f}",
                    "Gia_von": "{:,.0f}", "Loi_nhuan": "{:,.0f}",
                    "Tỷ trọng DT (%)": "{:.1f}%", "Biên LN (%)": "{:.1f}%"
                }),
                width="stretch",
                height=320
            )

# ==============================================================================
# TAB 2: PHÂN TÍCH PHÂN CẤP NHÓM 1 (CHA) -> SKU (CHI TIẾT) + 2 ĐỒ THỊ ĐƯỜNG
# ==============================================================================
with tab_n1_sku:
    st.subheader("📦 Cơ Cấu & Xu Hướng Bán Hàng: Nhóm 1 (Cha) → SKU (Chi tiết)")

    c_n1s_1, c_n1s_2, c_n1s_3 = st.columns([1.1, 1.3, 1.5])
    with c_n1s_1:
        n1s_metric_lbl = st.radio(
            "Chọn chỉ số phân tích N1 → SKU:",
            ["Doanh thu", "Lợi nhuận", "Số lượng"],
            horizontal=True,
            key="n1s_metric"
        )
    with c_n1s_2:
        hide_old_n1s = st.checkbox(
            "Ẩn dữ liệu '(Chưa phân Nhóm)' từ file cũ",
            value=True,
            key="hide_old_n1s"
        )
    with c_n1s_3:
        max_sku_per_n1 = st.slider(
            "Số SKU tối đa mỗi Nhóm 1 trên vòng tròn:",
            min_value=5, max_value=30, value=12, step=1
        )

    n1s_col_map = {"Doanh thu": "Doanh_thu", "Lợi nhuận": "Loi_nhuan", "Số lượng": "So_luong"}
    n1s_col = n1s_col_map[n1s_metric_lbl]

    df_n1s_src = df_filtered.copy()
    if hide_old_n1s:
        df_n1s_src = df_n1s_src[~df_n1s_src["Nhóm1"].str.contains("Chưa phân", na=False)]

    if df_n1s_src.empty:
        st.warning("⚠️ Các file đang chọn chưa có dữ liệu Nhóm 1. Hãy chọn file mới hoặc bỏ tích ô 'Ẩn dữ liệu Chưa phân Nhóm'.")
        df_n1_summary = pd.DataFrame()
        df_n1_sku_hier = pd.DataFrame()
        df_sun_n1s = pd.DataFrame()
        df_n1_cha_line = pd.DataFrame()
        df_sku_con_line = pd.DataFrame()
    else:
        n1_ranked_list = (
            df_n1s_src.groupby("Nhóm1")[n1s_metric_lbl]
            .sum()
            .sort_values(ascending=False)
            .index.tolist()
        )
        picked_n1_groups = st.multiselect(
            "🎯 Chọn các Nhóm 1 (Cha) muốn phân tích sâu xuống từng SKU (mặc định Top 8 Nhóm 1 lớn nhất):",
            options=n1_ranked_list,
            default=n1_ranked_list[:8],
            key="picked_n1_for_sku"
        )

        df_n1s_active = df_n1s_src[df_n1s_src["Nhóm1"].isin(picked_n1_groups)] if picked_n1_groups else df_n1s_src
        n1s_rev_total = df_n1s_active["Doanh thu"].sum()

        df_n1_summary = df_n1s_active.groupby("Nhóm1", as_index=False).agg(
            So_SKU=("Mã hàng hóa", "nunique"),
            So_luong=("Số lượng", "sum"),
            Doanh_thu=("Doanh thu", "sum"),
            Gia_von=("Giá xuất kho", "sum"),
            Loi_nhuan=("Lợi nhuận", "sum")
        ).sort_values(n1s_col, ascending=False)
        df_n1_summary["Tỷ trọng DT (%)"] = (df_n1_summary["Doanh_thu"] / n1s_rev_total * 100) if n1s_rev_total > 0 else 0
        df_n1_summary["Biên LN (%)"] = (df_n1_summary["Loi_nhuan"] / df_n1_summary["Doanh_thu"] * 100).where(df_n1_summary["Doanh_thu"] > 0, 0)

        df_n1_sku_hier = df_n1s_active.groupby(["Nhóm1", "Mã hàng hóa", "Tên hàng hóa", "SKU_Label", "Đơn vị tính"], as_index=False).agg(
            So_don=("Mã đơn hàng eShop", "nunique"),
            So_luong=("Số lượng", "sum"),
            Doanh_thu=("Doanh thu", "sum"),
            Gia_von=("Giá xuất kho", "sum"),
            Loi_nhuan=("Lợi nhuận", "sum")
        ).sort_values(["Nhóm1", n1s_col], ascending=[True, False])
        df_n1_sku_hier["Tỷ trọng DT (%)"] = (df_n1_sku_hier["Doanh_thu"] / n1s_rev_total * 100) if n1s_rev_total > 0 else 0
        df_n1_sku_hier["Biên LN (%)"] = (df_n1_sku_hier["Loi_nhuan"] / df_n1_sku_hier["Doanh_thu"] * 100).where(df_n1_sku_hier["Doanh_thu"] > 0, 0)

        ns_col1, ns_col2 = st.columns([1.1, 1.2])
        with ns_col1:
            df_sun_n1s = (
                df_n1_sku_hier[df_n1_sku_hier[n1s_col] > 0]
                .groupby("Nhóm1", group_keys=False)
                .head(max_sku_per_n1)
                .copy()
            )
            df_sun_n1s["SKU_Short"] = df_sun_n1s["Tên hàng hóa"].apply(lambda x: x[:26] + "..." if len(x) > 26 else x)

            if not df_sun_n1s.empty:
                fig_sun_n1s = px.sunburst(
                    df_sun_n1s,
                    path=["Nhóm1", "SKU_Short"],
                    values=n1s_col,
                    color="Nhóm1",
                    color_discrete_sequence=px.colors.qualitative.Vivid,
                    title=f"Cơ Cấu Tỷ Trọng {n1s_metric_lbl}: Nhóm 1 (Vòng trong) → SKU (Vòng ngoài)",
                    hover_data=["Mã hàng hóa", "Tên hàng hóa", "So_luong", "Doanh_thu", "Loi_nhuan"]
                )
                fig_sun_n1s.update_traces(textinfo="label+percent parent")
                fig_sun_n1s.update_layout(height=470, margin=dict(t=50, l=10, r=10, b=10))
                st.plotly_chart(fig_sun_n1s, width="stretch")
            else:
                st.info("Không có giá trị dương để vẽ biểu đồ tỷ trọng.")

        with ns_col2:
            df_sku_rank_n1 = df_n1_sku_hier.sort_values(n1s_col, ascending=True).tail(15)
            fig_sku_n1_bar = px.bar(
                df_sku_rank_n1,
                x=n1s_col,
                y="Tên hàng hóa",
                color="Nhóm1",
                color_discrete_sequence=px.colors.qualitative.Vivid,
                orientation="h",
                text=n1s_col,
                title=f"Top 15 SKU Cao Nhất Theo {n1s_metric_lbl} (Tô màu theo Nhóm 1)",
                hover_data=["Nhóm1", "Mã hàng hóa", "So_luong", "Doanh_thu", "Loi_nhuan"]
            )
            txt_fmt_n1s = "%{text:,.2f}" if n1s_col == "So_luong" else "%{text:,.0f}"
            fig_sku_n1_bar.update_traces(texttemplate=txt_fmt_n1s, textposition="outside")
            fig_sku_n1_bar.update_layout(height=470, yaxis_title="Sản phẩm (SKU)", xaxis_title=n1s_metric_lbl, legend_title="Nhóm 1")
            st.plotly_chart(fig_sku_n1_bar, width="stretch")

        st.markdown("---")
        st.markdown(f"#### 📈 Xu Hướng Tăng/Giảm Theo Ngày Của Cấp Cha (Nhóm 1) & Chi Tiết (SKU) — Theo {n1s_metric_lbl}")

        sku_line_rank = (
            df_n1s_active.groupby("SKU_Label")[n1s_metric_lbl]
            .sum()
            .sort_values(ascending=False)
            .index.tolist()
        )
        picked_sku_lines = st.multiselect(
            "Chọn các SKU chi tiết hiển thị trên đồ thị đường bên phải (mặc định Top 8 SKU cao nhất):",
            options=sku_line_rank,
            default=sku_line_rank[:8],
            key="picked_sku_lines_n1s"
        )

        ns_line1, ns_line2 = st.columns(2)
        with ns_line1:
            active_n1_cha = df_n1_summary["Nhóm1"].tolist()
            df_n1_cha_line = build_complete_daily_series(df_n1s_active, MASTER_DAYS_DF, "Nhóm1", active_n1_cha, n1s_metric_lbl)
            fig_line_n1_cha = build_line_chart_with_sum(
                df_n1_cha_line,
                x_col="Ngày hiển thị",
                y_col=n1s_metric_lbl,
                group_col="Nhóm1",
                ordered_items=active_n1_cha,
                master_ordered_days=MASTER_ORDERED_DAYS,
                title=f"1. Đồ Thị Đường Cấp Cha (Nhóm 1) — {n1s_metric_lbl}",
                legend_title="Nhóm 1 (Cha)",
                line_width=2.8,
                marker_size=7
            )
            st.plotly_chart(fig_line_n1_cha, width="stretch")

        with ns_line2:
            active_sku_con = picked_sku_lines if picked_sku_lines else sku_line_rank[:8]
            df_sku_con_line = build_complete_daily_series(df_n1s_active, MASTER_DAYS_DF, "SKU_Label", active_sku_con, n1s_metric_lbl)
            fig_line_sku_con = build_line_chart_with_sum(
                df_sku_con_line,
                x_col="Ngày hiển thị",
                y_col=n1s_metric_lbl,
                group_col="SKU_Label",
                ordered_items=active_sku_con,
                master_ordered_days=MASTER_ORDERED_DAYS,
                title=f"2. Đồ Thị Đường Chi Tiết (SKU) — {n1s_metric_lbl}",
                legend_title="Sản phẩm (SKU)",
                line_width=2.5,
                marker_size=7
            )
            st.plotly_chart(fig_line_sku_con, width="stretch")

        st.markdown("---")
        nb_col1, nb_col2 = st.columns([1, 1.4])
        with nb_col1:
            st.markdown("#### 📌 Tổng hợp theo Nhóm 1")
            st.dataframe(
                df_n1_summary.style.format({
                    "So_luong": "{:,.2f}", "Doanh_thu": "{:,.0f}",
                    "Gia_von": "{:,.0f}", "Loi_nhuan": "{:,.0f}",
                    "Tỷ trọng DT (%)": "{:.1f}%", "Biên LN (%)": "{:.1f}%"
                }),
                width="stretch",
                height=340
            )
        with nb_col2:
            st.markdown("#### 📋 Chi tiết phân cấp Nhóm 1 → SKU")
            cols_n1_sku_tbl = ["Nhóm1", "Mã hàng hóa", "Tên hàng hóa", "Đơn vị tính", "So_luong", "Doanh_thu", "Loi_nhuan", "Tỷ trọng DT (%)", "Biên LN (%)"]
            st.dataframe(
                df_n1_sku_hier[cols_n1_sku_tbl].style.format({
                    "So_luong": "{:,.2f}", "Doanh_thu": "{:,.0f}",
                    "Loi_nhuan": "{:,.0f}", "Tỷ trọng DT (%)": "{:.1f}%",
                    "Biên LN (%)": "{:.1f}%"
                }),
                width="stretch",
                height=340
            )

# ==============================================================================
# TAB 3: XU HƯỚNG THEO NGÀY
# ==============================================================================
with tab1:
    c1, c2 = st.columns([1.5, 1])
    with c1:
        fig_daily = go.Figure()
        fig_daily.add_trace(go.Bar(
            x=df_daily["Ngày hiển thị"], y=df_daily["Doanh_thu"],
            name="Doanh thu", marker_color="#2E86C1",
            text=df_daily["Doanh_thu"], texttemplate="%{text:,.0f}", textposition="outside"
        ))
        fig_daily.add_trace(go.Bar(
            x=df_daily["Ngày hiển thị"], y=df_daily["Loi_nhuan"],
            name="Lợi nhuận", marker_color="#28B463",
            text=df_daily["Loi_nhuan"], texttemplate="%{text:,.0f}", textposition="outside"
        ))
        fig_daily.update_layout(
            title="Doanh Thu & Lợi Nhuận Theo Ngày (DD/MM/YYYY)",
            barmode="group", height=430, yaxis_title="VNĐ", xaxis_title="Ngày bán"
        )
        fig_daily.update_xaxes(type="category", categoryorder="array", categoryarray=MASTER_ORDERED_DAYS)
        st.plotly_chart(fig_daily, width="stretch")

    with c2:
        fig_orders = px.line(
            df_daily, x="Ngày hiển thị", y="So_don", markers=True,
            title="Số Lượng Đơn Hàng Theo Ngày", text="So_don"
        )
        fig_orders.update_traces(textposition="top center", line_color="#E67E22", line_width=3)
        fig_orders.update_layout(height=430, yaxis_title="Số đơn hàng", xaxis_title="Ngày bán")
        fig_orders.update_xaxes(type="category", categoryorder="array", categoryarray=MASTER_ORDERED_DAYS)
        st.plotly_chart(fig_orders, width="stretch")

    st.subheader("Bảng đối soát tổng hợp theo từng ngày")
    st.dataframe(
        df_daily.drop(columns=["Ngày_dt"]).style.format({
            "Doanh_thu": "{:,.0f}", "Loi_nhuan": "{:,.0f}",
            "Gia_von": "{:,.0f}", "So_luong": "{:,.2f}"
        }),
        width="stretch"
    )

# ==============================================================================
# TAB 4: BIẾN ĐỘNG TĂNG GIẢM SỐ LƯỢNG TỪNG SKU THEO NGÀY
# ==============================================================================
with tab2:
    st.subheader("📈 Biến Động Số Lượng Bán Từng SKU Theo Ngày")
    sku_ranking = (
        df_filtered.groupby("SKU_Label", as_index=False)["Số lượng"]
        .sum()
        .sort_values("Số lượng", ascending=False)
    )
    all_skus = sku_ranking["SKU_Label"].tolist()
    default_skus = all_skus[:5]

    col_sku1, col_sku2 = st.columns([3, 1])
    with col_sku1:
        selected_skus = st.multiselect(
            "Chọn 1 hoặc nhiều mã hàng (SKU) để theo dõi biến động theo ngày:",
            options=all_skus,
            default=default_skus,
            placeholder="Gõ mã hàng hoặc tên sản phẩm để tìm và chọn..."
        )
    with col_sku2:
        sku_chart_style = st.radio(
            "Kiểu hiển thị:",
            ["Biểu đồ đường (Line)", "Biểu đồ cột nhóm (Bar)"],
            horizontal=False
        )

    active_skus = selected_skus if selected_skus else default_skus
    df_sku_complete = build_complete_daily_series(df_filtered, MASTER_DAYS_DF, "SKU_Label", active_skus, "Số lượng")

    if sku_chart_style == "Biểu đồ đường (Line)":
        fig_sku = build_line_chart_with_sum(
            df_sku_complete,
            x_col="Ngày hiển thị",
            y_col="Số lượng",
            group_col="SKU_Label",
            ordered_items=active_skus,
            master_ordered_days=MASTER_ORDERED_DAYS,
            title="Biến Động Số Lượng Bán Ra Của Từng SKU Qua Các Ngày",
            legend_title="Sản phẩm (SKU)",
            line_width=2.5,
            marker_size=7
        )
        fig_sku.update_layout(height=470)
    else:
        fig_sku = px.bar(
            df_sku_complete, x="Ngày hiển thị", y="Số lượng", color="SKU_Label",
            color_discrete_sequence=px.colors.qualitative.Bold,
            barmode="group", text="Số lượng", title="So Sánh Số Lượng Bán Ra Của Từng SKU Qua Các Ngày"
        )
        fig_sku.update_traces(texttemplate="%{text:,.1f}", textposition="outside")
        fig_sku.update_layout(height=470, xaxis_title="Ngày bán", yaxis_title="Số lượng bán", legend_title="Sản phẩm (SKU)", hovermode="x unified")
        fig_sku.update_xaxes(type="category", categoryorder="array", categoryarray=MASTER_ORDERED_DAYS)

    st.plotly_chart(fig_sku, width="stretch")

    st.markdown("#### 📊 Bảng ma trận số lượng & mức tăng/giảm theo ngày")
    pivot_sku = df_sku_complete.pivot(index="SKU_Label", columns="Ngày hiển thị", values="Số lượng")[MASTER_ORDERED_DAYS]
    pivot_sku["Tổng cộng"] = pivot_sku.sum(axis=1)
    pivot_sku["TB / Ngày"] = pivot_sku[MASTER_ORDERED_DAYS].mean(axis=1)
    if len(MASTER_ORDERED_DAYS) >= 2:
        last_day, prev_day = MASTER_ORDERED_DAYS[-1], MASTER_ORDERED_DAYS[-2]
        pivot_sku[f"Tăng/Giảm ({last_day} vs {prev_day})"] = pivot_sku[last_day] - pivot_sku[prev_day]
    st.dataframe(pivot_sku.style.format("{:,.2f}"), width="stretch")

# ==============================================================================
# TAB 5: TOP SẢN PHẨM
# ==============================================================================
with tab3:
    metric_choice = st.radio(
        "Xếp hạng Top sản phẩm theo tiêu chí:",
        ["Doanh thu", "Lợi nhuận", "Số lượng"],
        horizontal=True
    )
    col_map = {"Doanh thu": "Doanh_thu", "Lợi nhuận": "Loi_nhuan", "Số lượng": "So_luong"}
    sort_col = col_map[metric_choice]

    df_top = df_prod.nlargest(top_n, sort_col).sort_values(sort_col, ascending=True)
    fig_top = px.bar(
        df_top, x=sort_col, y="Tên hàng hóa", orientation="h",
        color="Nhóm2",
        color_discrete_sequence=px.colors.qualitative.Bold,
        text=sort_col, title=f"Top {top_n} Mặt Hàng Cao Nhất Theo {metric_choice}",
        hover_data=["Nhóm2", "Nhóm1", "Mã hàng hóa", "So_luong", "Doanh_thu", "Loi_nhuan"]
    )
    fig_top.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
    fig_top.update_layout(height=520, yaxis_title="", xaxis_title=metric_choice, legend_title="Nhóm 2")
    st.plotly_chart(fig_top, width="stretch")

# ==============================================================================
# TAB 6: CẢNH BÁO LỖ & GIÁ VỐN
# ==============================================================================
with tab4:
    c_warn1, c_warn2 = st.columns(2)
    c_warn1.error(f"🔻 Có **{len(df_loss)} dòng giao dịch bán lỗ** (Tổng lỗ: **{df_loss['Lợi nhuận'].sum():,.0f} đ**)")
    c_warn2.warning(f"❓ Có **{len(df_zero_cost)} dòng chưa có giá xuất kho (= 0)** (Doanh thu: **{df_zero_cost['Doanh thu'].sum():,.0f} đ**)")

    cols_show = ["Ngày hiển thị", "Mã đơn hàng eShop", "Nhóm2", "Nhóm1", "Mã hàng hóa", "Tên hàng hóa", "Số lượng", "Đơn giá trước thuế", "Doanh thu", "Giá xuất kho", "Lợi nhuận"]
    st.subheader("1. Chi tiết các giao dịch có Lợi nhuận âm")
    st.dataframe(df_loss[cols_show], width="stretch")

    st.subheader("2. Chi tiết các giao dịch Giá xuất kho = 0")
    st.dataframe(df_zero_cost[cols_show], width="stretch")

# ==============================================================================
# TAB 7: BẢNG SỔ CHI TIẾT
# ==============================================================================
with tab5:
    display_cols = [c for c in df_filtered.columns if c not in ["Ngày_dt", "Ngày hiển thị", "SKU_Label"]]
    st.dataframe(df_filtered[display_cols], width="stretch", height=450)

# ==============================================================================
# HÀM XUẤT BÁO CÁO EXCEL & POWERPOINT (FULL-FRAME 16:9 KHÔNG CẮT XÉN)
# ==============================================================================
def render_chart_fitted(fig, box_w_in, box_h_in, fallback_df=None, x_col=None, y_col=None, color_col=None, kind="bar", title=""):
    px_w = int(box_w_in * 120)
    px_h = int(box_h_in * 120)
    try:
        fig_export = go.Figure(fig)
        fig_export.update_layout(
            template="plotly_white",
            paper_bgcolor="white",
            plot_bgcolor="white",
            width=px_w,
            height=px_h,
            autosize=False,
            font=dict(color="#111111", size=11),
            title=dict(font=dict(size=13, color="#17375E"), x=0.02, xanchor="left"),
            margin=dict(t=48, l=45, r=25, b=45),
            legend=dict(font=dict(size=9.5))
        )
        return fig_export.to_image(format="png", width=px_w, height=px_h, scale=2)
    except Exception:
        if not HAS_MPL or fallback_df is None or fallback_df.empty:
            return None
        plt.style.use("default")
        fig_m, ax = plt.subplots(figsize=(box_w_in, box_h_in), dpi=180, facecolor="white")
        ax.set_facecolor("white")

        if kind == "line" and color_col and color_col in fallback_df.columns:
            for idx, (grp_val, grp_df) in enumerate(fallback_df.groupby(color_col)):
                c = PALETTE_COLORS[idx % len(PALETTE_COLORS)]
                ax.plot(grp_df[x_col].astype(str), grp_df[y_col], marker="o", linewidth=2.0, markersize=4, color=c, label=str(grp_val)[:24])
            ax.legend(fontsize=7, loc="best", frameon=True, facecolor="white")
        elif kind == "barh":
            sub_df = fallback_df.tail(15)
            colors = [PALETTE_COLORS[i % len(PALETTE_COLORS)] for i in range(len(sub_df))]
            ax.barh(sub_df[y_col].astype(str).str[:28], sub_df[x_col], color=colors)
        elif kind == "pie":
            sub_df = fallback_df.groupby(color_col, as_index=False)[y_col].sum()
            colors = [PALETTE_COLORS[i % len(PALETTE_COLORS)] for i in range(len(sub_df))]
            ax.pie(
                sub_df[y_col],
                labels=sub_df[color_col].astype(str).str[:22],
                autopct="%1.1f%%",
                colors=colors,
                startangle=90,
                textprops={"fontsize": 7.5, "color": "#111111"}
            )
        elif kind == "daily_group":
            x_vals = range(len(fallback_df))
            w = 0.38
            ax.bar([i - w/2 for i in x_vals], fallback_df["Doanh_thu"], width=w, color="#2E86C1", label="Doanh thu")
            ax.bar([i + w/2 for i in x_vals], fallback_df["Loi_nhuan"], width=w, color="#28B463", label="Lợi nhuận")
            ax.set_xticks(list(x_vals))
            ax.set_xticklabels(fallback_df[x_col].astype(str))
            ax.legend(fontsize=7.5, facecolor="white")
        else:
            ax.bar(fallback_df[x_col].astype(str), fallback_df[y_col], color="#E67E22")

        ax.set_title(title, fontsize=10, fontweight="bold", color="#17375E", pad=8)
        ax.tick_params(colors="#222222", labelsize=7.5)
        for spine in ax.spines.values():
            spine.set_edgecolor("#CCCCCC")
        if kind != "pie":
            ax.grid(axis="y", linestyle="--", alpha=0.45, color="#CCCCCC")
            plt.xticks(rotation=30, ha="right")
        plt.tight_layout()
        buf = io.BytesIO()
        fig_m.savefig(buf, format="png", facecolor="white", edgecolor="none")
        plt.close(fig_m)
        return buf.getvalue()

def build_excel_report():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_fmt = workbook.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E78"})
        num_fmt = workbook.add_format({"num_format": "#,##0"})
        pct_fmt = workbook.add_format({"num_format": "0.0%"})

        df_daily_out = df_daily.drop(columns=["Ngày_dt"]).copy()
        df_daily_out.columns = ["Ngày", "Doanh thu", "Lợi nhuận", "Giá vốn", "Số đơn hàng", "Số lượng bán", "Số dòng GD"]
        df_daily_out.to_excel(writer, sheet_name="1_Tong_Quan_Ngay", index=False, startrow=1)
        ws1 = writer.sheets["1_Tong_Quan_Ngay"]
        ws1.write(0, 0, "BÁO CÁO TỔNG HỢP DOANH THU & LỢI NHUẬN THEO NGÀY - ABL", title_fmt)
        ws1.set_column("A:A", 14)
        ws1.set_column("B:D", 18, num_fmt)
        ws1.set_column("E:G", 14, num_fmt)

        num_rows = len(df_daily_out)
        if num_rows > 0:
            chart1 = workbook.add_chart({"type": "column"})
            chart1.add_series({
                "name": ["1_Tong_Quan_Ngay", 1, 1],
                "categories": ["1_Tong_Quan_Ngay", 2, 0, 1 + num_rows, 0],
                "values": ["1_Tong_Quan_Ngay", 2, 1, 1 + num_rows, 1],
                "fill": {"color": "#2E86C1"},
            })
            chart1.add_series({
                "name": ["1_Tong_Quan_Ngay", 1, 2],
                "categories": ["1_Tong_Quan_Ngay", 2, 0, 1 + num_rows, 0],
                "values": ["1_Tong_Quan_Ngay", 2, 2, 1 + num_rows, 2],
                "fill": {"color": "#28B463"},
            })
            chart1.set_title({"name": "Biểu đồ Doanh thu & Lợi nhuận theo Ngày"})
            chart1.set_size({"width": 720, "height": 380})
            ws1.insert_chart(f"A{num_rows + 5}", chart1)

        if not df_nhom2_sum.empty:
            df_n2_out = df_nhom2_sum.copy()
            df_n2_out["Tỷ trọng DT (%)"] = df_n2_out["Tỷ trọng DT (%)"] / 100.0
            df_n2_out["Biên LN (%)"] = df_n2_out["Biên LN (%)"] / 100.0
            df_n2_out.to_excel(writer, sheet_name="2_Phan_Tich_N2_N1", index=False, startrow=1)

            df_h_out = df_hierarchy.copy()
            df_h_out["Tỷ trọng DT (%)"] = df_h_out["Tỷ trọng DT (%)"] / 100.0
            df_h_out["Biên LN (%)"] = df_h_out["Biên LN (%)"] / 100.0
            start_h = len(df_n2_out) + 5
            df_h_out.to_excel(writer, sheet_name="2_Phan_Tich_N2_N1", index=False, startrow=start_h)

            ws2 = writer.sheets["2_Phan_Tich_N2_N1"]
            ws2.write(0, 0, "TỔNG HỢP THEO NHÓM 2 (NGÀNH HÀNG CHA)", title_fmt)
            ws2.write(start_h - 1, 0, "CHI TIẾT PHÂN CẤP NHÓM 2 -> NHÓM 1 (NHÓM CON)", title_fmt)
            ws2.set_column("A:B", 32)
            ws2.set_column("C:G", 16, num_fmt)
            ws2.set_column("H:I", 14, pct_fmt)

        df_prod_out = df_prod.sort_values(["Nhóm1", "Doanh_thu"], ascending=[True, False]).copy()
        df_prod_out.columns = ["Nhóm 2", "Nhóm 1", "Mã hàng hóa", "Tên hàng hóa", "ĐVT", "Số lượng", "Doanh thu", "Giá vốn", "Lợi nhuận", "Số đơn"]
        df_prod_out.to_excel(writer, sheet_name="3_Phan_Tich_N1_SKU", index=False)
        ws3 = writer.sheets["3_Phan_Tich_N1_SKU"]
        ws3.set_column("A:B", 26)
        ws3.set_column("C:C", 14)
        ws3.set_column("D:D", 40)
        ws3.set_column("F:J", 16, num_fmt)

        pivot_sku.reset_index().to_excel(writer, sheet_name="4_Bien_Dong_SKU", index=False)
        ws4 = writer.sheets["4_Bien_Dong_SKU"]
        ws4.set_column("A:A", 45)
        ws4.set_column("B:Z", 15, num_fmt)

        cols_warn = ["Ngày hiển thị", "Mã đơn hàng eShop", "Nhóm2", "Nhóm1", "Mã hàng hóa", "Tên hàng hóa", "Số lượng", "Đơn giá trước thuế", "Doanh thu", "Giá xuất kho", "Lợi nhuận"]
        df_loss[cols_warn].to_excel(writer, sheet_name="5_Canh_Bao_Lo", index=False)

        detail_cols = [c for c in df_filtered.columns if c not in ["Ngày_dt", "Ngày hiển thị", "SKU_Label"]]
        df_filtered[detail_cols].to_excel(writer, sheet_name="6_So_Chi_Tiet", index=False)

    return output.getvalue()

def build_pptx_report():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def init_dashboard_slide(title_text, badge_text=""):
        slide = prs.slides.add_slide(blank_layout)
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(248, 250, 252)

        hdr = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.78))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = RGBColor(23, 55, 94)
        hdr.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0.45), Inches(0.16), Inches(9.5), Inches(0.48))
        p = tb.text_frame.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        if badge_text:
            tb_r = slide.shapes.add_textbox(Inches(9.6), Inches(0.22), Inches(3.3), Inches(0.4))
            pr = tb_r.text_frame.paragraphs[0]
            pr.text = badge_text
            pr.alignment = PP_ALIGN.RIGHT
            pr.font.size = Pt(12)
            pr.font.color.rgb = RGBColor(174, 214, 241)

        return slide

    def add_fitted_pic(slide, img_bytes, left_in, top_in, width_in, height_in):
        if not img_bytes:
            return
        card = slide.shapes.add_shape(1, Inches(left_in - 0.04), Inches(top_in - 0.04), Inches(width_in + 0.08), Inches(height_in + 0.08))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(220, 226, 232)
        card.line.width = Pt(1)
        slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(left_in), Inches(top_in), width=Inches(width_in), height=Inches(height_in))

    date_label = f"{MASTER_ORDERED_DAYS[0]} → {MASTER_ORDERED_DAYS[-1]}" if len(MASTER_ORDERED_DAYS) > 1 else f"Ngày {MASTER_ORDERED_DAYS[0]}"

    # SLIDE 1: TRANG BÌA
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = RGBColor(23, 55, 94)
    bg1.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(0.9), Inches(2.4), Inches(11.5), Inches(2.8))
    tf1 = tb1.text_frame
    p1 = tf1.paragraphs[0]
    p1.text = "BÁO CÁO PHÂN TÍCH BÁN HÀNG & CƠ CẤU NHÓM HÀNG - ABL"
    p1.font.size = Pt(34)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(255, 255, 255)

    p2 = tf1.add_paragraph()
    p2.text = f"\nGiai đoạn dữ liệu: {date_label}   |   Tổng đơn hàng: {total_orders:,} đơn   |   Doanh thu: {total_rev:,.0f} đ"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(174, 214, 241)

    # SLIDE 2: TỔNG QUAN KPI + XU HƯỚNG NGÀY
    s2 = init_dashboard_slide("1. TỔNG QUAN DASHBOARD: CHỈ SỐ KPI & XU HƯỚNG THEO NGÀY", date_label)
    kpis = [
        ("💰 DOANH THU (TRƯỚC THUẾ)", f"{total_rev:,.0f} đ", RGBColor(46, 134, 193)),
        ("📈 LỢI NHUẬN GỘP", f"{total_profit:,.0f} đ", RGBColor(40, 180, 99)),
        ("📊 BIÊN LỢI NHUẬN", f"{margin_pct:.1f}%", RGBColor(22, 160, 133)),
        ("🧾 TỔNG SỐ ĐƠN HÀNG", f"{total_orders:,} đơn", RGBColor(230, 126, 34)),
        ("📦 TỔNG SL / AOV", f"{total_qty:,.0f} ({aov:,.0f}đ)", RGBColor(142, 68, 173)),
    ]
    card_w = 2.38
    gap_w = 0.16
    for idx, (label, val, color) in enumerate(kpis):
        left = 0.4 + idx * (card_w + gap_w)
        box = s2.shapes.add_shape(1, Inches(left), Inches(0.98), Inches(card_w), Inches(1.15))
        box.fill.solid()
        box.fill.fore_color.rgb = color
        box.line.fill.background()
        tf_k = box.text_frame
        tf_k.word_wrap = True
        pk1 = tf_k.paragraphs[0]
        pk1.text = label
        pk1.font.size = Pt(10.5)
        pk1.font.bold = True
        pk1.font.color.rgb = RGBColor(255, 255, 255)
        pk2 = tf_k.add_paragraph()
        pk2.text = val
        pk2.font.size = Pt(17)
        pk2.font.bold = True
        pk2.font.color.rgb = RGBColor(255, 255, 255)

    img_daily = render_chart_fitted(fig_daily, 7.4, 4.85, df_daily, "Ngày hiển thị", "Doanh_thu", kind="daily_group", title="Doanh Thu & Lợi Nhuận Theo Ngày")
    add_fitted_pic(s2, img_daily, 0.4, 2.35, 7.4, 4.85)

    img_orders = render_chart_fitted(fig_orders, 4.8, 4.85, df_daily, "Ngày hiển thị", "So_don", kind="bar", title="Số Lượng Đơn Hàng Theo Ngày")
    add_fitted_pic(s2, img_orders, 8.13, 2.35, 4.8, 4.85)

    # SLIDE 3 & 4: NHÓM 2 -> NHÓM 1
    if not df_grp_src.empty:
        s3 = init_dashboard_slide(f"2. PHÂN TÍCH NHÓM HÀNG (N2 → N1) — CHỈ SỐ: {grp_metric_lbl.upper()}", date_label)
        if not df_sun.empty:
            img_sun = render_chart_fitted(fig_sun, 5.8, 6.2, df_nhom2_sum, y_col=grp_col, color_col="Nhóm2", kind="pie", title=f"Cơ Cấu Tỷ Trọng {grp_metric_lbl}: N2 → N1")
            add_fitted_pic(s3, img_sun, 0.4, 0.98, 5.8, 6.2)

        img_n1 = render_chart_fitted(fig_n1, 6.45, 6.2, df_n1_rank, x_col=grp_col, y_col="Nhóm1", kind="barh", title=f"Top 15 Nhóm 1 (Con) Cao Nhất Theo {grp_metric_lbl}")
        add_fitted_pic(s3, img_n1, 6.48, 0.98, 6.45, 6.2)

        s4 = init_dashboard_slide(f"3. XU HƯỚNG NGÀY (N2 → N1): CẤP CHA & CẤP CON — {grp_metric_lbl.upper()}", date_label)
        img_ln2 = render_chart_fitted(fig_line_n2, 6.12, 6.2, df_n2_line, "Ngày hiển thị", grp_metric_lbl, "Nhóm2", kind="line", title=f"1. Đồ Thị Đường Cấp Cha (Nhóm 2) — {grp_metric_lbl}")
        add_fitted_pic(s4, img_ln2, 0.4, 0.98, 6.12, 6.2)

        img_ln1 = render_chart_fitted(fig_line_n1, 6.12, 6.2, df_n1_line, "Ngày hiển thị", grp_metric_lbl, "N1_Unique", kind="line", title=f"2. Đồ Thị Đường Cấp Con (Nhóm 1) — {grp_metric_lbl}")
        add_fitted_pic(s4, img_ln1, 6.81, 0.98, 6.12, 6.2)

    # SLIDE 5 & 6: NHÓM 1 -> SKU
    if not df_n1s_src.empty:
        s5 = init_dashboard_slide(f"4. PHÂN TÍCH NHÓM HÀNG (N1 → SKU) — CHỈ SỐ: {n1s_metric_lbl.upper()}", date_label)
        if not df_sun_n1s.empty:
            img_sun_n1s = render_chart_fitted(fig_sun_n1s, 5.8, 6.2, df_n1_summary, y_col=n1s_col, color_col="Nhóm1", kind="pie", title=f"Cơ Cấu Tỷ Trọng {n1s_metric_lbl}: N1 → SKU")
            add_fitted_pic(s5, img_sun_n1s, 0.4, 0.98, 5.8, 6.2)

        img_sku_n1_bar = render_chart_fitted(fig_sku_n1_bar, 6.45, 6.2, df_sku_rank_n1, x_col=n1s_col, y_col="Tên hàng hóa", kind="barh", title=f"Top 15 SKU Cao Nhất Theo {n1s_metric_lbl}")
        add_fitted_pic(s5, img_sku_n1_bar, 6.48, 0.98, 6.45, 6.2)

        s6 = init_dashboard_slide(f"5. XU HƯỚNG NGÀY (N1 → SKU): CẤP CHA & CHI TIẾT — {n1s_metric_lbl.upper()}", date_label)
        img_ln1_cha = render_chart_fitted(fig_line_n1_cha, 6.12, 6.2, df_n1_cha_line, "Ngày hiển thị", n1s_metric_lbl, "Nhóm1", kind="line", title=f"1. Đồ Thị Đường Cấp Cha (Nhóm 1) — {n1s_metric_lbl}")
        add_fitted_pic(s6, img_ln1_cha, 0.4, 0.98, 6.12, 6.2)

        img_lsku_con = render_chart_fitted(fig_line_sku_con, 6.12, 6.2, df_sku_con_line, "Ngày hiển thị", n1s_metric_lbl, "SKU_Label", kind="line", title=f"2. Đồ Thị Đường Chi Tiết (SKU) — {n1s_metric_lbl}")
        add_fitted_pic(s6, img_lsku_con, 6.81, 0.98, 6.12, 6.2)

    # SLIDE 7: TOP SẢN PHẨM & BIẾN ĐỘNG SKU
    s7 = init_dashboard_slide(f"6. TOP {top_n} SẢN PHẨM ({metric_choice.upper()}) & BIẾN ĐỘNG SỐ LƯỢNG SKU", date_label)
    img_top = render_chart_fitted(fig_top, 6.12, 6.2, df_top, x_col=sort_col, y_col="Tên hàng hóa", kind="barh", title=f"Top {top_n} Mặt Hàng Theo {metric_choice}")
    add_fitted_pic(s7, img_top, 0.4, 0.98, 6.12, 6.2)

    img_sku_all = render_chart_fitted(fig_sku, 6.12, 6.2, df_sku_complete, "Ngày hiển thị", "Số lượng", "SKU_Label", kind="line", title="Biến Động Số Lượng Bán Ra Của Từng SKU")
    add_fitted_pic(s7, img_sku_all, 6.81, 0.98, 6.12, 6.2)

    ppt_out = io.BytesIO()
    prs.save(ppt_out)
    return ppt_out.getvalue()

# --- KHU VỰC XUẤT BÁO CÁO Ở MENU TRÁI ---
st.sidebar.markdown("---")
st.sidebar.header("📤 Xuất file báo cáo")

if st.sidebar.button("📗 Tạo file Excel (.xlsx + Đồ thị)", width="stretch"):
    with st.spinner("Đang đóng gói file Excel..."):
        try:
            excel_bytes = build_excel_report()
            st.sidebar.download_button(
                label="⬇️ Nhấn để lưu file Excel",
                data=excel_bytes,
                file_name="Bao_Cao_Ban_Hang_ABL_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )
        except Exception as e:
            st.sidebar.warning(f"Cần cài xlsxwriter để xuất Excel ({e})")

if HAS_PPTX:
    if st.sidebar.button("📙 Tạo file PowerPoint (.pptx)", width="stretch"):
        with st.spinner("Đang xuất đồ thị Full Dashboard sang PowerPoint..."):
            try:
                pptx_bytes = build_pptx_report()
                st.sidebar.download_button(
                    label="⬇️ Nhấn để lưu file PowerPoint",
                    data=pptx_bytes,
                    file_name="Bao_Cao_Ban_Hang_ABL_2026.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    width="stretch"
                )
            except Exception as e:
                st.sidebar.error(f"Lỗi xuất PowerPoint: {e}")
