import io
import re
import hashlib
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from sqlalchemy import create_engine

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
st.set_page_config(page_title="Báo Cáo Bán Hàng ABL 2026 (v18 - Đồng nhất màu)", page_icon="📊", layout="wide")
st.title("📊 Dashboard Phân Tích Bán Hàng & Cơ Cấu Nhóm Hàng - ABL")

# 2. Khóa mật khẩu đăng nhập nội bộ
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    pwd = st.text_input("🔐 Nhập mật khẩu truy cập báo cáo nội bộ:", type="password")
    if pwd == st.secrets["APP_PASSWORD"]:
        st.session_state.authenticated = True
        st.rerun()
    elif pwd:
        st.error("❌ Sai mật khẩu truy cập!")
    st.stop()

# Bảng màu mở rộng 24 màu phân biệt rõ nét
MASTER_PALETTE = [
    "#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD",
    "#8C564B", "#E377C2", "#17BECF", "#BCBD22", "#393B79",
    "#E6550D", "#31A354", "#756BB1", "#636363", "#AD494A",
    "#008080", "#E7298A", "#66A61E", "#E6AB02", "#A6761D",
    "#1B9E77", "#D95F02", "#7570B3", "#B2182B"
]

def build_fixed_color_map(df_full):
    """Tạo từ điển màu cố định cho tất cả Nhóm 2, Nhóm 1 và SKU từ dữ liệu gốc"""
    color_map = {}

    # 1. Cố định màu cho Nhóm 2
    if "Nhóm2" in df_full.columns:
        all_n2 = sorted(df_full["Nhóm2"].dropna().unique().tolist())
        for idx, n2 in enumerate(all_n2):
            color_map[n2] = MASTER_PALETTE[idx % len(MASTER_PALETTE)]

    # 2. Cố định màu cho Nhóm 1
    if "Nhóm1" in df_full.columns:
        all_n1 = sorted(df_full["Nhóm1"].dropna().unique().tolist())
        for idx, n1 in enumerate(all_n1):
            c = MASTER_PALETTE[(idx + 5) % len(MASTER_PALETTE)]
            color_map[n1] = c

    # Map thêm cho N1_Unique (dùng ở Tab 1)
    if "Nhóm2" in df_full.columns and "Nhóm1" in df_full.columns:
        pairs = df_full[["Nhóm1", "Nhóm2"]].dropna().drop_duplicates()
        for _, row in pairs.iterrows():
            u_name = f"{row['Nhóm1']} ({str(row['Nhóm2'])[:1]})"
            color_map[u_name] = color_map.get(row["Nhóm1"], "#1F77B4")

    # 3. Cố định màu cho SKU (dùng hash MD5 để mỗi mã hàng luôn ra 1 màu duy nhất)
    if "SKU_Label" in df_full.columns:
        for sku in df_full["SKU_Label"].dropna().unique():
            h = int(hashlib.md5(str(sku).encode("utf-8")).hexdigest(), 16)
            color_map[sku] = MASTER_PALETTE[h % len(MASTER_PALETTE)]

    return color_map

def build_complete_daily_series(df_source, master_days_df, group_col, selected_items, value_col):
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

def build_line_chart_with_sum(df_line, x_col, y_col, group_col, ordered_items, master_ordered_days, title, legend_title, color_map=None, line_width=2.5, marker_size=7):
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
        if color_map and item in color_map:
            color = color_map[item]
        else:
            color = MASTER_PALETTE[idx % len(MASTER_PALETTE)]
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

# 3. Tải dữ liệu trực tiếp từ Supabase SQL
@st.cache_data(ttl=600, show_spinner="🔄 Đang tải dữ liệu từ máy chủ...")
def load_sql_data():
    engine = create_engine(st.secrets["DATABASE_URL"])
    df_sql = pd.read_sql("SELECT * FROM ban_hang_abl", engine)
    df_sql["Ngày_dt"] = pd.to_datetime(df_sql["Ngày_dt"], errors="coerce")
    return df_sql

df_all_sql = load_sql_data()

if df_all_sql.empty:
    st.error("Bảng dữ liệu ban_hang_abl trên Supabase đang trống.")
    st.stop()

# Khởi tạo từ điển màu cố định toàn cục cho toàn bộ báo cáo
GLOBAL_COLOR_MAP = build_fixed_color_map(df_all_sql)

# --- THANH BỘ LỌC BÊN TRÁI (SIDEBAR) ---
st.sidebar.header("📂 Dữ liệu đầu vào")
if st.sidebar.button("🔄 Làm mới dữ liệu từ SQL"):
    st.cache_data.clear()
    st.rerun()

all_sources = sorted(df_all_sql["Nguồn file"].dropna().unique().tolist())
selected_files = st.sidebar.multiselect(
    "Chọn các ngày / file báo cáo:",
    options=all_sources,
    default=all_sources
)

if not selected_files:
    st.warning("⚠️ Vui lòng chọn ít nhất 1 file báo cáo ở thanh bên trái.")
    st.stop()

df_raw_input = df_all_sql[df_all_sql["Nguồn file"].isin(selected_files)].copy()

st.sidebar.markdown("---")
st.sidebar.header("🔍 Bộ lọc phân tích")

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
# TAB 1: PHÂN TÍCH PHÂN CẤP NHÓM 2 (CHA) -> NHÓM 1 (CON)
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
                    color_discrete_map=GLOBAL_COLOR_MAP,
                    title=f"Cơ Cấu Tỷ Trọng {grp_metric_lbl}: Nhóm 2 (Vòng trong) → Nhóm 1 (Vòng ngoài)"
                )
                fig_sun.update_traces(textinfo="label+percent parent")
                fig_sun.update_layout(height=470, margin=dict(t=50, l=10, r=10, b=10))
                st.plotly_chart(fig_sun, use_container_width=True)
            else:
                st.info("Không có giá trị dương để vẽ biểu đồ tỷ trọng.")

        with g_col2:
            df_n1_rank = df_hierarchy.sort_values(grp_col, ascending=True).tail(15)
            fig_n1 = px.bar(
                df_n1_rank,
                x=grp_col,
                y="Nhóm1",
                color="Nhóm2",
                color_discrete_map=GLOBAL_COLOR_MAP,
                orientation="h",
                text=grp_col,
                title=f"Top 15 Nhóm 1 (Con) Cao Nhất Theo {grp_metric_lbl}",
                hover_data=["Nhóm2", "So_SKU", "So_luong", "Doanh_thu", "Loi_nhuan"]
            )
            txt_fmt = "%{text:,.2f}" if grp_col == "So_luong" else "%{text:,.0f}"
            fig_n1.update_traces(texttemplate=txt_fmt, textposition="outside")
            fig_n1.update_layout(height=470, yaxis_title="Nhóm 1 (Con)", xaxis_title=grp_metric_lbl, legend_title="Nhóm 2 (Cha)")
            st.plotly_chart(fig_n1, use_container_width=True)

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
                color_map=GLOBAL_COLOR_MAP,
                line_width=2.8,
                marker_size=7
            )
            st.plotly_chart(fig_line_n2, use_container_width=True)

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
                color_map=GLOBAL_COLOR_MAP,
                line_width=2.5,
                marker_size=7
            )
            st.plotly_chart(fig_line_n1, use_container_width=True)

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
                use_container_width=True,
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
                use_container_width=True,
                height=320
            )

# ==============================================================================
# TAB 2: PHÂN TÍCH PHÂN CẤP NHÓM 1 (CHA) -> SKU (CHI TIẾT)
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
                    color_discrete_map=GLOBAL_COLOR_MAP,
                    title=f"Cơ Cấu Tỷ Trọng {n1s_metric_lbl}: Nhóm 1 (Vòng trong) → SKU (Vòng ngoài)",
                    hover_data=["Mã hàng hóa", "Tên hàng hóa", "So_luong", "Doanh_thu", "Loi_nhuan"]
                )
                fig_sun_n1s.update_traces(textinfo="label+percent parent")
                fig_sun_n1s.update_layout(height=470, margin=dict(t=50, l=10, r=10, b=10))
                st.plotly_chart(fig_sun_n1s, use_container_width=True)
            else:
                st.info("Không có giá trị dương để vẽ biểu đồ tỷ trọng.")

        with ns_col2:
            df_sku_rank_n1 = df_n1_sku_hier.sort_values(n1s_col, ascending=True).tail(15)
            fig_sku_n1_bar = px.bar(
                df_sku_rank_n1,
                x=n1s_col,
                y="Tên hàng hóa",
                color="Nhóm1",
                color_discrete_map=GLOBAL_COLOR_MAP,
                orientation="h",
                text=n1s_col,
                title=f"Top 15 SKU Cao Nhất Theo {n1s_metric_lbl} (Tô màu theo Nhóm 1)",
                hover_data=["Nhóm1", "Mã hàng hóa", "So_luong", "Doanh_thu", "Loi_nhuan"]
            )
            txt_fmt_n1s = "%{text:,.2f}" if n1s_col == "So_luong" else "%{text:,.0f}"
            fig_sku_n1_bar.update_traces(texttemplate=txt_fmt_n1s, textposition="outside")
            fig_sku_n1_bar.update_layout(height=470, yaxis_title="Sản phẩm (SKU)", xaxis_title=n1s_metric_lbl, legend_title="Nhóm 1")
            st.plotly_chart(fig_sku_n1_bar, use_container_width=True)

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
                color_map=GLOBAL_COLOR_MAP,
                line_width=2.8,
                marker_size=7
            )
            st.plotly_chart(fig_line_n1_cha, use_container_width=True)

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
                color_map=GLOBAL_COLOR_MAP,
                line_width=2.5,
                marker_size=7
            )
            st.plotly_chart(fig_line_sku_con, use_container_width=True)

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
                use_container_width=True,
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
                use_container_width=True,
                height=340
            )

# ==============================================================================
# TAB 3: XU HƯỚNG THEO NGÀY
# ==============================================================================
with tab1:
    st.subheader("📅 Diễn biến Doanh thu & Lợi nhuận theo ngày")
    fig_day = go.Figure()
    fig_day.add_trace(go.Bar(
        x=df_daily["Ngày hiển thị"], y=df_daily["Doanh_thu"],
        name="Doanh thu", marker_color="#2E86C1",
        text=df_daily["Doanh_thu"].apply(lambda x: f"{x/1e6:.1f}M" if x >= 1e6 else f"{x:,.0f}"),
        textposition="outside"
    ))
    fig_day.add_trace(go.Scatter(
        x=df_daily["Ngày hiển thị"], y=df_daily["Loi_nhuan"],
        name="Lợi nhuận gộp", mode="lines+markers",
        line=dict(color="#28B463", width=3), marker=dict(size=8)
    ))
    fig_day.update_layout(height=430, xaxis_title="Ngày bán", yaxis_title="VNĐ", hovermode="x unified")
    fig_day.update_xaxes(type="category", categoryorder="array", categoryarray=MASTER_ORDERED_DAYS)
    st.plotly_chart(fig_day, use_container_width=True)

    st.markdown("#### 📋 Bảng tổng hợp theo từng ngày")
    df_daily_show = df_daily.drop(columns=["Ngày_dt"]).copy()
    df_daily_show["Biên LN (%)"] = (df_daily_show["Loi_nhuan"] / df_daily_show["Doanh_thu"] * 100).where(df_daily_show["Doanh_thu"] > 0, 0)
    st.dataframe(
        df_daily_show.style.format({
            "Doanh_thu": "{:,.0f}", "Loi_nhuan": "{:,.0f}", "Gia_von": "{:,.0f}",
            "So_luong": "{:,.2f}", "Biên LN (%)": "{:.1f}%"
        }),
        use_container_width=True
    )

# ==============================================================================
# TAB 4: BIẾN ĐỘNG TỪNG SKU THEO NGÀY
# ==============================================================================
with tab2:
    st.subheader("📈 Biến Động Doanh Số / Lợi Nhuận / Số Lượng Của Từng Mặt Hàng (SKU)")
    c_sku1, c_sku2 = st.columns([2, 1])
    with c_sku1:
        sku_by_rev = (
            df_filtered.groupby("SKU_Label")["Doanh thu"]
            .sum()
            .sort_values(ascending=False)
            .index.tolist()
        )
        selected_skus = st.multiselect(
            "Chọn các mặt hàng muốn so sánh (mặc định Top 5 doanh thu cao nhất):",
            options=sku_by_rev,
            default=sku_by_rev[:5]
        )
    with c_sku2:
        metric_choice = st.radio("Chỉ số hiển thị:", ["Doanh thu", "Lợi nhuận", "Số lượng"], horizontal=True)

    if selected_skus:
        df_sku_line = build_complete_daily_series(df_filtered, MASTER_DAYS_DF, "SKU_Label", selected_skus, metric_choice)
        fig_sku_line = build_line_chart_with_sum(
            df_sku_line,
            x_col="Ngày hiển thị",
            y_col=metric_choice,
            group_col="SKU_Label",
            ordered_items=selected_skus,
            master_ordered_days=MASTER_ORDERED_DAYS,
            title=f"So sánh biến động {metric_choice} từng mặt hàng theo ngày",
            legend_title="Mặt hàng (SKU)",
            color_map=GLOBAL_COLOR_MAP,
            line_width=2.5,
            marker_size=7
        )
        st.plotly_chart(fig_sku_line, use_container_width=True)

        df_pivot_sku = df_sku_line.pivot(index="SKU_Label", columns="Ngày hiển thị", values=metric_choice).fillna(0)
        df_pivot_sku = df_pivot_sku.reindex(index=selected_skus, columns=MASTER_ORDERED_DAYS)
        df_pivot_sku["TỔNG CỘNG"] = df_pivot_sku.sum(axis=1)
        fmt_str = "{:,.2f}" if metric_choice == "Số lượng" else "{:,.0f}"
        st.dataframe(df_pivot_sku.style.format(fmt_str), use_container_width=True)
    else:
        st.info("Vui lòng chọn ít nhất 1 mặt hàng ở ô phía trên để xem biểu đồ biến động.")

# ==============================================================================
# TAB 5: TOP SẢN PHẨM
# ==============================================================================
with tab3:
    col_a, col_b = st.columns(2)
    with col_a:
        top_rev_df = df_prod.sort_values("Doanh_thu", ascending=True).tail(top_n)
        fig_rev = px.bar(
            top_rev_df, x="Doanh_thu", y="Tên hàng hóa", color="Nhóm2",
            color_discrete_map=GLOBAL_COLOR_MAP, orientation="h",
            title=f"Top {top_n} Doanh Thu Cao Nhất", text="Doanh_thu",
            hover_data=["Nhóm1", "Mã hàng hóa", "So_luong", "Loi_nhuan"]
        )
        fig_rev.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig_rev.update_layout(height=480, yaxis_title="")
        st.plotly_chart(fig_rev, use_container_width=True)

    with col_b:
        top_prof_df = df_prod.sort_values("Loi_nhuan", ascending=True).tail(top_n)
        fig_prof = px.bar(
            top_prof_df, x="Loi_nhuan", y="Tên hàng hóa", color="Nhóm2",
            color_discrete_map=GLOBAL_COLOR_MAP, orientation="h",
            title=f"Top {top_n} Lợi Nhuận Gộp Cao Nhất", text="Loi_nhuan",
            hover_data=["Nhóm1", "Mã hàng hóa", "So_luong", "Doanh_thu"]
        )
        fig_prof.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig_prof.update_layout(height=480, yaxis_title="")
        st.plotly_chart(fig_prof, use_container_width=True)

    st.markdown("#### 📊 Bảng xếp hạng chi tiết theo mặt hàng")
    df_prod_show = df_prod.sort_values("Doanh_thu", ascending=False).copy()
    df_prod_show["Biên LN (%)"] = (df_prod_show["Loi_nhuan"] / df_prod_show["Doanh_thu"] * 100).where(df_prod_show["Doanh_thu"] > 0, 0)
    st.dataframe(
        df_prod_show.style.format({
            "So_luong": "{:,.2f}", "Doanh_thu": "{:,.0f}",
            "Gia_von": "{:,.0f}", "Loi_nhuan": "{:,.0f}", "Biên LN (%)": "{:.1f}%"
        }),
        use_container_width=True
    )

# ==============================================================================
# TAB 6: CẢNH BÁO LỖ & GIÁ VỐN
# ==============================================================================
with tab4:
    c_warn1, c_warn2 = st.columns(2)
    with c_warn1:
        st.subheader(f"🔴 Giao dịch bán lỗ ({len(df_loss)} dòng)")
        if df_loss.empty:
            st.success("Không có giao dịch nào bị âm lợi nhuận!")
        else:
            cols_w = ["Ngày hiển thị", "Nhóm2", "Nhóm1", "Mã đơn hàng eShop", "Mã hàng hóa", "Tên hàng hóa", "Số lượng", "Doanh thu", "Giá xuất kho", "Lợi nhuận"]
            st.dataframe(
                df_loss[cols_w].style.format({"Số lượng": "{:,.2f}", "Doanh thu": "{:,.0f}", "Giá xuất kho": "{:,.0f}", "Lợi nhuận": "{:,.0f}"}),
                use_container_width=True
            )

    with c_warn2:
        st.subheader(f"🟡 Giao dịch Giá xuất kho = 0 ({len(df_zero_cost)} dòng)")
        if df_zero_cost.empty:
            st.success("Tất cả mặt hàng có doanh thu đều đã cập nhật giá vốn!")
        else:
            cols_z = ["Ngày hiển thị", "Nhóm2", "Nhóm1", "Mã đơn hàng eShop", "Mã hàng hóa", "Tên hàng hóa", "Số lượng", "Doanh thu"]
            st.dataframe(
                df_zero_cost[cols_z].style.format({"Số lượng": "{:,.2f}", "Doanh thu": "{:,.0f}"}),
                use_container_width=True
            )

# ==============================================================================
# TAB 7: SỔ CHI TIẾT & XUẤT FILE EXCEL / POWERPOINT
# ==============================================================================
with tab5:
    st.subheader("📋 Dữ liệu chi tiết sau lọc & Xuất báo cáo")
    display_cols = [
        "Ngày hiển thị", "Nhóm2", "Nhóm1", "Mã đơn hàng eShop", "Số hóa đơn",
        "Mã hàng hóa", "Tên hàng hóa", "Đơn vị tính", "Số lượng",
        "Đơn giá trước thuế", "Doanh thu", "Thuế GTGT", "Giá xuất kho", "Lợi nhuận", "Nguồn file"
    ]
    st.dataframe(
        df_filtered[display_cols].style.format({
            "Số lượng": "{:,.2f}", "Đơn giá trước thuế": "{:,.0f}",
            "Doanh thu": "{:,.0f}", "Thuế GTGT": "{:,.0f}",
            "Giá xuất kho": "{:,.0f}", "Lợi nhuận": "{:,.0f}"
        }),
        use_container_width=True,
        height=420
    )

    st.markdown("---")
    st.markdown("#### 📥 Xuất Báo Cáo Đầy Đủ")

    def build_excel_report():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_kpi = pd.DataFrame({
                "Chỉ tiêu": [
                    "Thời gian báo cáo", "Doanh thu trước thuế (VNĐ)",
                    "Lợi nhuận gộp (VNĐ)", "Biên lợi nhuận (%)",
                    "Tổng số đơn hàng", "Giá trị TB / Đơn - AOV (VNĐ)", "Tổng số lượng bán ra"
                ],
                "Giá trị": [
                    f"{MASTER_ORDERED_DAYS[0]} -> {MASTER_ORDERED_DAYS[-1]}" if MASTER_ORDERED_DAYS else "",
                    round(total_rev, 0), round(total_profit, 0), round(margin_pct, 2),
                    total_orders, round(aov, 0), round(total_qty, 2)
                ]
            })
            df_kpi.to_excel(writer, sheet_name="Tong_Quan_KPI", index=False)
            df_daily_show.to_excel(writer, sheet_name="Xu_Huong_Ngay", index=False)
            df_prod_show.to_excel(writer, sheet_name="Xep_Hang_SKU", index=False)
            df_filtered[display_cols].to_excel(writer, sheet_name="So_Chi_Tiet", index=False)
        return output.getvalue()

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="📥 Tải Báo Cáo Tổng Hợp (.xlsx)",
            data=build_excel_report(),
            file_name="Bao_Cao_Ban_Hang_ABL_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
