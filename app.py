import io
import pandas as pd
import streamlit as st

# Cấu hình trang Web
st.set_page_config(page_title="员工数据多字段核对工具", page_icon="🎄", layout="wide")

st.title("🎄 员工数据多字段核对与自动匹配工具")
st.write(
    "基于 **工号** 比对两个 Excel 文件，自动提取姓名、考核日期、上岗证等详细信息。"
)

st.sidebar.header("📂 上传 Excel 文件")
file_a = st.sidebar.file_uploader(
    "上传主文件要核对（例如：9.15 表）", type=["xlsx", "xls"], key="file_a"
)
file_b = st.sidebar.file_uploader(
    "上传对比文件/源数据（包含姓名、考核日期等）", type=["xlsx", "xls"], key="file_b"
)

if file_a and file_b:
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    st.subheader("⚙️ 1. 选择比对与匹配字段")
    col1, col2 = st.columns(2)

    with col1:
        col_key_a = st.selectbox(
            "选择【主文件】的工号列:",
            options=df_a.columns,
            index=df_a.columns.get_loc("工号") if "工号" in df_a.columns else 0,
        )

    with col2:
        col_key_b = st.selectbox(
            "选择【对比文件】的工号列:",
            options=df_b.columns,
            index=df_b.columns.get_loc("工号") if "工号" in df_b.columns else 0,
        )

    # Chọn các cột thông tin cần lấy từ File B (Tên, Ngày thi, Mã thẻ...)
    available_cols_b = [c for c in df_b.columns if c != col_key_b]
    
    # Tự động gợi ý chọn các cột phổ biến nếu có trong file
    default_selected = [
        c for c in available_cols_b 
        if any(keyword in str(c) for keyword in ["姓名", "考核日期", "上岗证", "岗证", "日期"])
    ]

    selected_extra_cols = st.multiselect(
        "选择要从【对比文件】中提取的附加信息列（例如：姓名、考核日期、上岗证）：",
        options=available_cols_b,
        default=default_selected if default_selected else available_cols_b[:2]
    )

    # Dọn dẹp khoảng trắng và đồng bộ kiểu chuỗi (Text) để đảm bảo match chính xác 100%
    df_a["_clean_key"] = df_a[col_key_a].astype(str).str.strip()
    df_b["_clean_key"] = df_b[col_key_b].astype(str).str.strip()

    # Xóa dữ liệu trùng lặp trong File B (chỉ lấy dòng đầu tiên của mỗi 工号)
    df_b_unique = df_b.drop_duplicates(subset=["_clean_key"], keep="first")

    # Tiến hành ghé/tra cứu dữ liệu (tương tự INDEX+MATCH / VLOOKUP)
    cols_to_merge = ["_clean_key"] + selected_extra_cols
    df_merged = pd.merge(
        df_a, 
        df_b_unique[cols_to_merge], 
        on="_clean_key", 
        how="left"
    )

    # Tạo cột trạng thái đối soát
    mask_matched = df_merged["_clean_key"].isin(df_b_unique["_clean_key"])
    df_merged["比对结果"] = mask_matched.map({True: "匹配成功/已存在", False: "未找到/缺失"})

    # Xóa cột tạm
    df_merged.drop(columns=["_clean_key"], inplace=True)

    # Phân loại dữ liệu
    df_trung = df_merged[mask_matched].copy()
    df_khong_trung = df_merged[~mask_matched].copy()

    st.markdown("---")
    st.subheader("📈 比对结果统计")

    m1, m2, m3 = st.columns(3)
    m1.metric("主文件总行数", len(df_a))
    m2.metric("成功匹配数量", len(df_trung))
    m3.metric("未匹配/缺失数量", len(df_khong_trung))

    # Tab hiển thị dữ liệu
    tab1, tab2, tab3 = st.tabs(
        ["📑 完整提取结果表", "⚠️ 成功匹配列表", "❌ 未匹配/缺失列表"]
    )

    with tab1:
        st.dataframe(df_merged, use_container_width=True)

    with tab2:
        st.dataframe(df_trung, use_container_width=True)

    with tab3:
        st.dataframe(df_khong_trung, use_container_width=True)

    # Xuất file Excel kết quả ra bộ nhớ đệm
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_merged.to_excel(writer, sheet_name="完整匹配结果", index=False)
        df_trung.to_excel(writer, sheet_name="成功匹配列表", index=False)
        df_khong_trung.to_excel(writer, sheet_name="未匹配列表", index=False)

    buffer.seek(0)

    st.markdown("---")
    st.download_button(
        label="📥 下载提取与比对结果 Excel 文件 (.xlsx)",
        data=buffer,
        file_name="员工多字段比对结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("👈 请在左侧边栏上传两个 Excel 文件以开始比对与信息提取。")