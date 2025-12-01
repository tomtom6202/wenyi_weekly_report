import streamlit as st
import scraper_logic as sl
import pandas as pd
import time
import io

# ====================================================================
# 1. 頁面配置與標題
# ====================================================================

st.set_page_config(
    page_title="文一週報數據爬取工具",
    layout="centered",
    initial_sidebar_state="expanded"
)

st.title("📊 文一會所週報數據自動化")
st.markdown("請輸入要抓取的**年份和週數**，然後點擊『執行爬蟲與數據處理』按鈕。")

# ====================================================================
# 2. 應用程式主邏輯
# ====================================================================

# 獲取預設值 (避免在運行主邏輯時重複計算)
try:
    # 這裡假設 sl.show_X_days_ago 和 sl.year_week_X_days_ago 已經定義在 scraper_logic.py 中
    default_date_info = sl.show_X_days_ago(6)
    default_year_week = sl.year_week_X_days_ago(-1)
except Exception as e:
    # 設置防呆，如果 scraper_logic 初始計算失敗
    default_date_info = "無法計算"
    default_year_week = "2025,01"
    st.error(f"初始化錯誤：{e}")


# 顯示帶有預設值的提示信息
info_message = f"**請輸入年份和週數**（格式範例：`2025,48`）。"
info_message += f"或使用以下預設值：6天前那一週，即 **{default_date_info}** ==> 『**{default_year_week}**』"

st.info(info_message, icon="ℹ️")


# 創建使用者輸入框
year_week_input = st.text_input(
    label="👉 輸入週數 YYYY,WW:",
    value=default_year_week, 
    help="例如：2025,48。請確保週數格式正確。"
)

# 執行按鈕
if st.button("🚀 執行爬蟲與數據處理", type="primary"):
    
    # 檢查輸入格式
    if not year_week_input or ',' not in year_week_input:
        st.error("請輸入正確的週數格式，例如：`2025,48`。")
        st.session_state.data_ready = False
        # 參數檢查失敗，提前退出
        st.stop()
    
    # === 1. 讀取 Streamlit Secrets (獲取帳密資訊) ===
    try:
        # 從 Canvas secrets.toml 中的 [church_details] 區塊讀取
        church_details = st.secrets["church_details"] 
    except KeyError:
        st.error("❌ 錯誤：Streamlit Secrets 中缺少 `church_details` 配置。請檢查您的密鑰設置。")
        st.session_state.data_ready = False 
        # 密鑰未找到，停止執行後續爬蟲
        st.stop() 

    # 使用 st.status 來顯示進度，取代 Colab 中的 print
    with st.status(f"開始爬取 {year_week_input} 週的數據...", expanded=True) as status:
        try:
            status.update(label="1/3 🔑 正在嘗試登入並獲取數據...", state="running", expanded=True)
            time.sleep(1) # 模擬工作
            
            # === 2. 【已修正】調用後端主邏輯，並傳遞 church_details 密鑰字典 (不再接收 check_result) ===
            total_excel_bytes, report_excel_bytes, final_year_week, preview_df = \
                sl.run_scraper_and_process(year_week_input, church_details) # <-- 傳遞密鑰

            status.update(label="2/3 🧮 數據處理與表格生成中...", state="running", expanded=True)
            time.sleep(1) # 模擬工作

            # 將結果存入 session_state，讓數據和下載按鈕保持不變
            st.session_state.data_ready = True
            st.session_state.total_excel_bytes = total_excel_bytes
            st.session_state.report_excel_bytes = report_excel_bytes
            st.session_state.preview_df = preview_df
            st.session_state.file_prefix = final_year_week.replace(',', '_')
            
            status.update(label="3/3 ✅ 所有數據已成功處理！", state="complete", expanded=False)
            st.success(f"數據處理完成！請在下方下載 Excel 報告。")

        except Exception as e:
            status.update(label="❌ 爬蟲或數據處理失敗！", state="error", expanded=True)
            st.error(f"處理失敗。錯誤訊息：{e}")
            # st.exception(e) # 顯示更詳細的錯誤給開發者
            st.session_state.data_ready = False


# ====================================================================
# 3. 輸出與下載區塊
# ====================================================================

if 'data_ready' in st.session_state and st.session_state.data_ready:
    
    file_prefix = st.session_state.file_prefix

    st.header("下載區塊")
    st.markdown("---")

    col1, col2 = st.columns(2)

    # 下載第一個 Excel 檔案 (文一總數據)
    with col1:
        st.subheader("檔案 1：文一總數據")
        st.download_button(
            label="下載 文一總數據 Excel",
            data=st.session_state.total_excel_bytes,
            file_name=f"{file_prefix}_文一總數據.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="包含所有會所的原始爬取數據和專項數據。"
        )

    # 下載第二個 Excel 檔案 (文一每週報表 - 帶格式)
    with col2:
        st.subheader("檔案 2：文一每週報表")
        st.download_button(
            label="下載 文一每週報表 Excel (帶格式)",
            data=st.session_state.report_excel_bytes,
            file_name=f"{file_prefix}_文一每週報表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="包含計算後的結果和合併儲存格等格式。"
        )

    st.markdown("---")
    st.subheader("數據預覽 (表單填寫所需數據)")
    st.caption("以下是 `表單填寫所需數據` 的 DataFrame 預覽，用於快速檢查數據是否正確。")
    
    # 顯示 what_we_need_10 的預覽
    if st.session_state.preview_df is not None and not st.session_state.preview_df.empty:
        # 使用 to_html 並設置樣式來更好地呈現 DataFrame
        st.dataframe(st.session_state.preview_df.style.set_table_styles([
            {'selector': 'th', 'props': [('font-size', '10pt'), ('background-color', '#f0f2f6')]},
            {'selector': 'td', 'props': [('font-size', '10pt')]}
        ]), use_container_width=True)
    else:
        st.warning("預覽數據為空，請檢查爬蟲過程是否正確。")