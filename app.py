import streamlit as st
import pandas as pd
import io

# 設定網頁為寬螢幕模式，讓資料表更好閱讀
st.set_page_config(layout="wide")
st.title("📊 進階資料分析與重組工具")

# --- 系統狀態初始化 (用來記憶資料和當前步驟) ---
if 'current_df' not in st.session_state:
    st.session_state.current_df = None
if 'file_id' not in st.session_state:
    st.session_state.file_id = None
if 'action_mode' not in st.session_state:
    st.session_state.action_mode = 'select_cols'

# ==========================================
# 需求 1：能上傳 Excel 表格並進行更新
# ==========================================
st.markdown("### 1. 檔案上傳與更新")
uploaded_file = st.file_uploader("請上傳 Excel 檔案 (重新上傳即可更新資料)", type=["xlsx", "xls"])

if uploaded_file is not None:
    # 偵測是否為新上傳的檔案 (或內容被更新過的檔案)
    if st.session_state.file_id != uploaded_file.file_id:
        st.session_state.current_df = pd.read_excel(uploaded_file)
        st.session_state.file_id = uploaded_file.file_id
        st.session_state.action_mode = 'select_cols' # 重置回預設模式

    df = st.session_state.current_df

    # ==========================================
    # 需求 2：上傳過後可看到所有資料 (中區塊、10筆高度、滾輪)
    # ==========================================
    st.markdown("### 2. 資料總覽")
    st.info("提示：此區塊已固定高度，一次最多顯示約 10 筆資料，請使用滑鼠滾輪上下滑動檢視。")
    # height=380 像素，大約等於 10 筆資料列 + 標題列的高度
    st.dataframe(df, height=380, use_container_width=True)
    st.markdown("---")

    # ==========================================
    # 需求 3：視窗下方欄顯示 2 個按鈕 (切換功能)
    # ==========================================
    st.markdown("### 3. 操作功能選擇")
    
    # 建立兩個並排的按鈕來切換模式
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("🔘 選取所需欄位", use_container_width=True):
            st.session_state.action_mode = 'select_cols'
    with col_btn2:
        if st.button("🔘 設定數據範圍", use_container_width=True):
            st.session_state.action_mode = 'set_range'

    # ------------------------------------------
    # 需求 3-1：選取所需欄位
    # ------------------------------------------
    if st.session_state.action_mode == 'select_cols':
        st.subheader("📌 3-1. 選取所需欄位")
        
        # 3-1-1 顯示所有欄位名稱供點選 (多選)
        all_columns = df.columns.tolist()
        selected_cols = st.multiselect("請點選您需要的欄位 (可多選)：", all_columns)
        
        if selected_cols:
            # 3-1-2 抓出並分開工作表檢視
            df_selected = df[selected_cols]
            df_unselected = df.drop(columns=selected_cols)
            
            col_view1, col_view2 = st.columns(2)
            with col_view1:
                st.write("✅ **已選取的欄位資料：**")
                st.dataframe(df_selected, height=300, use_container_width=True)
            with col_view2:
                st.write("❌ **未選取的欄位資料：**")
                st.dataframe(df_unselected, height=300, use_container_width=True)
            
            # 匯出與狀態傳遞區塊
            st.markdown("#### 導出與下一步")
            out_col1, out_col2 = st.columns(2)
            
            with out_col1:
                # 導出新 Excel (兩個工作表)
                output1 = io.BytesIO()
                with pd.ExcelWriter(output1, engine='openpyxl') as writer:
                    df_selected.to_excel(writer, index=False, sheet_name='已選取欄位')
                    df_unselected.to_excel(writer, index=False, sheet_name='未選取欄位')
                
                st.download_button(
                    label="📥 下載【欄位分割】Excel",
                    data=output1.getvalue(),
                    file_name="欄位分割結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with out_col2:
                # 顯示按鈕，按下去以新資料設定數據
                def proceed_to_range():
                    st.session_state.current_df = df_selected # 覆寫記憶體中的資料
                    st.session_state.action_mode = 'set_range' # 切換到 3-2 介面
                    
                st.button(
                    "➡️ 以【已選取欄位】進入設定數據範圍", 
                    on_click=proceed_to_range, 
                    type="primary", 
                    use_container_width=True
                )

    # ------------------------------------------
    # 需求 3-2：設定數據按鈕功能
    # ------------------------------------------
    elif st.session_state.action_mode == 'set_range':
        st.subheader("📌 3-2. 設定數據範圍")
        st.info("目前正使用您最後確認的資料結構進行數據設定。")
        
        # 3-2-1 設定範圍區塊
        set_col1, set_col2, set_col3 = st.columns([2, 1, 2])
        with set_col1:
            target_col = st.selectbox("1. 選擇要套用條件的欄位", df.columns)
        with set_col2:
            operator = st.selectbox("2. 邏輯條件", ["<=", ">=", "==", "<", ">"])
        with set_col3:
            value = st.number_input("3. 輸入數值 (X)", value=0.0)
            
        if target_col:
            # 抓出「未填」或空白的資料 (處理包含字串"未填"、空值 NaN、或純空白)
            is_missing = (
                df[target_col].isna() | 
                (df[target_col].astype(str).str.strip() == "未填") | 
                (df[target_col].astype(str).str.strip() == "")
            )
            df_missing = df[is_missing]
            df_valid = df[~is_missing] # 剩下的就是有填寫的資料
            
            # 將要判斷的資料強制轉為數字 (無法轉數字的會暫時變成 NaN)
            valid_numeric = pd.to_numeric(df_valid[target_col], errors='coerce')
            
            # 根據使用者選擇的邏輯執行判斷
            if operator == "<=": mask = valid_numeric <= value
            elif operator == ">=": mask = valid_numeric >= value
            elif operator == "==": mask = valid_numeric == value
            elif operator == "<": mask = valid_numeric < value
            elif operator == ">": mask = valid_numeric > value
            
            # 處理無法判斷的錯誤值 (將它們預設為不符合)
            mask = mask.fillna(False)
            
            df_in_range = df_valid[mask]
            df_out_range = df_valid[~mask]
            
            # 分開檢視工作表
            st.markdown("#### 分類檢視結果")
            tab1, tab2, tab3 = st.tabs(["✅ 範圍內的數據", "❌ 範圍外的數據", "⚠️ 未填/格式異常數據"])
            
            with tab1:
                st.write(f"符合 `{target_col} {operator} {value}` (共 {len(df_in_range)} 筆)")
                st.dataframe(df_in_range, height=250, use_container_width=True)
            with tab2:
                st.write(f"不符合條件 (共 {len(df_out_range)} 筆)")
                st.dataframe(df_out_range, height=250, use_container_width=True)
            with tab3:
                st.write(f"包含空值或標示為「未填」 (共 {len(df_missing)} 筆)")
                st.dataframe(df_missing, height=250, use_container_width=True)
                
            # 導出新 Excel
            st.markdown("#### 導出分類結果")
            output2 = io.BytesIO()
            with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                df_in_range.to_excel(writer, index=False, sheet_name='範圍內數據')
                df_out_range.to_excel(writer, index=False, sheet_name='範圍外數據')
                
                # 若有未填文字，額外生成第三個工作表
                if not df_missing.empty:
                    df_missing.to_excel(writer, index=False, sheet_name='未填數據')
                    sheet_count = 3
                else:
                    sheet_count = 2
                    
            st.download_button(
                label=f"📥 下載【數據分類】Excel (包含 {sheet_count} 個工作表)",
                data=output2.getvalue(),
                file_name=f"數據分類結果_{target_col}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )