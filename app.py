import os
import re
import io
import json
import time
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from tavily import TavilyClient

# 頁面標題與佈局
st.set_page_config(page_title="AI 批量資料處理服務", page_icon="🤖", layout="wide")
st.title("🤖 AI 批量資料處理與搜尋系統")

# 側邊欄：設定參數
with st.sidebar:
    st.header("⚙️ 系統參數設定")
    model_name = st.text_input("Gemini 模型名稱", value="gemini-2.5-flash")
    gemini_api_key = st.text_input("Gemini API Key", type="password")
    tavily_api_key = st.text_input("Tavily API Key", type="password")
    
    st.divider()
    run_tavily = st.checkbox("執行 Tavily 搜尋", value=True)
    run_gemini = st.checkbox("執行 Gemini 生成", value=True)
    tavily_max_results = st.number_input("Tavily 搜尋最大回傳筆數", value=10, min_value=1, max_value=20)
    target_group = st.number_input("目標組別 / 項次", value=1, min_value=1)

# 主要內容區：上傳檔案
uploaded_file = st.file_uploader("📂 上傳共用 Excel 檔 (包含設定與資料分頁)", type=["xlsx"])

def search_web_tavily(tavily_client, query, max_results):
    clean_query = query.replace("．", "").replace("・", "").strip()
    try:
        response = tavily_client.search(
            query=clean_query,
            max_results=max_results,
            search_depth="basic",
            include_answer=True
        )
        results = response.get("results", [])
        ai_answer = response.get("answer", "").strip()

        context_parts = []
        if ai_answer:
            context_parts.append(f"【Tavily 綜合搜尋摘要】:\n{ai_answer}\n")
        if results:
            context_parts.append("【來源網頁詳細資料】:")
            for r in results:
                context_parts.append(
                    f"標題: {r.get('title', '')}\n摘要: {r.get('content', '')}\n網址: {r.get('url', '')}\n"
                )
        return "\n".join(context_parts)
    except Exception as e:
        st.warning(f"⚠️ Tavily 搜尋警告: {e}")
        return ""

def call_ai_engine(client, prompt, model_name, is_batch=False, max_retries=4):
    """具備自動重試（Exponential Backoff）機制的 Gemini 呼叫函數"""
    for attempt in range(max_retries):
        try:
            config = types.GenerateContentConfig(response_mime_type="application/json")
            response = client.models.generate_content(model=model_name, contents=prompt, config=config)
            raw_text = response.text.strip()
            
            if "```" in raw_text:
                raw_text = re.sub(r'```(?:json)?', '', raw_text).replace('```', '').strip()
            
            match = re.search(r'\[.*\]', raw_text, re.DOTALL) if is_batch else re.search(r'\{.*\}', raw_text, re.DOTALL)
            if not match and is_batch:
                match = re.search(r'\{.*\}', raw_text, re.DOTALL)
            if match:
                raw_text = match.group(0)

            return json.loads(raw_text)

        except Exception as e:
            error_str = str(e)
            if "503" in error_str or "UNAVAILABLE" in error_str or "429" in error_str or "RESOURCE_EXHAUSTED" in error_str:
                wait_time = (attempt + 1) * 15  # 遞增等待：5秒、10秒、15秒...
                st.warning(f"⚠️ Gemini 伺服器忙碌中，進行第 {attempt + 1}/{max_retries} 次重試，等待 {wait_time} 秒...")
                time.sleep(wait_time)
            else:
                st.error(f"❌ Gemini API 發生錯誤: {e}")
                break

    st.error("❌ 已達最大重試次數，Gemini 伺服器持續忙碌，請稍後再試。")
    return None

# 開始執行按鈕
if st.button("🚀 開始處理任務", type="primary"):
    if not uploaded_file:
        st.error("❌ 請先上傳 Excel 檔案！")
    elif run_gemini and not gemini_api_key:
        st.error("❌ 請填寫 Gemini API Key！")
    elif run_tavily and not tavily_api_key:
        st.error("❌ 請填寫 Tavily API Key！")
    else:
        st.info("🔄 正在讀取與處理資料...")
        
        # 初始化 API Client
        client = genai.Client(api_key=gemini_api_key) if run_gemini else None
        tavily_client = TavilyClient(api_key=tavily_api_key) if run_tavily else None

        # 讀取 Excel
        input_sheet_name = f"指令{target_group}參數"
        try:
            snippet_sheet = pd.read_excel(uploaded_file, sheet_name='共用片段庫')
            instruction_sheet = pd.read_excel(uploaded_file, sheet_name='指令模板')
            input_fields_sheet = pd.read_excel(uploaded_file, sheet_name='輸入欄位')
            output_fields_sheet = pd.read_excel(uploaded_file, sheet_name='輸出結構')
            input_df = pd.read_excel(uploaded_file, sheet_name=input_sheet_name)
        except Exception as e:
            st.error(f"❌ 讀取 Excel 分頁失敗，請確認分頁標籤是否包含『{input_sheet_name}』: {e}")
            st.stop()

        # 解析設定
        def get_row(df):
            return df[df["項次"] == target_group].iloc[0] if "項次" in df.columns and not df[df["項次"] == target_group].empty else df.iloc[0]

        row_cmd = get_row(instruction_sheet)
        row_in = get_row(input_fields_sheet)
        row_out = get_row(output_fields_sheet)

        task_mode = str(row_cmd.get("任務模式", "evaluate")).strip().lower()
        global_instruction = str(row_cmd.get("指令內容模板", ""))
        
        global_snippets = ""
        snippet_code = row_cmd.get("套用片段代號", "")
        if pd.notna(snippet_code) and "片段代號" in snippet_sheet.columns:
            matched = snippet_sheet[snippet_sheet["片段代號"] == str(snippet_code).strip()]
            if not matched.empty:
                global_snippets = str(matched.iloc[0].get("片段內容", ""))

        input_field_names = [str(row_in[col]).strip() for col in input_fields_sheet.columns if col != "項次" and pd.notna(row_in[col])]
        
        output_fields_meta = []
        for c in [c for c in output_fields_sheet.columns if c.startswith("欄位名稱")]:
            idx = c.replace("欄位名稱", "")
            if pd.notna(row_out.get(c)):
                output_fields_meta.append({
                    "name": str(row_out.get(c)).strip(),
                    "desc": str(row_out.get(f"欄位說明{idx}", ""))
                })

        output_keys = [m["name"] for m in output_fields_meta]
        fields_desc = ", ".join([f'"{k}"' for k in output_keys])
        field_rules_desc = "\n".join([f"- \"{m['name']}\": {m['desc']}" for m in output_fields_meta if m['desc']])

        processed_results = []
        progress_bar = st.progress(0)
        total_rows = len(input_df)

        for index, row in input_df.iterrows():
            dynamic_kwargs = {f: ("" if pd.isna(row.get(f)) else str(row.get(f)).strip()) for f in input_field_names}
            if not any(dynamic_kwargs.values()):
                continue

            formatted_inst = global_instruction
            for k, v in dynamic_kwargs.items():
                formatted_inst = formatted_inst.replace("{" + k + "}", v)

            if task_mode == "generate":
                tavily_template = str(row_cmd.get("Tavily指令內容模板", "{國家} {城市} {年月}"))
                tavily_formatted = tavily_template
                for k, v in dynamic_kwargs.items():
                    tavily_formatted = tavily_formatted.replace("{" + k + "}", v)
                tavily_prompt = re.sub(r'\{[a-zA-Z0-9_\s\(\)]+\}', '', tavily_formatted).strip()

                web_context = search_web_tavily(tavily_client, tavily_prompt, tavily_max_results) if run_tavily else ""

                pure_prompt = (
                    f"{global_snippets}\n\n{formatted_inst}\n\n"
                    f"欄位規範：\n{field_rules_desc}\n\n"
                    f"請回傳 JSON 陣列，欄位需包含：[{fields_desc}]"
                )
                actual_send = f"搜尋參考資料：\n{web_context}\n\n" + pure_prompt if web_context else pure_prompt
                res = call_ai_engine(client, actual_send, model_name, is_batch=True) if run_gemini else None

            else:
                tavily_prompt = "未支援 (Evaluate 模式)"
                web_context = ""
                pure_prompt = (
                    f"{global_snippets}\n\n{formatted_inst}\n\n"
                    f"欄位規範：\n{field_rules_desc}\n\n"
                    f"請回傳 JSON 物件，欄位需包含：[{fields_desc}]"
                )
                res = call_ai_engine(client, pure_prompt, model_name, is_batch=False) if run_gemini else None

            res_list = res if isinstance(res, list) else ([res] if res else [{}])

            for item in res_list:
                row_dict = {}
                for k in output_keys:
                    k_clean = k.replace("\n", " ").strip().lower()
                    if k == "狀態":
                        row_dict[k] = "顯示"
                    elif k == "第一階段搜尋結果":
                        row_dict[k] = web_context if run_tavily else "未執行"
                    elif "tavily" in k_clean and "prompt" in k_clean:
                        row_dict[k] = tavily_prompt
                    elif "gemini" in k_clean and "prompt" in k_clean:
                        row_dict[k] = pure_prompt
                    elif k in dynamic_kwargs:
                        row_dict[k] = dynamic_kwargs[k]
                    else:
                        val = ""
                        if isinstance(item, dict):
                            for ak, av in item.items():
                                if str(ak).replace(" ", "").lower() == k.replace(" ", "").lower():
                                    val = av
                                    break
                        row_dict[k] = str(val).strip() if val not in [None, "nan", "NaN"] else ""
                processed_results.append(row_dict)

            progress_bar.progress((index + 1) / total_rows)
            # 加入 1.5 秒間隔，保護 API 頻率不被鎖住
            time.sleep(3)

        if processed_results:
            df_out = pd.DataFrame(processed_results)
            
            # 使用 BytesIO 在記憶體中建構 Excel 檔案，避免實體硬碟寫入權限問題
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='AI通用清單')
            excel_buffer.seek(0)

            # 將結果與檔名寫入 session_state
            st.session_state["processed_data"] = excel_buffer.getvalue()
            st.session_state["output_filename"] = f"AI產出結果_{input_sheet_name}.xlsx"
            st.success("🎉 處理完成！")

# 繪製下載按鈕（獨立於處理邏輯外，避免點擊下載時按鈕消失）
if "processed_data" in st.session_state:
    st.download_button(
        label="📥 下載處理結果 Excel 檔案",
        data=st.session_state["processed_data"],
        file_name=st.session_state["output_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
