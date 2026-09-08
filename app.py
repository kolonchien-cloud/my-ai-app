import io
import json
import re
import time
from google import genai
from google.genai import types
import openpyxl
import pandas as pd
import streamlit as st
from tavily import TavilyClient

# 網頁版面配置
st.set_page_config(
    page_title="AI 聯網自動化生成工具", page_icon="🤖", layout="wide"
)

st.title("🤖 AI 聯網自動化與結構化生成工具")
st.write("透過網頁上傳您的設定 Excel 檔，動態執行 Tavily 搜尋與 Gemini AI 生成。")

# 側邊欄：API 金鑰與控制開關
with st.sidebar:
  st.header("⚙️ 參數與金鑰設定")
  gemini_api_key = st.text_input("Gemini API Key", type="password")
  tavily_api_key = st.text_input("Tavily API Key", type="password")

  model_name = st.selectbox(
      "選擇 Gemini 模型", ["gemini-3.8-flash", "gemini-3.6-flash"]
  )
  tavily_max_results = st.slider("Tavily 搜尋最大回傳筆數", 1, 20, 10)

  st.divider()
  run_tavily = st.checkbox("執行 Tavily 搜尋", value=True)
  run_gemini = st.checkbox("執行 Gemini 生成", value=True)

# 主畫面：上傳 Excel 設定檔
uploaded_file = st.file_uploader(
    "📂 請上傳共用 Excel 檔案 (config_universal.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
  try:
    # 讀取上傳的 Excel
    snippet_sheet = pd.read_excel(uploaded_file, sheet_name="共用片段庫")
    instruction_sheet = pd.read_excel(uploaded_file, sheet_name="指令模板")
    input_fields_sheet = pd.read_excel(uploaded_file, sheet_name="輸入欄位")
    output_fields_sheet = pd.read_excel(uploaded_file, sheet_name="輸出結構")

    # 選擇目標組別
    available_groups = (
        instruction_sheet["項次"].tolist()
        if "項次" in instruction_sheet.columns
        else [1]
    )
    target_group = st.selectbox(
        "🎯 選擇目標組別/項次",
        available_groups,
        format_func=lambda x: f"第 {x} 組",
    )

    input_sheet_name = (
        "指令1參數" if target_group == 1 else f"指令{target_group}參數"
    )

    # 讀取對應分頁的輸入資料
    input_df = pd.read_excel(uploaded_file, sheet_name=input_sheet_name)
    st.success(
        f"成功讀取設定！當前指定分頁：【{input_sheet_name}】，共 {len(input_df)}"
        " 筆輸入資料。"
    )

    # 執行按鈕
    if st.button("🚀 開始執行自動化任務", type="primary"):
      if run_gemini and not gemini_api_key:
        st.error("❌ 請在側邊欄輸入 Gemini API Key！")
      elif run_tavily and not tavily_api_key:
        st.error("❌ 請在側邊欄輸入 Tavily API Key！")
      else:
        # 初始化客戶端
        client = (
            genai.Client(api_key=gemini_api_key)
            if run_gemini and gemini_api_key
            else None
        )
        tavily_client = (
            TavilyClient(api_key=tavily_api_key)
            if run_tavily and tavily_api_key
            else None
        )

        # 獲取對應組別設定
        def get_row_by_item_id(df, item_id):
          if "項次" not in df.columns:
            return df.iloc[0]
          matched = df[df["項次"] == item_id]
          return matched.iloc[0] if not matched.empty else df.iloc[0]

        row_cmd = get_row_by_item_id(instruction_sheet, target_group)
        row_in = get_row_by_item_id(input_fields_sheet, target_group)
        row_out = get_row_by_item_id(output_fields_sheet, target_group)

        global_instruction = str(row_cmd.get("指令內容模板", ""))
        snippet_code = row_cmd.get("套用片段代號", "")
        global_snippets = ""
        if pd.notna(snippet_code) and "片段代號" in snippet_sheet.columns:
          matched_snippet = snippet_sheet[
              snippet_sheet["片段代號"] == str(snippet_code).strip()
          ]
          if not matched_snippet.empty:
            global_snippets = str(
                matched_snippet.iloc[0].get("片段內容", "")
            )

        input_field_names = [
            str(row_in[col]).strip()
            for col in input_fields_sheet.columns
            if col != "項次" and pd.notna(row_in[col]) and str(row_in[col]).strip()
        ]

        output_fields_meta = []
        for c_name_col in [
            c for c in output_fields_sheet.columns if c.startswith("欄位名稱")
        ]:
          idx_str = c_name_col.replace("欄位名稱", "")
          name_val = row_out.get(c_name_col, "")
          if pd.notna(name_val) and str(name_val).strip():
            output_fields_meta.append({
                "name": str(name_val).strip(),
                "desc": str(row_out.get(f"欄位說明{idx_str}", "")),
            })

        output_keys = [meta["name"] for meta in output_fields_meta]
        fields_desc = ", ".join([f'"{k}"' for k in output_keys])
        field_rules_desc = "\n".join(
            [f'- "{m["name"]}": {m["desc"]}' for m in output_fields_meta if m["desc"]]
        )

        processed_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        total_rows = len(input_df)
        for index, row in input_df.iterrows():
          status_text.text(
              f"正在處理第 {index + 1} / {total_rows} 筆資料..."
          )
          dynamic_kwargs = {
              field: (
                  "" if pd.isna(row.get(field)) else str(row.get(field))
              )
              for field in input_field_names
          }

          if not any(dynamic_kwargs.values()):
            continue

          # Tavily 搜尋組裝
          tavily_template = str(
              row_cmd.get("Tavily指令內容模板", "{國家} {城市} {年月}")
          )
          if pd.isna(tavily_template) or not tavily_template.strip():
            tavily_template = "{國家} {城市} {年月}"

          tavily_formatted = tavily_template
          for k, v in dynamic_kwargs.items():
            tavily_formatted = tavily_formatted.replace("{" + k + "}", v)
          tavily_prompt = (
              re.sub(r"\s+", " ", tavily_formatted).replace("．", "").strip()
          )

          web_context = ""
          if run_tavily and tavily_client:
            try:
              response = tavily_client.search(
                  query=tavily_prompt,
                  max_results=tavily_max_results,
                  search_depth="basic",
                  include_answer=True,
              )
              results = response.get("results", [])
              ai_answer = response.get("answer", "").strip()
              if results or ai_answer:
                context_parts = []
                if ai_answer:
                  context_parts.append(
                      f"【Tavily 綜合搜尋摘要】:\n{ai_answer}\n"
                  )
                if results:
                  context_parts.append("【來源網頁詳細資料】:")
                  for r in results:
                    context_parts.append(
                        f"標題: {r.get('title', '')}\n摘要:"
                        f" {r.get('content', '')}\n網址: {r.get('url', '')}\n"
                    )
                web_context = "\n".join(context_parts)
            except Exception as e:
              st.warning(f"Tavily 搜尋發生錯誤: {e}")

          # Gemini 提示詞組裝
          formatted_inst = global_instruction
          for k, v in dynamic_kwargs.items():
            formatted_inst = formatted_inst.replace("{" + k + "}", v)

          pure_gemini_prompt = (
              f"{global_snippets}\n\n{formatted_inst}\n\n各欄位規範：\n{field_rules_desc}\n\n【絕對指令】：請根據上述提供的真實聯網資料，嚴格萃取並整理出符合格式的標準"
              f" JSON 陣列（List of Objects）。絕對不要包含任何前言後記與 Markdown"
              f" 程式碼區塊。每個物件必須包含且僅包含以下 Key：[{fields_desc}]。"
          )
          actual_gemini_send = (
              f"以下是從網路上即時搜尋到的真實參考資料：\n{web_context}\n\n"
              + pure_gemini_prompt
              if web_context
              else pure_gemini_prompt
          )

          res = None
          if run_gemini and client:
            try:
              config = types.GenerateContentConfig(
                  response_mime_type="application/json"
              )
              response = client.models.generate_content(
                  model=model_name, contents=actual_gemini_send, config=config
              )
              raw_text = response.text.strip()
              if "```" in raw_text:
                raw_text = (
                    re.sub(r"```(?:json)?", "", raw_text)
                    .replace("```", "")
                    .strip()
                )
              match = re.search(r"\[.*\]", raw_text, re.DOTALL)
              if match:
                raw_text = match.group(0)
              res = json.loads(raw_text)
            except Exception as e:
              st.error(f"Gemini 執行錯誤: {e}")

          # 結果組織
          if run_gemini and res is not None:
            res_list = res if isinstance(res, list) else [res]
            for item in res_list:
              row_dict = {}
              for k in output_keys:
                k_clean = k.replace("\n", " ").strip()
                if k == "狀態":
                  row_dict[k] = "顯示"
                elif k == "第一階段搜尋結果":
                  row_dict[k] = web_context if run_tavily else "未執行"
                elif "tavily" in k_clean.lower() and "prompt" in k_clean.lower():
                  row_dict[k] = tavily_prompt
                elif "gemini" in k_clean.lower() and "prompt" in k_clean.lower():
                  row_dict[k] = pure_gemini_prompt
                elif k in dynamic_kwargs:
                  row_dict[k] = dynamic_kwargs[k]
                else:
                  val = ""
                  if isinstance(item, dict):
                    if k in item:
                      val = item[k]
                    else:
                      for actual_key, actual_val in item.items():
                        if actual_key.replace(" ", "") == k.replace(" ", ""):
                          val = actual_val
                          break
                  row_dict[k] = (
                      ""
                      if val is None or str(val).lower() == "nan"
                      else str(val).strip()
                  )
              processed_results.append(row_dict)
          else:
            row_dict = {}
            for k in output_keys:
              row_dict[k] = dynamic_kwargs.get(k, "未執行")
            processed_results.append(row_dict)

          progress_bar.progress((index + 1) / total_rows)
          time.sleep(1)

        status_text.text("✅ 所有任務執行完畢！")

        if processed_results:
          out_df = pd.DataFrame(processed_results)
          st.subheader("📊 執行結果預覽")
          st.dataframe(out_df)

          # 產生下載按鈕
          output_buffer = io.BytesIO()
          with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="AI通用清單")
          output_buffer.seek(0)

          st.download_button(
              label="📥 下載產出的 Excel 報表",
              data=output_buffer,
              file_name=f"AI通用產出結果_第{target_group}組.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

  except Exception as e:
    st.error(f"❌ 讀取或處理 Excel 檔案時發生錯誤: {e}")
