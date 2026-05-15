import os
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from csrc_agent import (
    WorkbookReadError,
    extract_supplement_dates_from_filename,
    format_date_slash,
    is_valid_xlsx,
    parse_supplement_doc,
    sync_filing_xlsx,
    update_completed_filings,
    update_master_with_issues,
)

# 把 Streamlit Cloud Secrets 写入环境变量，供 csrc_agent.py 中的 OpenAI SDK 读取
for _key in ["OPENAI_API_KEY", "DEEPSEEK_API_KEY", "OPENAI_MODEL", "DEEPSEEK_MODEL"]:
    try:
        if _key in st.secrets:
            os.environ[_key] = str(st.secrets[_key])
    except Exception:
        pass

st.set_page_config(page_title="CSRC 备案表自动处理 Agent", layout="wide")
st.title("CSRC 境外上市备案表自动处理 Agent")
st.caption("API版：规则/历史库优先，低置信度问题可调用 OpenAI 或 DeepSeek 进行分类。")

with st.expander("使用说明", expanded=False):
    st.markdown(
        """
        1. 上传你的主表 Excel（必须是标准 `.xlsx`，不要上传 `.xls` 或网页下载失败的 HTML 文件）。  
        2. 可选上传官网备案情况表 `.xlsx`，用于同步新增/更新企业。  
        3. 可选上传补充材料要求 `.docx`，用于按公司拆分“一、二、三……”大问题。  
        4. 已完成备案公司按 `公司名,日期` 每行一个填写。  
        5. 如果 Excel 读取失败，请先用 Excel/WPS 打开后“另存为 Excel 工作簿 .xlsx”，再上传。
        """
    )

st.sidebar.header("输入文件")
master = st.sidebar.file_uploader("1）你的总表 xlsx", type=["xlsx"])
supp_doc = st.sidebar.file_uploader("2）补充材料要求 docx（可选）", type=["docx"])
filing_xlsx = st.sidebar.file_uploader("3）官网备案情况表 xlsx（可选）", type=["xlsx"])
st.sidebar.caption("补充材料公告当周 / 公告日期会优先从 Word 文件名自动提取。")

st.sidebar.markdown("---")
st.sidebar.header("API 分类设置")
use_llm = st.sidebar.checkbox("启用 API 智能分类（仅低置信度问题调用）", value=False)
provider = st.sidebar.selectbox("模型服务商", ["openai", "deepseek"], index=0)
llm_threshold = st.sidebar.slider("历史匹配分数低于多少时调用 API", min_value=50, max_value=95, value=80, step=5)

if use_llm:
    if provider == "openai" and not os.getenv("OPENAI_API_KEY"):
        st.sidebar.warning("未检测到 OPENAI_API_KEY。请在 Streamlit Secrets 中配置。")
    if provider == "deepseek" and not os.getenv("DEEPSEEK_API_KEY"):
        st.sidebar.warning("未检测到 DEEPSEEK_API_KEY。请在 Streamlit Secrets 中配置。")


st.sidebar.markdown("---")
nst_completed = st.sidebar.text_area(
    "已完成备案公司（可选，每行：公司名,日期）",
    value="广州豪特节能环保科技股份有限公司,2026/5/9\n华健未来（成都）科技股份有限公司,2026/5/9\n南京海纳医药科技股份有限公司,2026/5/9",
    height=120,
)

run = st.sidebar.button("开始处理", type="primary")


def save_upload(uploaded_file, target_path: Path) -> Path:
    target_path.write_bytes(uploaded_file.getvalue())
    return target_path


def validate_uploaded_xlsx(path: Path, label: str):
    if not is_valid_xlsx(str(path)):
        st.error(
            f"{label} 不是有效的 .xlsx 文件。\n\n"
            "请用 Excel/WPS 打开后，另存为“Excel 工作簿 (*.xlsx)”，文件名尽量改成英文后重新上传。"
        )
        st.stop()


def parse_completed_items(raw_text: str):
    items = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line or "," not in line:
            continue
        company, date = line.split(",", 1)
        company, date = company.strip(), format_date_slash(date.strip())
        if company and date:
            items.append((company, date))
    return items


if run:
    if not master:
        st.error("请先上传你的总表 xlsx。")
        st.stop()

    try:
        with tempfile.TemporaryDirectory() as td_raw:
            td = Path(td_raw)

            master_path = td / "master.xlsx"
            save_upload(master, master_path)
            validate_uploaded_xlsx(master_path, "主表 Excel")
            current_path = master_path

            # Step 1: 同步官网备案情况表
            if filing_xlsx:
                fx = td / "filing_table.xlsx"
                save_upload(filing_xlsx, fx)
                validate_uploaded_xlsx(fx, "官网备案情况表 Excel")

                synced_path = td / "01_synced_filing.xlsx"
                changes = sync_filing_xlsx(str(current_path), str(fx), str(synced_path))
                current_path = synced_path
                st.success(f"备案情况表同步完成：{len(changes)} 处新增/更新")
                if changes:
                    st.dataframe(pd.DataFrame(changes, columns=["动作", "企业名称", "列号"]), use_container_width=True)

            # Step 2: 已完成备案更新
            completed_items = parse_completed_items(nst_completed)
            if completed_items:
                completed_path = td / "02_completed.xlsx"
                update_completed_filings(str(current_path), completed_items, str(completed_path))
                current_path = completed_path
                st.success(f"已完成备案状态更新完成：{len(completed_items)} 家")

            # Step 3: 补充材料 docx 解析和写回
            issues_df = None
            if supp_doc:
                supp_week_start, supp_notice_date = extract_supplement_dates_from_filename(supp_doc.name)
                if supp_week_start and supp_notice_date:
                    st.info(f"从 Word 文件名识别到：备案补充材料公告当周 = {supp_week_start}，备案补充材料公告日期 = {supp_notice_date}")
                else:
                    st.warning("未能从 Word 文件名识别日期。请确认文件名类似：境外发行上市备案补充材料要求公示（2026年4月27日—2026年5月8日）.docx")

                sp = td / "supplement.docx"
                save_upload(supp_doc, sp)

                issues = parse_supplement_doc(str(sp), str(current_path), use_llm=use_llm, provider=provider, llm_threshold=llm_threshold)
                issues_df = pd.DataFrame([i.__dict__ for i in issues])
                st.success(f"补充材料解析完成：{len(issues)} 个大问题" + (f"；API分类：{provider}" if use_llm else "；未启用API分类"))
                if not issues_df.empty:
                    st.dataframe(issues_df, use_container_width=True)

                final_path = td / "output_csrc_agent.xlsx"
                update_master_with_issues(
                    str(current_path),
                    issues,
                    supp_notice_date,
                    str(final_path),
                    supp_week_start=supp_week_start,
                )
                current_path = final_path

            st.markdown("### 下载结果")
            st.download_button(
                "下载处理后的 xlsx",
                data=Path(current_path).read_bytes(),
                file_name="output_csrc_agent.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            if issues_df is not None:
                st.download_button(
                    "下载本次问题明细 CSV",
                    data=issues_df.to_csv(index=False, encoding="utf-8-sig"),
                    file_name="issues_detail.csv",
                    mime="text/csv",
                )

    except WorkbookReadError as e:
        st.error(str(e))
        st.info("建议：把两个 Excel 都用 Excel/WPS 打开，另存为新的 .xlsx 文件，文件名改成 master.xlsx / filing_table.xlsx 后再上传。")
        st.stop()
    except zipfile.BadZipFile:
        st.error("上传的 Excel 文件结构损坏，无法作为 .xlsx 读取。请重新下载或另存为新的 .xlsx 文件。")
        st.stop()
    except Exception as e:
        st.error("处理失败。请先确认上传的是标准 .xlsx / .docx 文件；如果仍失败，把 Manage app 里的完整日志发给我。")
        with st.expander("显示错误详情"):
            st.exception(e)
        st.stop()
else:
    st.info("请在左侧上传文件，然后点击“开始处理”。")
