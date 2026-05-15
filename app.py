import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from csrc_agent import (
    parse_supplement_doc,
    sync_filing_xlsx,
    update_master_with_issues,
    update_completed_filings,
)

st.set_page_config(page_title="CSRC 备案表自动处理 Agent", layout="wide")
st.title("CSRC 境外上市备案表自动处理 Agent")
st.caption("第一版：先跑通自动解析、匹配、写表；复杂分类后续再接 API。")

st.sidebar.header("输入文件")
master = st.sidebar.file_uploader("1）你的总表 xlsx", type=["xlsx"])
supp_doc = st.sidebar.file_uploader("2）补充材料要求 docx", type=["docx"])
filing_xlsx = st.sidebar.file_uploader("3）官网备案情况表 xlsx", type=["xlsx"])
supp_date = st.sidebar.text_input("补充材料公告日期", value="2026-05-08")

st.sidebar.markdown("---")
nst_completed = st.sidebar.text_area(
    "已完成备案公司（可选，每行：公司名,日期）",
    value="广州豪特节能环保科技股份有限公司,2026-05-09\n华健未来（成都）科技股份有限公司,2026-05-09\n南京海纳医药科技股份有限公司,2026-05-09",
    height=120,
)

run = st.sidebar.button("开始处理", type="primary")

if run:
    if not master:
        st.error("请先上传你的总表 xlsx。")
        st.stop()

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        master_path = td / master.name
        master_path.write_bytes(master.getvalue())
        current_path = master_path

        # Step 1: 同步官网备案情况表
        if filing_xlsx:
            fx = td / filing_xlsx.name
            fx.write_bytes(filing_xlsx.getvalue())
            synced_path = td / "01_synced_filing.xlsx"
            changes = sync_filing_xlsx(str(current_path), str(fx), str(synced_path))
            current_path = synced_path
            st.success(f"备案情况表同步完成：{len(changes)} 处新增/更新")
            if changes:
                st.dataframe(pd.DataFrame(changes, columns=["动作", "企业名称", "列号"]))

        # Step 2: 已完成备案更新
        completed_items = []
        for line in nst_completed.splitlines():
            line = line.strip()
            if not line or "," not in line:
                continue
            company, date = line.split(",", 1)
            completed_items.append((company.strip(), date.strip()))
        if completed_items:
            completed_path = td / "02_completed.xlsx"
            update_completed_filings(str(current_path), completed_items, str(completed_path))
            current_path = completed_path
            st.success(f"已完成备案状态更新完成：{len(completed_items)} 家")

        # Step 3: 补充材料 docx 解析和写回
        issues_df = None
        if supp_doc:
            sp = td / supp_doc.name
            sp.write_bytes(supp_doc.getvalue())
            issues = parse_supplement_doc(str(sp), str(current_path))
            issues_df = pd.DataFrame([i.__dict__ for i in issues])
            st.success(f"补充材料解析完成：{len(issues)} 个大问题")
            st.dataframe(issues_df, use_container_width=True)

            final_path = td / "output_csrc_agent.xlsx"
            update_master_with_issues(str(current_path), issues, supp_date, str(final_path))
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
