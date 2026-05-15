# CSRC 境外上市备案表自动处理 Agent - Starter

## 你要做的三件事
1. 备案通知书公告：把总表里对应公司的“备案状态/补充材料”改成“已完成备案”，并填写完成备案时间。
2. 补充材料要求公告：解析 docx，把每家公司的一、二、三大问题写回总表 Sheet1 的问题1-问题9；同时追加到问题库 Sheet2/Sheet3，用历史问题库做标准化概括和大类分类。
3. 备案情况表公告：读取官网新 xlsx，与旧总表按企业名称匹配，把新增企业或更新信息同步进总表。

## 安装
```bash
cd csrc_agent_starter
python -m venv .venv
# Windows
.venv\Scripts\activate
# Mac/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## 运行网页工具
```bash
streamlit run app.py
```

## 直接命令行运行示例
```bash
python csrc_agent.py \
  --master "境内企业境外发行证券和上市备案情况表（截至20260508)(1)(1).xlsx" \
  --supp-doc "境外发行上市备案补充材料要求公示（2026年4月27日—2026年5月8日）.docx" \
  --filing-xlsx "境内企业境外发行证券和上市备案情况表（首次公开发行及全流通）（截至2026年5月8日）.xlsx" \
  --supp-date 2026-05-08 \
  --out output.xlsx
```

## 工作原则
- 补充材料按“一、二、三……”大问题切割，不拆（1）（2）（3）。
- 原问题列尽量保留监管原文。
- 问题概括优先复用历史问题库里相似问题的标准表达。
- 大类优先复用历史问题库分类。
- API 只是辅助，第一版先用“关键词 + 历史相似问题匹配”跑通。
