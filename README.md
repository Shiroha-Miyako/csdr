# CSRC Filing Agent

一个用于处理证监会境外发行上市备案材料的 Streamlit 工具。

## 功能

- 上传主表 Excel
- 上传补充材料要求 DOCX
- 上传官网备案情况表 XLSX
- 自动同步新增/更新企业信息
- 自动更新“已完成备案”状态和备案完成时间
- 自动按公司拆分补充材料问题
- 按“一、二、三……”大问题切分，不拆（1）（2）（3）
- 根据历史问题库和关键词规则生成“问题概括”和“大类”
- 导出处理后的 Excel 和本次问题明细 CSV

## 本次修正版解决的问题

- 固定 Python 依赖版本，降低 Streamlit Cloud 兼容性问题
- 增加 `runtime.txt`，建议云端使用 Python 3.11
- 增加 Excel 文件有效性检查，避免上传了无效 `.xlsx` 后直接报错
- 增加 `safe_load_workbook()`，给出更清晰的错误提示
- 增加 `.gitignore`，避免真实 Excel、DOCX、API Key 被上传到 GitHub

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 部署

1. 把本项目上传到 GitHub。
2. Streamlit Cloud 选择该 GitHub 仓库。
3. Main file path 填：

```text
app.py
```

4. 如果可以选择 Python 版本，选：

```text
3.11
```

## API Key 放哪里

当前版本不强制需要 API。后续如果接 OpenAI / DeepSeek，不要写进代码。

本地开发时，新建 `.streamlit/secrets.toml`：

```toml
OPENAI_API_KEY = "your_key_here"
DEEPSEEK_API_KEY = "your_key_here"
```

Streamlit Cloud 部署时，在 App 的 Settings → Secrets 里填写同样内容。

## 不要上传到 Public GitHub 的内容

- 真实主表 Excel
- 真实补充材料 DOCX
- 处理结果 output.xlsx
- `.streamlit/secrets.toml`
- API Key

