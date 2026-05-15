import argparse
import json
import os
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document
from openpyxl import load_workbook
# 某些证监会表格或手工模板里带有 Excel 图表/绘图 XML，openpyxl 读取时可能报 pitchFamily 错误。
# 这个工具只处理单元格数据，不需要读取图片/图表，因此禁用图片/图表解析。
try:
    import openpyxl.reader.excel as _openpyxl_excel
    _openpyxl_excel.find_images = lambda archive, path: ([], [])
except Exception:
    pass
from rapidfuzz import fuzz, process

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

CHINESE_NUM = "一二三四五六七八九十"

# 你可以持续维护这个词典：优先级高于 AI
KEYWORD_RULES = [
    (["股权架构", "返程并购", "37 号文", "37号文"], "股权架构搭建及返程并购的合规性", "业务合规"),
    (["业务经营", "数字广告", "网络推广", "职业中介", "广播电视节目制作"], "业务经营合规性", "业务合规"),
    (["境内运营实体", "重大诉讼", "行政处罚", "历次股权变动"], "境内运营实体情况", "业务合规"),
    (["高耗能", "高排放", "排污", "固定污染物"], "已建、在建及此次募投项目是否属于“高耗能”“高排放”项目", "环保"),
    (["全流通", "质押", "冻结", "权利瑕疵"], "全流通股份的瑕疵情形", "全流通"),
    (["国有股东", "国有股标识", "国资"], "国有股东标识办理进展情况", "国有股份"),
    (["股权激励", "员工持股", "外部顾问", "期权激励"], "股权激励计划", "股权激励"),
    (["最近12个月", "新增股东", "入股价格", "利益输送"], "新增股东入股价格的定价依据及合理性", "股东情况"),
    (["股份代持", "股权代持", "代持"], "股份代持情况", "股份代持"),
    (["历次增资", "股权转让", "实缴出资", "抽逃出资", "出资方式"], "历次股权变动的合法合规性", "股权变更"),
    (["持股5%以上", "5%以上", "穿透", "禁止持股"], "持股5%以上股东信息", "股东情况"),
    (["董事", "高级管理人员", "国籍", "永久境外居留权"], "董事、高级管理人员国籍及境外永久居留权情况", "股东情况"),
    (["APP", "小程序", "公众号", "个人信息", "数据收集", "数据安全"], "主要境内运营实体开发、运营的网站、APP等产品情况；个人信息保护、数据安全情况", "数据安全"),
    (["境外子公司", "境外投资", "外汇登记"], "公司境外子公司涉及的境外投资、外汇登记等监管程序具体履行情况说明", "业务合规"),
    (["外商投资准入", "负面清单", "外资禁止", "外资限制"], "公司及下属公司经营范围是否涉及国家禁止或限制外商投资的领域", "业务合规"),
    (["技术进出口", "技术出口"], "最近三年技术出口业务的开展情况及合规性说明", "业务合规"),
    (["募集资金", "募投", "境外投资"], "募集资金用途涉及境外投资所履行的相关审批、核准或备案程序情况", "备案情况"),
    (["发行方案", "招股说明书", "备案材料"], "本次发行上市方案", "上市方案"),
    (["审计意见", "会计师事务所"], "会计师事务所审计意见类型", "备案情况"),
    (["A股上市", "辅导备案", "撤回原因"], "曾经申请A股上市的相关情况以及是否存在对本次发行上市产生重大影响的事项", "历史上市"),
    (["诉讼", "仲裁"], "未决诉讼情况", "未决诉讼"),
    (["安全生产", "行政处罚"], "安全生产法律法规落实情况", "生产安全"),
    (["第八条", "不得境外发行上市", "禁止境外发行上市"], "专项核查（是否存在不得境外发行上市的情形）", "法律合规"),
]



class WorkbookReadError(ValueError):
    """用户上传的 Excel 文件无法被 openpyxl 读取时使用的友好错误。"""


def is_valid_xlsx(path: str) -> bool:
    """xlsx 本质是 zip 包。先用 zipfile 做快速校验，避免云端报红色 ValueError。"""
    try:
        return zipfile.is_zipfile(path)
    except Exception:
        return False


def safe_load_workbook(path: str, **kwargs):
    """
    更稳地读取 Excel 工作簿。
    常见失败原因：文件不是标准 .xlsx、下载成 HTML、.xls 改后缀、文件损坏、外部链接/绘图 XML 兼容问题。
    """
    path = str(path)
    if not is_valid_xlsx(path):
        raise WorkbookReadError(
            "上传的文件不是有效的 .xlsx 工作簿。请用 Excel 或 WPS 打开该文件，"
            "选择“另存为”→“Excel 工作簿 (*.xlsx)”，再重新上传。"
        )

    options = {
        "data_only": False,
        "read_only": False,
        "keep_vba": False,
        "keep_links": False,
    }
    options.update(kwargs)

    try:
        return load_workbook(path, **options)
    except TypeError:
        # 兼容不同 openpyxl 版本的参数差异
        options.pop("rich_text", None)
        try:
            return load_workbook(path, **options)
        except Exception as e:
            raise WorkbookReadError(
                "openpyxl 无法读取该 Excel 文件。请先用 Excel/WPS 打开并另存为新的 .xlsx 文件，"
                "文件名尽量使用英文，例如 master.xlsx、filing_table.xlsx。"
            ) from e
    except Exception as e:
        raise WorkbookReadError(
            "openpyxl 无法读取该 Excel 文件。请先用 Excel/WPS 打开并另存为新的 .xlsx 文件，"
            "文件名尽量使用英文，例如 master.xlsx、filing_table.xlsx。"
        ) from e


@dataclass
class Issue:
    company: str
    original: str
    summary: str = ""
    category: str = ""


def clean_text(s: str) -> str:
    s = (s or "").replace("\u3000", " ").replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s.strip()


def read_docx_text(path: str) -> str:
    doc = Document(path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    text = "\n".join(paras)
    # 去掉独立页码行
    text = re.sub(r"\n\d+\n", "\n", text)
    return clean_text(text)


def split_company_sections(text: str) -> Dict[str, str]:
    """识别 docx 中的公司段落。公司名通常是单独一行，后面跟“请你公司...”"""
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    start_idx = 0
    for i, line in enumerate(lines):
        if "具体如下" in line:
            start_idx = i + 1
            break
    lines = lines[start_idx:]

    sections: Dict[str, List[str]] = {}
    current_company = None
    buf: List[str] = []

    def is_company_line(line: str, next_line: str = "") -> bool:
        if len(line) > 40:
            return False
        if re.search(r"[，。；：、（）()]", line):
            return False
        if line.startswith(tuple(CHINESE_NUM)):
            return False
        if "请你公司" in next_line or "请发行人" in next_line:
            return True
        # 部分公告公司名后一行可能不是“请你公司”，保守兜底
        return bool(re.search(r"(电子|网络|通信|科技|新能源|医疗|医药|控股|集团|股份|有限|公司|生物|材料|智能|环保)$", line)) and len(line) <= 25

    for i, line in enumerate(lines):
        nxt = lines[i+1] if i+1 < len(lines) else ""
        if is_company_line(line, nxt):
            if current_company and buf:
                sections[current_company] = buf[:]
            current_company = line
            buf = []
        elif current_company:
            buf.append(line)
    if current_company and buf:
        sections[current_company] = buf

    return {k: clean_text("\n".join(v)) for k, v in sections.items()}


def remove_intro(section_text: str) -> str:
    # 删除“请你公司补充说明...法律意见”等引导句，保留实质问题
    section_text = re.sub(r"^请你公司(?:就以下事项)?补充说明以下事项[，,]?请律师(?:进行)?核查并出具明确的?法律意见[:：]?", "", section_text)
    section_text = re.sub(r"^请你公司就以下事项补充说明[，,]?请律师(?:进行)?核查并出具明确的?法律意见[:：]?", "", section_text)
    section_text = re.sub(r"^请你公司补充说明以下事项[，,]?请律师进行核查并出具明确的法律意见[:：]?", "", section_text)
    return clean_text(section_text)


def split_big_issues(section_text: str) -> List[str]:
    body = remove_intro(section_text)
    # 标记“一、二、三”大序号；不匹配（1）（2）
    pattern = re.compile(rf"(?<![（(])([{CHINESE_NUM}]{{1,3}})、")
    matches = list(pattern.finditer(body))
    if not matches:
        return [clean_text(body)] if body else []
    issues = []
    for idx, m in enumerate(matches):
        start = m.start()
        end = matches[idx+1].start() if idx + 1 < len(matches) else len(body)
        q = body[start:end]
        q = re.sub(rf"^[{CHINESE_NUM}]{{1,3}}、", "", q).strip()
        issues.append(clean_text(q))
    return [x for x in issues if x]


def load_history_examples(wb) -> List[Tuple[str, str, str]]:
    """从问题类型统计、本次新增、当期新增读取历史：原问题/标准问题/大类。"""
    examples = []
    for ws_name in ["问题类型统计", "本次新增", "当期新增"]:
        if ws_name not in wb.sheetnames:
            continue
        ws = wb[ws_name]
        for row in ws.iter_rows(min_row=1, values_only=True):
            vals = list(row)
            if ws_name == "问题类型统计":
                original, category = vals[0], vals[1]
                summary = vals[0]
            else:
                original = vals[1] if len(vals) > 1 else None
                summary = vals[2] if len(vals) > 2 else None
                category = vals[3] if len(vals) > 3 else None
            if original and summary and category and str(original).strip() not in ["问题", "请说明"]:
                examples.append((str(original).strip(), str(summary).strip(), str(category).strip()))
    return examples


def classify_by_rules(issue_text: str, examples: List[Tuple[str, str, str]]) -> Tuple[str, str, str]:
    # 1) 关键词规则
    for kws, summary, category in KEYWORD_RULES:
        issue_no_space = re.sub(r"\s+", "", issue_text)
        if any(k in issue_text or re.sub(r"\s+", "", k) in issue_no_space for k in kws):
            return summary, category, "keyword"
    # 2) 历史相似问题匹配
    if examples:
        candidates = [x[0] for x in examples]
        best = process.extractOne(issue_text, candidates, scorer=fuzz.token_set_ratio)
        if best and best[1] >= 72:
            idx = best[2]
            return examples[idx][1], examples[idx][2], f"history:{best[1]:.0f}"
    # 3) 兜底
    return "待人工确认", "待分类", "fallback"


def parse_supplement_doc(docx_path: str, master_wb_path: str) -> List[Issue]:
    text = read_docx_text(docx_path)
    sections = split_company_sections(text)
    wb = safe_load_workbook(master_wb_path)
    examples = load_history_examples(wb)

    issues: List[Issue] = []
    for company, sec in sections.items():
        for q in split_big_issues(sec):
            summary, cat, _ = classify_by_rules(q, examples)
            issues.append(Issue(company=company, original=q, summary=summary, category=cat))
    return issues


def find_header_row(ws, header="企业名称") -> int:
    for row in range(1, min(ws.max_row, 20) + 1):
        for col in range(1, ws.max_column + 1):
            if ws.cell(row, col).value == header:
                return row
    return 3


def find_company_row(ws, company: str, name_col: int = 2) -> Optional[int]:
    names = []
    rows = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, name_col).value
        if v:
            names.append(str(v))
            rows.append(r)
    # 公司简称匹配全称：用 partial_ratio
    best = process.extractOne(company, names, scorer=fuzz.partial_ratio)
    if best and best[1] >= 70:
        return rows[best[2]]
    return None


def update_master_with_issues(master_path: str, issues: List[Issue], supp_date: str, out_path: str):
    wb = safe_load_workbook(master_path)
    ws = wb["备案情况总表"] if "备案情况总表" in wb.sheetnames else wb.active
    history_ws = wb["问题类型统计"] if "问题类型统计" in wb.sheetnames else None
    add_ws = wb["本次新增"] if "本次新增" in wb.sheetnames else wb.create_sheet("本次新增")

    by_company: Dict[str, List[Issue]] = {}
    for it in issues:
        by_company.setdefault(it.company, []).append(it)

    # 备案情况总表：K为补充材料公告日期，O问题数量，P开始问题1
    for company, company_issues in by_company.items():
        r = find_company_row(ws, company, name_col=2)
        if not r:
            print(f"[WARN] 总表未找到公司：{company}")
            continue
        ws.cell(r, 11).value = supp_date  # K: 备案补充材料公告日期
        ws.cell(r, 15).value = len(company_issues)  # O: 问题数量
        for i, it in enumerate(company_issues[:9], start=0):
            ws.cell(r, 16+i).value = it.summary  # P-X: 问题1-问题9，填标准化问题名称；如果你要填原文，改成 it.original

    # 本次新增：追加原问题、概括、大类
    start = add_ws.max_row + 1
    for it in issues:
        add_ws.cell(start, 2).value = it.original
        add_ws.cell(start, 3).value = it.summary
        add_ws.cell(start, 4).value = it.category
        start += 1

    # 问题类型统计：追加“标准问题 + 大类”用于下次复用
    if history_ws:
        hrow = history_ws.max_row + 1
        for it in issues:
            history_ws.cell(hrow, 1).value = it.summary
            history_ws.cell(hrow, 2).value = it.category
            history_ws.cell(hrow, 3).value = 1
            hrow += 1

    wb.save(out_path)
    return out_path


def update_completed_filings(master_path: str, completed_items: List[Tuple[str, str]], out_path: str):
    """completed_items: [(公司名, 完成备案日期)]"""
    wb = safe_load_workbook(master_path)
    ws = wb["备案情况总表"] if "备案情况总表" in wb.sheetnames else wb.active
    for company, completed_date in completed_items:
        r = find_company_row(ws, company, name_col=2)
        if r:
            ws.cell(r, 9).value = "已完成备案"   # I: 备案状态
            ws.cell(r, 12).value = completed_date # L: 已完成备案时间
        else:
            print(f"[WARN] 未找到完成备案公司：{company}")
    wb.save(out_path)
    return out_path


def sync_filing_xlsx(master_path: str, new_filing_xlsx: str, out_path: str):
    """把官网最新备案情况表中的新增企业或更新信息同步到主表。"""
    master_wb = safe_load_workbook(master_path)
    master_ws = master_wb["备案情况总表"] if "备案情况总表" in master_wb.sheetnames else master_wb.active
    new_wb = safe_load_workbook(new_filing_xlsx, data_only=False)
    new_ws = new_wb.active

    # 新表列 A-J，对应主表 A-I 基本信息。主表 C-I 与新表 C-I 基本一致。
    master_names = {str(master_ws.cell(r, 2).value).strip(): r for r in range(5, master_ws.max_row+1) if master_ws.cell(r, 2).value}
    append_row = master_ws.max_row + 1
    changed = []

    for r in range(5, new_ws.max_row+1):
        name = new_ws.cell(r, 2).value
        if not name:
            continue
        name_s = str(name).strip()
        # 精确或模糊匹配
        if name_s in master_names:
            mr = master_names[name_s]
        else:
            best = process.extractOne(name_s, list(master_names.keys()), scorer=fuzz.token_set_ratio) if master_names else None
            mr = master_names[best[0]] if best and best[1] >= 92 else None
        if mr:
            # 更新 C-I 基础字段；不覆盖后面的补充材料和问题列
            for c in range(3, 10):
                old = master_ws.cell(mr, c).value
                new = new_ws.cell(r, c).value
                if new is not None and old != new:
                    master_ws.cell(mr, c).value = new
                    changed.append(("update", name_s, c))
        else:
            # 追加新企业，复制 A-I
            for c in range(1, 10):
                master_ws.cell(append_row, c).value = new_ws.cell(r, c).value
            changed.append(("append", name_s, None))
            append_row += 1

    master_wb.save(out_path)
    return changed


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--master", required=True, help="你的总表 xlsx")
    parser.add_argument("--supp-doc", help="补充材料要求 docx")
    parser.add_argument("--filing-xlsx", help="官网备案情况表 xlsx")
    parser.add_argument("--supp-date", default=datetime.today().strftime("%Y-%m-%d"))
    parser.add_argument("--out", default="output.xlsx")
    args = parser.parse_args()

    current = args.master
    if args.filing_xlsx:
        tmp = args.out.replace(".xlsx", "_synced.xlsx")
        changed = sync_filing_xlsx(current, args.filing_xlsx, tmp)
        print(f"同步备案情况表完成：{len(changed)} 处新增/更新")
        current = tmp

    if args.supp_doc:
        issues = parse_supplement_doc(args.supp_doc, current)
        print(f"解析补充材料完成：{len(issues)} 个大问题")
        for it in issues[:10]:
            print(it.company, "|", it.summary, "|", it.category)
        update_master_with_issues(current, issues, args.supp_date, args.out)
        print(f"已输出：{args.out}")

if __name__ == "__main__":
    main()
