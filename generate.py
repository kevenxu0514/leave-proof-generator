# -*- coding: utf-8 -*-
"""请假证明一键生成 Skill v1.0（乙方案：每一页末尾都出现落款）

用法:
    python generate.py <输入文件> [--out <输出目录>] [--name <输出文件名>]

输入支持三种载体（含 名单 + 活动信息）：
    .xlsx  Sheet「名单」(列头=模板列头) + Sheet「活动信息」(键值对: template/event_date/periods/
           location/activity/dorm/special_note/sign_date)
    .csv   名单文件（列头=模板列头），同目录下的 活动信息.csv 存放键值对
    .json  单文件含全部字段（见 样例数据/*.json）

流程：读取输入 → 复制模板 → 填正文黄高亮槽位 → 填名单表格 →
      乙方案分段（每页=表头+行+2空行+落款组+分页符）→ 渲染核验 → 容量自适应收敛 → 输出 docx
"""
import argparse
import json
import math
import os
import re
import subprocess
import sys
from copy import deepcopy
from datetime import date, timedelta

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls
import pymupdf

WORK = os.path.dirname(os.path.abspath(__file__))
def _slot_ranges_sk(full):
    """上课请假：固定前缀 | S1 日期(星期) | S2 节次 | S3 到{地点}出席{活动}，志愿者活动， | 固定后缀"""
    m1 = re.search(r'\d{4}年\d{1,2}月\d{1,2}日\(星期[一二三四五六日]\)', full)
    m3 = re.search(r'到.+?，志愿者活动，', full)
    if not m1 or not m3:
        raise RuntimeError(f'上课模板正文槽位定位失败: {full!r}')
    return [(m1.start(), m1.end()), (m1.end(), m3.start()), (m3.start(), m3.end())]


def _slot_ranges_fs(full):
    """分寝室：固定 兹证明 | S1 宿舍楼 | 固定 下列名单中同学，因参加 | S2 日期（星期） | S3 参加…， | 固定后缀"""
    m1 = re.search(r'[一二三四五六七八九十]+舍', full)
    m2 = re.search(r'\d{4}年\d{1,2}月\d{1,2}日（星期[一二三四五六日]）', full)
    end3 = full.find('故需请假，特此证明！')
    if not m1 or not m2 or end3 < 0:
        raise RuntimeError(f'分寝室模板正文槽位定位失败: {full!r}')
    return [(m1.start(), m1.end()), (m2.start(), m2.end()), (m2.end(), end3)]


TEMPLATES = {
    '上课请假': dict(
        file='上课请假模板.docx',
        headers=['学院', '姓名', '班级', '学号'],
        required=['event_date', 'periods', 'location', 'activity'],
        slot_ranges=_slot_ranges_sk,
    ),
    '分寝室请假': dict(
        file='分寝室请假模板.docx',
        headers=['姓名', '班级', '学号', '寝室号'],
        required=['event_date', 'dorm', 'activity', 'special_note'],
        slot_ranges=_slot_ranges_fs,
    ),
}
BLANK_LINES = 2
MAX_ROUNDS = 6
WEEKDAY_CN = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']


# ---------------- 输入读取 ----------------

def read_input(path):
    ext = path.rsplit('.', 1)[-1].lower()
    if ext == 'json':
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        return data.get('info', data), data.get('headers') or data.get('header'), data.get('students', [])
    if ext in ('xlsx', 'xls'):
        return _read_xlsx(path)
    if ext == 'csv':
        return _read_csv(path)
    raise RuntimeError(f'不支持的输入类型: {ext}（支持 xlsx/csv/json）')


def _read_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    names = wb.sheetnames
    sheet_n = next((s for s in names if '名单' in s or '学生' in s), names[0])
    sheet_i = next((s for s in names if '活动' in s or '信息' in s), None)
    rows = [r for r in wb[sheet_n].iter_rows(values_only=True)]
    header = [str(c).strip() for c in rows[0] if c is not None]
    students = [list(r) for r in rows[1:] if any(c is not None and str(c).strip() for c in r)]
    info = {}
    if sheet_i:
        for r in wb[sheet_i].iter_rows(values_only=True):
            if r[0] is None or str(r[0]).strip() == '':
                continue
            k = str(r[0]).strip()
            v = r[1] if len(r) > 1 else None
            if v is not None:
                info[k] = str(v).strip() if not isinstance(v, (int, float)) else str(v)
    return info, header, students


def _read_csv(path):
    import csv
    with open(path, encoding='utf-8-sig') as f:
        rows = [r for r in csv.reader(f)]
    header = [c.strip() for c in rows[0]]
    students = [r for r in rows[1:] if any(c.strip() for c in r)]
    info = {}
    import glob
    d = os.path.dirname(path) or '.'
    base = os.path.basename(path)
    prefix = base.split('-名单')[0] if '-名单' in base else None
    cands = []
    if prefix:
        cands = sorted(glob.glob(os.path.join(d, f'活动信息-{prefix}*.csv')))
    if not cands:
        cands = sorted(glob.glob(os.path.join(d, '活动信息*.csv')))
    seen = set()
    for info_path in cands:
        if info_path in seen:
            continue
        seen.add(info_path)
        with open(info_path, encoding='utf-8-sig') as f:
            for r in csv.reader(f):
                if r and r[0].strip():
                    info[r[0].strip()] = r[1].strip() if len(r) > 1 else ''
    return info, header, students


def parse_date(s):
    s = str(s).strip().replace('年', '-').replace('月', '-').replace('日', '')
    parts = [int(x) for x in re.split(r'[-/.]', s) if x]
    return date(parts[0], parts[1], parts[2])


def fmt_cn_date(d):
    return f'{d.year}年{d.month}月{d.day}日'


# ---------------- 模板编辑 ----------------

def set_flag(p_el, tag):
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        pPr = parse_xml(f'<w:pPr {nsdecls("w")}/>')
        p_el.insert(0, pPr)
    if pPr.find(qn(tag)) is None:
        pPr.append(parse_xml(f'<{tag} {nsdecls("w")}/>'))


def make_runs(p, values, sz_half=32):
    for r in p.findall(qn('w:r')):
        p.remove(r)
    for val in values:
        r = parse_xml(
            f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:hint="eastAsia" '
            f'w:ascii="仿宋" w:hAnsi="仿宋" w:eastAsia="仿宋"/>'
            f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
            f'<w:t xml:space="preserve">{val}</w:t></w:r>')
        p.append(r)


def set_date(ps, date_str):
    runs = ps.findall(qn('w:r'))
    if len(runs) >= 2:
        r2 = runs[1]
        ts = r2.findall(qn('w:t'))
        if ts:
            ts[0].text = date_str
            ts[0].set(qn('xml:space'), 'preserve')
        else:
            r2.append(parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve">{date_str}</w:t>'))
        for extra in runs[2:]:
            ps.remove(extra)
    elif runs:
        nr = deepcopy(runs[0])
        ts = nr.findall(qn('w:t'))
        if ts:
            ts[0].text = date_str
            ts[0].set(qn('xml:space'), 'preserve')
        ps.append(nr)


def _set_run_text(r_el, text):
    for t in r_el.findall(qn('w:t')):
        r_el.remove(t)
    r_el.append(parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve">{text}</w:t>'))


def fill_body(body_p, slot_values, slot_ranges):
    """按模板语义槽位（正则定位文本区间）替换正文，保留固定 run 原位置。
    槽位从左到右依次替换；每次替换后重算全文与区间（槽位值长度与原文本
    不同会使后续偏移失效，必须重算），并清理空 run。"""
    p_el = body_p._p
    for slot_idx, new_val in enumerate(slot_values):
        runs = p_el.findall(qn('w:r'))
        texts = [''.join(t.text or '' for t in r.findall(qn('w:t'))) for r in runs]
        offsets, pos = [], 0
        for t in texts:
            offsets.append(pos)
            pos += len(t)
        full = ''.join(texts)
        ranges = slot_ranges(full)
        if len(ranges) != len(slot_values):
            raise RuntimeError(f'槽位数量不匹配: 模板={len(ranges)} 输入={len(slot_values)}')
        s, e = ranges[slot_idx]
        idxs = [i for i in range(len(runs))
                if offsets[i] < e and offsets[i] + len(texts[i]) > s]
        if not idxs:
            raise RuntimeError(f'槽位区间 [{s},{e}) 未命中任何 run')
        i0, i1 = idxs[0], idxs[-1]
        if i0 == i1:
            prefix = texts[i0][: s - offsets[i0]]
            suffix = texts[i0][e - offsets[i0]:]
            _set_run_text(runs[i0], prefix + new_val + suffix)
        else:
            prefix = texts[i0][: s - offsets[i0]]
            suffix = texts[i1][e - offsets[i1]:]
            _set_run_text(runs[i0], prefix + new_val)
            _set_run_text(runs[i1], suffix)
            for i in range(i0 + 1, i1):
                runs[i].getparent().remove(runs[i])
            if suffix == '':
                runs[i1].getparent().remove(runs[i1])
        # 清理空 run
        for r in list(p_el.findall(qn('w:r'))):
            if not ''.join(t.text or '' for t in r.findall(qn('w:t'))):
                r.getparent().remove(r)


def build(src, dst, students, segs, slot_values, sign_date, slot_ranges, blank_lines=2):
    doc = Document(src)
    body = doc.element.body
    src_tbl = doc.tables[0]._tbl

    # 正文段 = 含「兹证明」的段落
    body_p = next(p for p in doc.paragraphs if '兹证明' in p.text)
    fill_body(body_p, slot_values, slot_ranges)

    paras = doc.paragraphs
    date_re = re.compile(r'\s*\d{4}年\d{1,2}月\d{1,2}日\s*$')
    di = [i for i, p in enumerate(paras) if date_re.match(p.text)][-1]
    date_p = paras[di]
    sig_ps = [paras[di - 2], paras[di - 1], date_p]
    blanks_ps = [p for p in paras[2:di - 2] if p.text.strip() == ''] or [paras[2]]

    src_tbl.getparent().remove(src_tbl)
    for p in blanks_ps + sig_ps:
        p._p.getparent().remove(p._p)

    anchor = paras[1]._p
    student_iter = iter(students)
    for si, seg in enumerate(segs):
        new_tbl = deepcopy(src_tbl)
        rows = new_tbl.findall(qn('w:tr'))
        header_tr = rows[0]
        data_trs = rows[1:]
        for tr in data_trs[seg:]:
            new_tbl.remove(tr)
        trPr0 = header_tr.find(qn('w:trPr'))
        if trPr0 is not None and trPr0.find(qn('w:tblHeader')) is None:
            trPr0.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        for tr in data_trs[:seg]:
            tcs = tr.findall(qn('w:tc'))
            vals = next(student_iter)
            for ci, tc in enumerate(tcs[:len(vals)]):
                make_runs(tc.find(qn('w:p')), [str(vals[ci])])
        anchor.addnext(new_tbl)
        anchor = new_tbl

        for _ in range(blank_lines):
            pb = deepcopy(blanks_ps[0]._p)
            set_flag(pb, 'w:keepNext')
            anchor.addnext(pb)
            anchor = pb
        for idx, sp in enumerate(sig_ps):
            ps = deepcopy(sp._p)
            set_flag(ps, 'w:keepLines')
            if idx < 2:
                set_flag(ps, 'w:keepNext')
            if sp is date_p:
                set_date(ps, sign_date)
                if si < len(segs) - 1:
                    # 分页符内嵌日期段末尾：不产生独立段落，杜绝空页/分页段孤行
                    ps.append(parse_xml(f'<w:r {nsdecls("w")}><w:br w:type="page"/></w:r>'))
            anchor.addnext(ps)
            anchor = ps
    doc.save(dst)


# ---------------- 渲染与核验 ----------------

def render_pdf(docx_path):
    pdf = docx_path[:-5] + '.pdf'
    ps = (
        "$ErrorActionPreference = 'Stop'\n"
        "try {\n"
        "  $word = New-Object -ComObject Word.Application\n"
        "  $word.Visible = $false\n"
        "  $word.DisplayAlerts = 0\n"
        f"  $doc = $word.Documents.Open('{docx_path}', $false, $true)\n"
        f"  $doc.ExportAsFixedFormat('{pdf}', 17)\n"
        "  $doc.Close($false)\n"
        "  $word.Quit()\n"
        "} catch {\n"
        "  Write-Output ('RENDER_ERR: ' + $_.Exception.Message)\n"
        "  try { $word.Quit() } catch {}\n"
        "  exit 1\n"
        "}\n"
    )
    r = subprocess.run(['powershell', '-NoProfile', '-Command', ps],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError('Word 渲染失败: ' + r.stdout + r.stderr)
    return pdf


def verify_pdf(pdf):
    doc = pymupdf.open(pdf)
    problems, infos = [], []
    for i, page in enumerate(doc):
        blocks = [b for b in page.get_text('blocks') if b[6] == 0 and b[4].strip()]
        blocks.sort(key=lambda b: b[3])
        text = ''.join(b[4] for b in blocks).replace(' ', '')
        has_sig = ('共青团' in text and '服务站' in text
                   and re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text) is not None)
        hdr_blocks = [b for b in page.get_text('blocks') if ('学号' in b[4] or '寝室号' in b[4])]
        has_table = len(hdr_blocks) > 0
        hdr_bottom = max(b[3] for b in hdr_blocks) if hdr_blocks else 0
        rows_cnt = len([b for b in page.get_text('blocks')
                        if b[0] < 150 and b[1] > hdr_bottom and b[4].strip()])
        imgs = len(page.get_images(full=True))
        last = blocks[-1][4].strip().replace(' ', '') if blocks else ''
        if not blocks:
            problems.append((i, 'EMPTY', rows_cnt))
        if not has_sig:
            problems.append((i, 'NO_SIG', rows_cnt))
        if has_sig and not has_table:
            problems.append((i, 'ORPHAN', rows_cnt))
        if imgs < 1:
            problems.append((i, 'NO_HEADER', rows_cnt))
        if not (last and '年' in last and '日' in last):
            problems.append((i, 'NO_DATE_END', rows_cnt))
        infos.append(dict(page=i + 1, has_sig=has_sig, has_table=has_table,
                          rows=rows_cnt, last_y=round(blocks[-1][3], 1) if blocks else 0))
    doc.close()
    return problems, infos


def compute_segs(n, F1, F2):
    if n <= 0:
        raise ValueError('名单为空')
    if n <= F1:
        return [n]
    segs = [F1]
    rest = n - F1
    npg = math.ceil(rest / F2)
    base, rem = divmod(rest, npg)
    segs += [base + (1 if i < rem else 0) for i in range(npg)]
    return segs


def solve(src, dst, students, slot_values, sign_date, slot_ranges, init_caps=(7, 11)):
    F1, F2 = init_caps
    last_pdf = None
    for rnd in range(1, MAX_ROUNDS + 1):
        segs = compute_segs(len(students), F1, F2)
        build(src, dst, students, segs, slot_values, sign_date, slot_ranges, BLANK_LINES)
        pdf = render_pdf(dst)
        last_pdf = pdf
        problems, infos = verify_pdf(pdf)
        log = [f'[R{rnd}] caps=({F1},{F2}) segs={segs} pages={len(infos)}']
        for inf in infos:
            log.append(f'  P{inf["page"]}: rows={inf["rows"]:2d} last_y={inf["last_y"]:6.1f}')
        if not problems:
            log.append('  核验全部通过 ✓')
            print('\n'.join(log))
            return segs
        for i, kind, rows in problems:
            if kind in ('NO_SIG', 'NO_DATE_END'):
                cap = 'F1' if i == 0 else 'F2'
                new = max(rows - 1, 1)
                if cap == 'F1':
                    F1 = min(F1, new)
                else:
                    F2 = min(F2, new)
                log.append(f'  ADJ {cap} {rows}->{new} (P{i + 1}: {kind})')
            elif kind == 'ORPHAN':
                cap = 'F1' if (i - 1) == 0 else 'F2'
                new = max(infos[i - 1]['rows'] - 1, 1)
                if cap == 'F1':
                    F1 = min(F1, new)
                else:
                    F2 = min(F2, new)
                log.append(f'  ADJ {cap}->{new} (P{i + 1}: 孤行落款)')
            elif kind == 'EMPTY':
                cap = 'F1' if (i - 1) == 0 else 'F2'
                new = max(infos[i - 1]['rows'] - 1, 1)
                if cap == 'F1':
                    F1 = min(F1, new)
                else:
                    F2 = min(F2, new)
                log.append(f'  ADJ {cap}->{new} (P{i + 1}: 空白页)')
            elif kind == 'NO_HEADER':
                log.append(f'  WARN P{i + 1}: 页首题头缺失')
        print('\n'.join(log))
    raise RuntimeError('多次调整仍无法通过分页核验，请检查模板或数据')


# ---------------- 主流程 ----------------

def main():
    ap = argparse.ArgumentParser(description='请假证明一键生成（乙方案）')
    ap.add_argument('input', help='输入文件 xlsx/csv/json')
    ap.add_argument('--out', default=WORK, help='输出目录（默认脚本目录）')
    ap.add_argument('--name', default=None, help='输出文件名（默认自动）')
    args = ap.parse_args()

    info, header, students = read_input(args.input)

    template_id = info.get('template') or info.get('模板')
    if not template_id:
        if '学院' in header:
            template_id = '上课请假'
        elif '寝室号' in header:
            template_id = '分寝室请假'
    if template_id not in TEMPLATES:
        raise RuntimeError(f'未知模板: {template_id}（可选: {list(TEMPLATES)}）')
    tpl = TEMPLATES[template_id]

    # 校验列头
    missing = [h for h in tpl['headers'] if h not in header]
    if missing:
        raise RuntimeError(f'名单列头缺少: {missing}（应为 {tpl["headers"]}）')
    col_idx = {h: header.index(h) for h in tpl['headers']}
    students = [[row[col_idx[h]] if col_idx[h] < len(row) else '' for h in tpl['headers']]
                for row in students]

    missing_params = [f for f in tpl['required'] if not info.get(f)]
    if missing_params:
        raise RuntimeError(f'活动信息缺少字段: {missing_params}')

    ev = parse_date(info['event_date'])
    sign = parse_date(info.get('sign_date') or info.get('落款日期') or (ev - timedelta(days=1)))
    wd = WEEKDAY_CN[ev.weekday()]
    p = dict(y=ev.year, m=ev.month, d=ev.day, wd=wd)
    if template_id == '上课请假':
        slot_values = [f'{p["y"]}年{p["m"]}月{p["d"]}日({p["wd"]})',
                       info['periods'].rstrip() + ' ',
                       f"到{info['location']}出席{info['activity']}，志愿者活动，"]
    else:
        slot_values = [info['dorm'],
                       f'{p["y"]}年{p["m"]}月{p["d"]}日（{p["wd"]}）',
                       f"参加{info['activity']}志愿者工作，{info['special_note']}，"]
    sign_date = fmt_cn_date(sign)

    src = None
    for cand in (os.path.join(WORK, '模板', tpl['file']), os.path.join(WORK, tpl['file'])):
        if os.path.exists(cand):
            src = cand
            break
    if src is None:
        raise RuntimeError(f'模板文件不存在: {tpl["file"]}（需与 generate.py 同目录或 模板/ 子目录）')
    name = args.name or f'{template_id.replace("请假", "")}请假证明-{ev:%Y%m%d}.docx'
    dst = os.path.join(args.out, name)

    solve(src, dst, students, slot_values, sign_date, tpl["slot_ranges"])
    print(f'\n输出: {dst}')
    print(f'落款日期: {sign_date}  |  学生数: {len(students)}')
    return dst


if __name__ == '__main__':
    main()
