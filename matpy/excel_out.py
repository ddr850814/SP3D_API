# -*- coding: utf-8 -*-
"""Excel 输出：管道材料表（对应 CLB Form1 单快照导出的 NPOI 写表逻辑，
字段对齐 HQSHSP3DReport frmReportPipePart 管道部件报告）。

Sheet「管道材料」，双模式：
  by_line=False（按商品代码汇总，跨管线合并数量）：目录级字段
  by_line=True （按管线分组）：全字段（含 run 级属性：管道等级/名称、
                操作与设计压力温度、保温三件套、单元/区域名称）

目录属性缺失时回退：商品代码 ContractorCommodityCode → CC 码
（IndustryCommodityCode）→ PartNo 前段；长描述回退 PartNo 后段
（对齐 CLB GetPart 失败时的行为）。列宽自适应（AutoSizeColumn 等价：
先扫描内容长度再设宽）。
"""
from typing import Dict, List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter

from mat_core import MatRow
from weld_core import WeldInfo

# 类型码（mat_core 类型表 JOIN 判定：直管/管件/仪表件/特殊件）
_CLS_PIPE = "pipe"
_CLS_SPECIALTY = "specialty"
_CLS_INSTRUMENT = "instrument"


def _attr_key(r: MatRow) -> str:
    return r.partno.split("|")[0]


def write_excel(rows: List[MatRow], attrs: Dict[str, dict], by_line: bool,
                out_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "管道材料"

    def _unit(r: MatRow) -> str:
        return "米" if r.cls == _CLS_PIPE else "只"

    def _type_text(r: MatRow, a: dict) -> str:
        if r.cls == _CLS_INSTRUMENT:
            return "仪表件"
        if r.cls == _CLS_SPECIALTY:
            return "特殊件"
        return a.get("commodity_type", "")

    def _unit_weight(weight, qty) -> float | None:
        # 直管 weight=切段总重合计 → /米数 = kg/m；管件 = 平均单件重
        if weight is None or not qty or qty <= 0:
            return None
        return weight / qty

    if by_line:
        headers = ["管线号", "部件类型", "商品代码", "CC码", "通径1", "通径2",
                   "壁厚1", "壁厚2", "管道等级", "管道名称", "操作最大压力(MPa)",
                   "操作最大温度(℃)", "设计最大压力(MPa)", "设计最大温度(℃)",
                   "保温目的",
                   "保温材料", "保温厚度(mm)", "材料长描述", "单位重量", "数量",
                   "单位", "单元名称", "区域名称"]
        ordered = sorted(
            rows, key=lambda r: (r.line,
                                 (attrs.get(_attr_key(r)) or {}).get("class_sort", 7),
                                 _attr_key(r), r.partno))
        width = [len(h) for h in headers]
        for j, r in enumerate(ordered):
            row = j + 2
            pn_seg = r.partno.split("|")
            a = attrs.get(pn_seg[0]) or {}
            # 商品代码：Contractor → CC 码保底 → PartNo 前段
            code = (a.get("commodity_code") if a.get("commodity_code") != pn_seg[0]
                    else None) or a.get("industry_code") or pn_seg[0]
            cc = a.get("industry_code") or pn_seg[0]
            desc = a.get("long_desc") or (pn_seg[1] if len(pn_seg) > 1 else "")
            uw = _unit_weight(r.weight, r.qty)
            thk = round(r.ins_thk, 1) if r.ins_thk is not None else None
            vals = [r.line, _type_text(r, a), code, cc,
                    a.get("primary", 0.0), a.get("secondary", 0.0),
                    a.get("sch1", ""), a.get("sch2", ""),
                    r.spec, r.run_name, r.op_p, r.op_t, r.de_p, r.de_t,
                    r.ins_purpose, r.ins_material, thk, desc,
                    uw, r.qty, _unit(r), r.unit_name, r.area_name]
            for i, v in enumerate(vals, 1):
                ws.cell(row=row, column=i, value=v).font = Font(
                    name="Times New Roman", size=10)
                w = min(len(str(v)) + 2, 60)
                if w > width[i - 1]:
                    width[i - 1] = w
    else:
        # 按商品代码汇总：跨管线合并数量与重量（特殊件带 |位号 后缀，天然单列）；
        # run 级属性跨管线不同，不生成对应列
        headers = ["商品代码", "CC码", "部件类型", "数量", "材料长描述",
                   "通径1", "通径2", "壁厚1", "壁厚2", "单位重量", "单位"]
        merged: Dict[str, list] = {}
        for r in rows:
            m = merged.setdefault(r.partno, [0.0, None, r.cls])
            m[0] += r.qty
            if r.weight is not None:
                m[1] = (m[1] or 0.0) + r.weight
        ordered = sorted(
            merged.items(),
            key=lambda kv: ((attrs.get(kv[0].split("|")[0]) or {}).get("class_sort", 7),
                            kv[0].split("|")[0], kv[0]))
        width = [len(h) for h in headers]
        for j, (pn, (qty, wt, cls)) in enumerate(ordered):
            row = j + 2
            pn_seg = pn.split("|")
            a = attrs.get(pn_seg[0]) or {}
            code = (a.get("commodity_code") if a.get("commodity_code") != pn_seg[0]
                    else None) or a.get("industry_code") or pn_seg[0]
            cc = a.get("industry_code") or pn_seg[0]
            desc = a.get("long_desc") or (pn_seg[1] if len(pn_seg) > 1 else "")
            fake = MatRow(line="", partno=pn, qty=qty, cls=cls)
            vals = [code, cc, _type_text(fake, a), qty, desc,
                    a.get("primary", 0.0), a.get("secondary", 0.0),
                    a.get("sch1", ""), a.get("sch2", ""),
                    _unit_weight(wt, qty), _unit(fake)]
            for i, v in enumerate(vals, 1):
                ws.cell(row=row, column=i, value=v).font = Font(
                    name="Times New Roman", size=10)
                w = min(len(str(v)) + 2, 60)
                if w > width[i - 1]:
                    width[i - 1] = w

    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(name="宋体", size=10, bold=True)
    for i, w in enumerate(width, 1):
        ws.column_dimensions[get_column_letter(i)].width = w + 2

    wb.save(out_path)


# ------------------------------------------------------- 螺栓/垫片报表 ----


def _write_detail_sheet(ws, headers: list, rows: List[list]) -> None:
    """明细表通用写法：第 1 行表头、数据从第 2 行，列宽自适应。"""
    width = [len(h) for h in headers]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h).font = Font(
            name="宋体", size=10, bold=True)
    for j, vals in enumerate(rows):
        row = j + 2
        for i, v in enumerate(vals, 1):
            ws.cell(row=row, column=i, value=v).font = Font(
                name="Times New Roman", size=10)
            w = min(len(str(v)) + 2, 60)
            if w > width[i - 1]:
                width[i - 1] = w
    for i, w in enumerate(width, 1):
        ws.column_dimensions[get_column_letter(i)].width = w + 2


def write_bolts(rows, out_path: str, by_line: bool = True) -> int:
    """螺栓报告：by_line=True 逐套明细（对齐 VB frmReportBolt 列）；
    by_line=False 按商品代码汇总——(直径, 长度, CC码) 合计套内数量，
    描述/业主码取组内首个非空（采购汇总口径，VB 无此形态）。

    返回写入的数据行数。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "螺栓"
    if by_line:
        headers = ["OID", "名称", "螺栓直径", "长度(mm)", "数量", "CC码",
                   "材料长描述", "业主编码", "个数", "单位", "部件名称",
                   "管道名称", "管线名称", "区域名称", "单元名称"]
        data = [[r.oid, "螺栓", r.diameter, r.length_mm, r.qty, r.cc_code,
                 r.long_desc, r.owner_code, 1, "套", r.part_pns,
                 r.run_names, r.line_names, r.area_name, r.unit_name]
                for r in rows]
    else:
        merged: Dict[tuple, list] = {}
        for r in rows:
            m = merged.setdefault((r.diameter, r.length_mm, r.cc_code),
                                  [0, "", ""])
            m[0] += r.qty
            m[1] = m[1] or r.long_desc
            m[2] = m[2] or r.owner_code
        headers = ["螺栓直径", "长度(mm)", "CC码", "材料长描述",
                   "业主编码", "数量", "单位"]
        data = [[dia, length, cc, desc, owner, qty, "套"]
                for (dia, length, cc), (qty, desc, owner)
                in sorted(merged.items())]
    _write_detail_sheet(ws, headers, data)
    wb.save(out_path)
    return len(data)


def write_gaskets(rows, out_path: str, by_line: bool = True) -> int:
    """垫片报告：by_line=True 逐片明细（对齐 VB frmReportGasket 列）；
    by_line=False 按商品代码汇总——(通径, CC码) 合计片数，
    描述/业主采购码取组内首个非空（采购汇总口径，VB 无此形态）。

    返回写入的数据行数。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "垫片"
    if by_line:
        headers = ["OID", "名称", "通径", "CC码", "材料长描述", "业主采购码",
                   "数量", "单位", "部件名称", "管道名称", "管线名称",
                   "区域名称", "单元名称"]
        data = [[r.oid, "垫片", r.nd, r.cc_code, r.long_desc, r.owner_code,
                 1, "个", r.part_pns, r.run_names, r.line_names,
                 r.area_name, r.unit_name] for r in rows]
    else:
        merged: Dict[tuple, list] = {}
        for r in rows:
            m = merged.setdefault((r.nd, r.cc_code), [0, "", ""])
            m[0] += 1
            m[1] = m[1] or r.long_desc
            m[2] = m[2] or r.owner_code
        headers = ["通径", "CC码", "材料长描述", "业主采购码", "数量", "单位"]
        data = [[nd, cc, desc, owner, qty, "个"]
                for (nd, cc), (qty, desc, owner)
                in sorted(merged.items())]
    _write_detail_sheet(ws, headers, data)
    wb.save(out_path)
    return len(data)


def write_instruments(rows, out_path: str) -> None:
    """仪表报告（对齐 VB frmReportInstrument 列，逐台一行）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "仪表"
    headers = ["OID", "位号", "管道名称", "管线名称", "单元名称", "区域名称",
               "部件码", "坐标X", "坐标Y", "坐标Z"]
    data = [[r.oid, r.tag, r.run_name, r.line_name, r.unit_name,
             r.area_name, r.part_code,
             r.x if r.x is not None else "",
             r.y if r.y is not None else "",
             r.z if r.z is not None else ""] for r in rows]
    _write_detail_sheet(ws, headers, data)
    wb.save(out_path)


def write_piping_runs(rows, out_path: str) -> None:
    """管道报告（对齐 VB frmReportPipingRun 列，逐 run 一行）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "管道属性"
    headers = ["OID", "管道名称", "管线名称", "管道等级",
               "Operating Max Temp(℃)", "Operating Max Pressure(MPa)",
               "Design Max Temp(℃)", "Design Max Pressure(MPa)",
               "IsInsulated", "保温目的", "保温材料", "保温温度(℃)",
               "保温厚度(mm)", "单元名称", "区域名称"]

    def _b(v):
        return "" if v is None else ("True" if v else "False")

    data = [[r.oid, r.name, r.line_name, r.spec,
             r.op_t if r.op_t is not None else "",
             r.op_p if r.op_p is not None else "",
             r.de_t if r.de_t is not None else "",
             r.de_p if r.de_p is not None else "",
             _b(r.is_insulated), r.ins_purpose, r.ins_material,
             r.ins_temp if r.ins_temp is not None else "",
             r.ins_thk if r.ins_thk is not None else "",
             r.unit_name, r.area_name] for r in rows]
    _write_detail_sheet(ws, headers, data)
    wb.save(out_path)


def write_pid_numbers(rows, out_path: str) -> None:
    """PID/应力号报告（对齐 VB frmReportPIDStressNo 列，逐条一行）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "PID应力号"
    headers = ["OID", "模型名称", "PID号", "应力分析题号", "模型类型"]
    data = [[r.oid, r.name, r.pid_no, r.stress_no, r.model_type]
            for r in rows]
    _write_detail_sheet(ws, headers, data)
    wb.save(out_path)


def write_pid_correlations(rows, out_path: str) -> None:
    """PID 关联报告（对齐 VB frmReportCorrelate 列，逐对象一行）。

    属性比较列依赖 COM API 不可复刻：未关联行填「未关联PID」。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "PID关联"
    headers = ["OID", "模型名称", "模型类", "模型路径", "关联PID",
               "关联状态", "关联批准", "属性比较"]
    data = [[r.oid, r.name, r.model_class, r.path, r.basis,
             r.status, r.approval, r.comparison] for r in rows]
    _write_detail_sheet(ws, headers, data)
    wb.save(out_path)


def write_iso_drawings(rows, out_path: str) -> None:
    """ISO 图纸报告（对齐 VB frmReportISODrawing 列，每图一行）。

    页数取自图纸内容 XML（Deflate64 zip 解压后的 HEADERTotalSheet）；
    无图纸内容的留空。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "ISO图纸"
    headers = ["OID", "名称", "路径", "终版版次", "终版描述",
               "终版设计人", "终版校核人", "终版审核人", "终版日期", "页数"]
    data = [[r.oid, r.name, r.path, r.rev_minor, r.rev_desc,
             r.rev_by, r.chk_by, r.app_by, r.rev_date, r.pages]
            for r in rows]
    _write_detail_sheet(ws, headers, data)
    wb.save(out_path)


def write_weld_excel(welds: Dict[WeldInfo, int], met: int, thk: bool,
                     out_path: str) -> None:
    """焊接量报告（weldpy 的 NPOI 写表逻辑，列与原 template.xlsx 逐字一致）。

    列按选项动态裁剪：met 0=不分组（无材料列）/1=按材料等级/2=按材料组；
    thk=忽略壁厚时不生成管表号/壁厚两列。备注列始终保留作手写栏。
    """
    wb = Workbook()
    ws = wb.active

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_border = Border(left=thin, right=thin)   # 原模板表头仅左右框线
    font_data = Font(name="Times New Roman", size=10)
    font_title = Font(name="宋体", size=10)
    fill_green = PatternFill(start_color="FF00FF00", end_color="FF00FF00",
                             fill_type="solid")

    s1 = NamedStyle(name="ws1", font=font_data, border=border)
    s2 = NamedStyle(name="ws2", font=font_data, border=border, fill=fill_green)
    s3 = NamedStyle(name="ws3", font=font_title, border=border, fill=fill_green,
                    alignment=Alignment(horizontal="center"))
    wb.add_named_style(s1)
    wb.add_named_style(s2)
    wb.add_named_style(s3)

    headers = ["管线号", "焊接点口径/寸"]
    if met == 1:
        headers.append("材料等级")
    elif met == 2:
        headers.append("材料组")
    if not thk:
        headers += ["管表号", "壁厚/mm"]
    headers += ["焊口数量", "壁厚系数", "焊缝量1", "备注\nNOTE"]
    col = {h: i for i, h in enumerate(headers, 1)}
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(size=10)
        c.border = hdr_border

    # 按管线号排序（对应 WeldL = Welds.OrderBy(PipelineId)）
    wl = sorted(welds.items(), key=lambda kv: kv[0].pipeline_id)
    n = len(wl)

    for i, (w, cnt) in enumerate(wl):
        row = i + 2  # Excel 行号（第 1 行是表头）
        ws.cell(row=row, column=col["管线号"], value=w.pipeline_id).style = s1
        ws.cell(row=row, column=col["焊接点口径/寸"], value=w.npd).style = s1
        if met == 1:
            ws.cell(row=row, column=col["材料等级"], value=w.spec).style = s1
        elif met == 2:
            ws.cell(row=row, column=col["材料组"], value=w.materials_group).style = s1
        c_cnt = col["焊口数量"]
        if not thk:
            ws.cell(row=row, column=col["管表号"],
                    value=w.schedule_thickness).style = s1
            ws.cell(row=row, column=col["壁厚/mm"],
                    value=w.wall_thickness).style = s1
            ws.cell(row=row, column=c_cnt, value=cnt).style = s1
            ws.cell(row=row, column=col["壁厚系数"],
                    value=w.wall_thickness / 10).style = s1
            ws.cell(row=row, column=col["焊缝量1"],
                    value=w.wall_thickness / 10 * w.npd * cnt).style = s1
        else:
            ws.cell(row=row, column=c_cnt, value=cnt).style = s1
            ws.cell(row=row, column=col["壁厚系数"], value=1).style = s1
            ws.cell(row=row, column=col["焊缝量1"],
                    value=w.npd * cnt).style = s1
        ws.cell(row=row, column=col["备注\nNOTE"]).style = s1  # 备注列保留边框

    # 总计
    if thk:
        total = sum(w.npd * cnt for w, cnt in wl)
    else:
        total = sum(w.wall_thickness / 10 * w.npd * cnt for w, cnt in wl)

    # 分组汇总（GroupSum）：met=1 按材料等级(SPEC)、met=2 按材料组，
    # 写在壁厚系数（名）/焊缝量1（值）两列下
    if met == 0:
        group_count = 0
    else:
        groups = []
        seen = []
        for w, cnt in wl:
            key = w.spec if met == 1 else w.materials_group
            if key not in seen:
                seen.append(key)
                groups.append([key, 0.0])
            idx = seen.index(key)
            groups[idx][1] += ((w.npd * cnt) if thk
                               else (w.wall_thickness / 10 * w.npd * cnt))
        for j, (gname, gtotal) in enumerate(groups):
            row = n + 4 + j  # 数据 n 行 + 表头 + SkipRowCount
            ws.cell(row=row, column=col["壁厚系数"], value=gname).style = s3
            ws.cell(row=row, column=col["焊缝量1"], value=gtotal).style = s2
        group_count = len(groups)

    # 总计行
    row_total = n + 4 + group_count
    ws.cell(row=row_total, column=col["壁厚系数"], value="总计").style = s3
    ws.cell(row=row_total, column=col["焊缝量1"], value=total).style = s2

    wb.save(out_path)


# ------------------------------------------------------- 快照/比较导出 ----


def write_snapshot_excel(report: str, headers: list, rows: List[list],
                         out_path: str) -> int:
    """单快照导出（快照管理对话框选 1 个时的全量输出）。

    快照文件自描述（headers+rows 原样落盘），不依赖数据库连接。"""
    wb = Workbook()
    ws = wb.active
    ws.title = report
    _write_detail_sheet(ws, headers, rows)
    wb.save(out_path)
    return len(rows)


def write_compare_excel(report: str, name_a: str, name_b: str, headers: list,
                        added: List[list], removed: List[list],
                        out_path: str) -> tuple:
    """比较导出：两 sheet「增加/减少」（语义同 CLB 增加材料/减少材料），
    表下附 A1 说明行（旧→新快照名）。返回 (增加行数, 减少行数)。"""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "增加"
    _write_detail_sheet(ws1, headers, added)
    ws2 = wb.create_sheet("减少")
    _write_detail_sheet(ws2, headers, removed)
    note = f"{report}比较：{name_a} → {name_b}（生成于工具）"
    for ws in (ws1, ws2):
        c = ws.cell(row=1, column=len(headers) + 2, value=note)
        c.font = Font(name="宋体", size=9, italic=True)
    wb.save(out_path)
    return len(added), len(removed)
