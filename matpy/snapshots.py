# -*- coding: utf-8 -*-
"""统计快照与同类型比较（CLB.vb「材料表比较」的 SQL 直连复刻）。

快照 = 一次统计结果落盘为 JSON（%APPDATA%/MatMat/snapshots/），文件自描述
（表头/数据行/比较键/数量列都存档），比较与导出不依赖数据库连接——这是
对 CLB 版（快照只存 3 列、导出时必须连目录库）的改进。

比较语义对齐 CLB CreateSQLiteViewAd/Re：
  增加 = 新快照独有 + 共同键数量增加（取新快照行）
  减少 = 旧快照独有 + 共同键数量减少（取旧快照行）
行尾追加「旧数量/新数量/差量」三列辅助核对。仅同 report 类型可比较；
管道材料/螺栓/垫片/管夹支持「按管线号分组」两档键，焊接量键即全规格维度。
"""
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from mat_core import MatRow
from weld_core import WeldInfo

# 类型码（与 excel_out 一致：直管/特殊件/仪表件）
_CLS_PIPE = "pipe"
_CLS_SPECIALTY = "specialty"
_CLS_INSTRUMENT = "instrument"


def snap_dir() -> Path:
    return Path(os.environ["APPDATA"]) / "MatMat" / "snapshots"


# ------------------------------------------------------- 各报表快照列 ----
# 快照统一存明细粒度；汇总档比较时按 sum_key_idx 重新聚合。
# 螺栓/垫片逐套（个）一行，数量列恒为 1——比较数量是套/片数，
# 不是套内条数（与明细导出的「个数=1」口径一致）。


def _mat_rows(rows: List[MatRow], attrs: Optional[dict]) -> List[list]:
    attrs = attrs or {}
    out = []
    for r in rows:
        seg = r.partno.split("|")
        a = attrs.get(seg[0]) or {}
        code = (a.get("commodity_code")
                if a.get("commodity_code") != seg[0] else None) \
            or a.get("industry_code") or seg[0]
        cc = a.get("industry_code") or seg[0]
        desc = a.get("long_desc") or (seg[1] if len(seg) > 1 else "")
        if r.cls == _CLS_INSTRUMENT:
            typ = "仪表件"
        elif r.cls == _CLS_SPECIALTY:
            typ = "特殊件"
        else:
            typ = a.get("commodity_type", "")
        out.append([r.line, r.partno, typ, code, cc,
                    a.get("primary", 0.0), a.get("secondary", 0.0),
                    a.get("sch1", ""), a.get("sch2", ""), r.spec, desc,
                    r.qty, "米" if r.cls == _CLS_PIPE else "只",
                    r.unit_name, r.area_name])
    return out


def _weld_rows(welds: Dict[WeldInfo, int], attrs) -> List[list]:
    merged: Dict[tuple, int] = {}
    for w, cnt in welds.items():
        k = (w.pipeline_id, w.npd, w.spec, w.materials_group,
             w.schedule_thickness, w.wall_thickness)
        merged[k] = merged.get(k, 0) + cnt
    return [list(k) + [n] for k, n in sorted(merged.items())]


def _bolt_rows(rows, attrs) -> List[list]:
    return [[r.line_names, r.diameter, r.length_mm, r.cc_code, r.long_desc,
             r.owner_code, 1, r.run_names, r.area_name, r.unit_name]
            for r in rows]


def _gasket_rows(rows, attrs) -> List[list]:
    return [[r.line_names, r.nd, r.cc_code, r.long_desc, r.owner_code, 1,
             r.run_names, r.area_name, r.unit_name]
            for r in rows]


def _clamp_rows(rows, attrs) -> List[list]:
    return [[r.line_names, r.npd, r.spec, r.cc_code, r.clamp_option, 1,
             r.run_names, r.area_name, r.unit_name]
            for r in rows]


SPECS: Dict[str, dict] = {
    "管道材料": dict(
        headers=["管线号", "部件码", "部件类型", "商品代码", "CC码", "通径1",
                 "通径2", "壁厚1", "壁厚2", "管道等级", "材料长描述",
                 "数量", "单位", "单元名称", "区域名称"],
        key_idx=(0, 1),                      # 管线号+部件码（同 CLB Line+PartNo）
        sum_key_idx=(2, 3, 4, 5, 6, 7, 8),   # 部件类型+代码+规格
        qty_idx=11,
        build=_mat_rows,
    ),
    "焊接量": dict(
        headers=["管线号", "焊接点口径/寸", "材料等级", "材料组", "管表号",
                 "壁厚/mm", "焊口数量"],
        key_idx=(0, 1, 2, 3, 4, 5),
        sum_key_idx=None,                    # 键即全规格维度，无汇总档
        qty_idx=6,
        build=_weld_rows,
    ),
    "螺栓": dict(
        headers=["管线名称", "螺栓直径", "长度(mm)", "CC码", "材料长描述",
                 "业主编码", "数量(套)", "管道名称", "区域名称", "单元名称"],
        key_idx=(0, 1, 2, 3),
        sum_key_idx=(1, 2, 3),
        qty_idx=6,
        build=_bolt_rows,
    ),
    "垫片": dict(
        headers=["管线名称", "通径", "CC码", "材料长描述", "业主采购码",
                 "数量(个)", "管道名称", "区域名称", "单元名称"],
        key_idx=(0, 1, 2, 3),
        sum_key_idx=(1, 2),
        qty_idx=5,
        build=_gasket_rows,
    ),
    "管夹": dict(
        headers=["管线名称", "通径", "管道等级", "商品代码", "卡箍型号",
                 "数量(个)", "管道名称", "区域名称", "单元名称"],
        key_idx=(0, 1, 2, 3),
        sum_key_idx=(1, 2, 3),
        qty_idx=5,
        build=_clamp_rows,
    ),
}


# ------------------------------------------------------------ 存取 ----


def save(report: str, plant_code: str, scope: str, rows, attrs=None,
         by_line=None, weld_opts=None) -> dict:
    """把一次统计结果写成快照文件，返回含 path 的元数据 dict。

    by_line：产出快照时「按管线号分组」勾选状态（True/False），固化进
    快照——比较时自动按此口径聚合，不依赖比较时的界面选择；
    焊接量无此开关恒 None。
    weld_opts：焊接量统计口径 {"met": 0/1/2, "thk": bool, "skip_pipes": bool}
    （分组模式/忽略壁厚/忽略直管焊缝）——统计时 WeldInfo 已按此降维，
    固化标记供比较前拦截不同口径互比；其它报表恒 None。
    """
    spec = SPECS[report]
    now = datetime.now()
    meta = {
        "v": 1,
        "report": report,
        "plant": plant_code,
        "time": now.strftime("%Y-%m-%d %H:%M:%S"),
        "scope": scope,
        "by_line": by_line,
        "weld_opts": weld_opts,
        "headers": spec["headers"],
        "key_idx": list(spec["key_idx"]),
        "sum_key_idx": (list(spec["sum_key_idx"])
                        if spec["sum_key_idx"] else None),
        "qty_idx": spec["qty_idx"],
        "rows": spec["build"](rows, attrs),
    }
    d = snap_dir()
    d.mkdir(parents=True, exist_ok=True)
    stamp = now.strftime("%Y%m%d_%H%M%S")
    path = d / "{}__{}__{}.json".format(report, plant_code, stamp)
    # 同秒再存（如缓存命中路径快速双击）不覆盖，追加序号
    n = 2
    while path.exists():
        path = d / "{}__{}__{}_{}.json".format(report, plant_code, stamp, n)
        n += 1
    tmp = path.with_suffix(".tmp")           # 原子写（与 save_config 同约定）
    tmp.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)
    meta["path"] = str(path)
    return meta


def list_snapshots() -> List[dict]:
    """扫描快照目录，按时间升序返回 [{path, report, plant, time, scope}]。
    损坏文件跳过。"""
    out = []
    d = snap_dir()
    if not d.is_dir():
        return out
    for p in sorted(d.glob("*.json")):
        try:
            m = json.loads(p.read_text(encoding="utf-8"))
            out.append({"path": str(p), "report": m.get("report", "?"),
                        "plant": m.get("plant", "?"),
                        "time": m.get("time", ""),
                        "scope": m.get("scope", ""),
                        "by_line": m.get("by_line"),
                        "weld_opts": m.get("weld_opts")})
        except Exception:
            continue
    # time 相同（同秒多存）时按文件名序——追加序号后即保存顺序
    out.sort(key=lambda m: (m["time"], m["path"]))
    return out


def load(path: str) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def delete(path: str) -> None:
    Path(path).unlink(missing_ok=True)


# ------------------------------------------------------------ 比较 ----


def _agg(meta: dict, key_idx, qty_idx: int) -> Dict[tuple, list]:
    """按键聚合：key -> [数量合计, 行（空缺列补组内首个非空）]。"""
    d: Dict[tuple, list] = {}
    for row in meta["rows"]:
        k = tuple(row[i] for i in key_idx)
        rec = d.setdefault(k, [0.0, None])
        q = row[qty_idx] or 0
        rec[0] += q
        if rec[1] is None:
            rec[1] = list(row)
        else:
            base = rec[1]
            for i, v in enumerate(row):
                if base[i] in (None, "") and v not in (None, ""):
                    base[i] = v
    return d


def compare(a: dict, b: dict):
    """比较同类型、同记录方式的两份快照（a 旧 b 新，调用方按 time 保证先后）。

    聚合口径取快照自身的 by_line（产出时固化的「按管线号分组」勾选状态），
    两份口径不同直接抛错——按管线号记录与不按管线号记录的快照不可比。
    返回 (headers, added_rows, removed_rows)；数据行取所属快照的原始行，
    行尾追加 [旧数量, 新数量, 差量]。数量 round 6 位防浮点累加误差。
    """
    if a.get("by_line") != b.get("by_line"):
        raise ValueError("两份快照记录方式不同（一份按管线号记录、一份不按），"
                         "不能比较。请选同记录方式的两份快照。")
    if a.get("weld_opts") != b.get("weld_opts"):
        raise ValueError("两份焊接量快照的统计设置不同"
                         "（分组模式/忽略壁厚/忽略直管焊缝不一致），"
                         "不能比较。请选同设置的两份快照。")
    qty_idx = a["qty_idx"]
    by_line = a.get("by_line")
    key_idx = a["key_idx"] if (by_line or not a.get("sum_key_idx")) \
        else a["sum_key_idx"]
    aa = _agg(a, key_idx, qty_idx)
    bb = _agg(b, key_idx, qty_idx)
    headers = list(a["headers"]) + ["旧数量", "新数量", "差量"]
    added: List[list] = []
    removed: List[list] = []
    for k in sorted(set(aa) | set(bb)):
        q1 = round(aa[k][0], 6) if k in aa else 0.0
        q2 = round(bb[k][0], 6) if k in bb else 0.0
        if q2 > q1:
            added.append(list(bb[k][1]) + [q1, q2, round(q2 - q1, 6)])
        elif q1 > q2:
            removed.append(list(aa[k][1]) + [q1, q2, round(q2 - q1, 6)])
    return headers, added, removed
