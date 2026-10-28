# -*- coding: utf-8 -*-
"""核心材料采集逻辑：SP3D 管道材料清单 SQL 直连（CLB.vb 的 API 路径 SQL 化）。

四类对象按 (管线, PartNo) 聚合：直管数量=CutLength（米），其余=件数；
特殊件/仪表 PartNo 带 JNamedItem 对象名后缀（PNo|Name 键）。类型判定一律
本体表/类型表 JOIN，不用 oid 前缀；PartNumber 两路（REFDATPipeComponent
直取 / 命名 def 载体名）。判定表映射、各报表链路口径与实测依据统一见
项目 skill matpy-summary。

双后端：SQL Server / Oracle，_t/_gl/nvl 按 provider 分流（同 weldpy）。
"""
from dataclasses import dataclass, field
import re
from typing import Callable, Dict, List, Optional, Set, Tuple

from db import Db

# SP3D 全局关系类型前缀（类 GUID）
_REL_PART2DEF = "B0E39EC4"   # 部件实例 → 目录定义（madeFrom part）
_REL_NAME_CARRIER = "5280312B"  # 命名关系：oidTarget 的对象名在 RelationName
_REL_RUNUSESSPEC = "A53D19B6"  # run → 所用规格（PathRunUsesSpec，Oracle 直查用）
_REL_SHEET2DOC = "52683F7A"

# Oracle RAW 等值比较用完整 GUID（可走索引；RAWTOHEX 前缀 LIKE 逼全扫，
# weld_core 同款常量；probe 实测每前缀 DISTINCT=1，跨站点全局一致）
_GUID_RUNUSESSPEC = "A53D19B621DE4563B7BC03DE5D632E29"
_GUID_NAME_CARRIER = "5280312BE69C11D1A966080036069A02"


def safe_sql(value: str) -> str:
    """转义单引号，避免拼接 SQL 报错/注入。"""
    return (value or "").replace("'", "''")


def _t(db: Db, dbname: str, table: str) -> str:
    """跨库表名：MSSQL 三段名 [db].dbo.T；ORACLE 同实例跨 schema 直接 db.T。"""
    if getattr(db, "provider", "mssql") == "oracle":
        return f"{dbname}.{table}"
    return f"[{dbname}].dbo.[{table}]"


def _gl(db: Db, guid: str) -> str:
    """GUID 字面量：MSSQL '{...}'；ORACLE HEXTORAW('大端 hex')。"""
    if getattr(db, "provider", "mssql") == "oracle":
        return f"HEXTORAW('{guid.replace('-', '').upper()}')"
    return f"'{guid}'"


def _oid_list(db: Db, oids: List[str]) -> str:
    # oid 字面量列表：Oracle 下 oid 列是 RAW(16)（大端原序，db.query 已按
    # RAWTOHEX 同款输出 plain hex），HEXTORAW 直出即物理字节；MSSQL '{oid}' 原样
    return ", ".join(_gl(db, o) for o in dict.fromkeys(oids))


def _rel(db: Db, prefix: str, alias: str = "") -> str:
    """关系类型前缀匹配条件：Oracle RAWTOHEX 前缀 LIKE；MSSQL 直接 LIKE。

    同一 SQL 出现多个 CoRERelationOrigin 别名时必须传 alias 限定列名（ORA-00918）。
    """
    col = f"{alias}.RelationType" if alias else "RelationType"
    if getattr(db, "provider", "mssql") == "oracle":
        return f"RAWTOHEX({col}) LIKE '{prefix}%'"
    return f"{col} LIKE '{prefix}%'"


# ------------------------------------------------------------- 数据载体 ----


@dataclass
class MatOptions:
    """一次采集的选项。"""
    sys_oids: List[str] = field(default_factory=list)   # 范围系统节点 oid
    sys_names: List[str] = field(default_factory=list)  # 显示名（日志用）
    hidden_only: bool = False  # True=仅统计隐藏材料（ReportingRequirements=2）
    dwg_folder_oids: List[str] = field(default_factory=list)  # ISO 图纸目录节点
    pid_kinds: Optional[frozenset] = None  # PID 关联对象类别白名单（None=全部）
    vol_oids: List[str] = field(default_factory=list)  # 范围 volume 条件引用的
    # 体积对象 oid（Named Spaces 条件；多体积并集语义：中心命中任一盒即计入）
    vol_boxes: List[List[float]] = field(default_factory=list)  # Planes 条件的
    # 平面组围合盒 [xmin,ymin,zmin,xmax,ymax,zmax]（组内全部平面聚合 MIN/MAX
    # 而来；单平面是零厚度盒，逐平面判定会失真必须聚合；多组间并集）


@dataclass
class MatRow:
    """聚合行：(管线, PartNo, run 属性组合) -> 数量。

    cls 为类型码（'pipe'=直管 / 'com'=管件 / 'instrument'=仪表件 /
    'specialty'=特殊件，类型表 JOIN 判定，失败回退本体表），用于判定单位
    （直管=米/其余=只）与仪表件/特殊件文案；weight 为组内有效
    DryWeight 总和（kg，直管=切段总重合计，件=单件重合计）；保温厚度
    已换算 mm。
    """
    line: str
    partno: str
    qty: float
    cls: str = ""
    spec: str = ""                    # 管道等级（SpecName）
    run_name: str = ""                # 管道名称（组内 MIN）
    op_p: Optional[float] = None      # 操作最大压力（MPa）
    op_t: Optional[float] = None      # 操作最大温度（℃）
    de_p: Optional[float] = None      # 设计最大压力（MPa）
    de_t: Optional[float] = None      # 设计最大温度（℃）
    ins_purpose: str = ""             # 保温目的（代码表翻译文本）
    ins_material: str = ""            # 保温材料（代码表翻译文本）
    ins_thk: Optional[float] = None   # 保温厚度 mm
    weight: Optional[float] = None    # 组内总重 kg（无有效重量为 None）
    unit_name: str = ""               # 单元名称（管线最近单元祖先）
    area_name: str = ""               # 区域名称（管线最近区域祖先）


# ------------------------------------------------------------- 材料采集 ----


_VOL_PT = ("s.xmin + s.xmax BETWEEN 2*v.xmin AND 2*v.xmax "
           "AND s.ymin + s.ymax BETWEEN 2*v.ymin AND 2*v.ymax "
           "AND s.zmin + s.zmax BETWEEN 2*v.zmin AND 2*v.zmax")


def _fnum(x: float) -> str:
    """浮点 SQL 字面量：repr 防科学计数法（1e-05 → 0.0001）。"""
    s = repr(float(x))
    return s if "e" not in s and "E" not in s else format(float(x), ".10f")


def _scope_sql(db: Db, opts: "MatOptions", run_col: str,
               mdb: str = "", vol_col: str = "",
               vol_mode: str = "") -> Tuple[str, str]:
    """过滤器范围 → (CTE 前缀, AND 条件)。run_col 为 run oid 列引用。

    mdb + vol_col 传入时追加 volume 交集，vol_mode 三种口径：
    ""（默认，材料表）——零件盒中心点落在勾选体积盒（MDB
    CORESpatialIndex）内才计入；零件无盒子行即未放置，天然排除。
    "self"（管夹）——对象自身盒中心点（点状盒即自身位置）判定，
    无盒子保留不过滤（找不到位置不硬丢）。
    "host"（螺栓/垫片）——经 XGeneratesConnectionItems 借宿主连接
    （origin，点状盒）的位置判定，宿主无盒子保留不过滤。"""
    if not opts.sys_oids and not opts.vol_oids and not opts.vol_boxes:
        return "", ""
    ctes: List[str] = []
    cond = ""
    if opts.sys_oids:
        ctes.append(f"""sysScope (oid) AS (
    SELECT h.oid FROM JNamedItem h WHERE h.oid IN ({_oid_list(db, opts.sys_oids)})
    UNION ALL
    SELECT x.oidDestination FROM XSystemHierarchy x JOIN sysScope s ON x.oidOrigin = s.oid
)""")
        cond += f" AND {run_col} IN (SELECT oid FROM sysScope)"
    if (opts.vol_oids or opts.vol_boxes) and mdb and vol_col:
        tbl = _t(db, mdb, "CORESpatialIndex")
        # volScope 两来源并集：Named 体积盒（oid 逐个）+ Planes 围合盒
        # （Python 端已聚合的字面量盒，VALUES/dual 双后端）
        parts = []
        if opts.vol_oids:
            parts.append(f"    SELECT v.xmin, v.ymin, v.zmin, v.xmax, v.ymax, v.zmax\n"
                         f"    FROM {tbl} v WHERE v.oid IN ({_oid_list(db, opts.vol_oids)})")
        if opts.vol_boxes:
            is_ora = getattr(db, "provider", "mssql") == "oracle"
            cols = ("xmin", "ymin", "zmin", "xmax", "ymax", "zmax")
            rows = []
            for b in opts.vol_boxes:
                nums = ", ".join(f"{_fnum(float(x))} AS {c}"
                                 for x, c in zip(b, cols))
                rows.append(f"    SELECT {nums} FROM dual" if is_ora
                            else f"    SELECT {nums}")
            joiner = "\n    UNION ALL\n" if is_ora else "\n  UNION ALL\n"
            parts.append("(" + joiner.join(rows) + "\n    )")
        ctes.append("volScope (xmin, ymin, zmin, xmax, ymax, zmax) AS (\n"
                    + "\n  UNION ALL\n".join(parts) + "\n)")
        if vol_mode == "host":
            # 宿主位置集合（有 = 可判定）；命中 = 宿主点落在任一勾选体积盒
            pos = (f"SELECT g.OidDestination FROM XGeneratesConnectionItems g "
                   f"JOIN {tbl} s ON s.oid = g.oidOrigin")
            hit = f"{pos} JOIN volScope v ON {_VOL_PT}"
        else:
            pos = f"SELECT s.oid FROM {tbl} s"
            hit = f"{pos} JOIN volScope v ON {_VOL_PT}"
            if vol_mode == "self":
                pos = f"SELECT oid FROM {tbl}"
            # 默认模式无盒即未放置，直接 IN 命中集合排除；
            # self 模式补"无盒保留"分支
        if vol_mode in ("self", "host"):
            cond += (f" AND ({vol_col} IN ({hit})"
                     f" OR {vol_col} NOT IN ({pos}))")
        else:
            cond += f" AND {vol_col} IN ({hit})"
    return "WITH " + ",\n".join(ctes) + "\n", cond


def extract_materials(db: Db, plant: dict,
                      opts: MatOptions,
                      on_log: Callable[[str], None] = None) -> List[MatRow]:
    """采集管道材料清单，返回按 (管线, PartNo) 聚合的行。

    plant 为 list_plants 的项目 dict（需含 reports/model/schema 键）。
   mdb_schema 即 plant["model"]（Oracle 跨 schema 直查用）。
    """
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    nvl = "NVL" if is_ora else "ISNULL"
    mdb = plant["model"]

    # 范围 CTE（weldpy 同款：勾选节点 + 全部子孙，恒取范围内）；
    # volume 交集仅材料表生效（vol_col=零件实例，见 _scope_sql 注释）
    scope_cte, scope_cond = _scope_sql(db, opts, "f.oidOrigin",
                                       mdb=mdb, vol_col="pp.oid")

    # 部件实例 → 目录定义（madeFrom）；PartNumber 两路：管道目录件表
    # 直取 / 命名 def 载体名（PART2DEF 目标仅这两类，分流依据见 skill）
    if is_ora:
        def_join = f"""JOIN {mdb}.CoRERelationOrigin d
       ON d.oid = pp.oid AND {_rel(db, _REL_PART2DEF, "d")}
LEFT JOIN {mdb}.REFDATPipeComponent pm
       ON pm.oid = d.oidTarget
LEFT JOIN {mdb}.CoRERelationOrigin cn
       ON cn.oidTarget = d.oidTarget AND {_rel(db, _REL_NAME_CARRIER, "cn")}"""
        # 各列 NVARCHAR2/VARCHAR2 混拼会 ORA-12704，统一 TO_NCHAR
        partno_expr = ("COALESCE(TO_NCHAR(pm.PartNumber), "
                       "TO_NCHAR(cn.RelationName))")
        concat = "||"
    else:
        def_join = f"""JOIN {_t(db, mdb, 'CoRERelationOrigin')} d
       ON d.oid = pp.oid AND {_rel(db, _REL_PART2DEF, "d")}
LEFT JOIN {_t(db, mdb, 'REFDATPipeComponent')} pm
       ON pm.oid = d.oidTarget
LEFT JOIN {_t(db, mdb, 'CoRERelationOrigin')} cn
       ON cn.oidTarget = d.oidTarget AND {_rel(db, _REL_NAME_CARRIER, "cn")}"""
        partno_expr = "ISNULL(pm.PartNumber, cn.RelationName)"
        concat = "+"

    # 特殊件/仪表：PartNo 带 JNamedItem 对象名后缀（PNo|Name 键）；
    # 类型判定默认 JOIN MDB 类型表，失败回退 RDB 本体表（成员相等）
    def _type_frags(typing: str) -> tuple:
        if typing == "table":
            type_join = (f"LEFT JOIN {_t(db, mdb, 'ROUTEPipeInstrumentOcc')} io"
                         f" ON io.Oid = pp.oid\n   "
                         f"LEFT JOIN {_t(db, mdb, 'ROUTEPipeSpecialtyOcc')} so"
                         f" ON so.Oid = pp.oid")
        else:
            type_join = ("LEFT JOIN JRteInstrument io ON io.Oid = pp.oid\n"
                         "LEFT JOIN JRteSpecialtyOccur so ON so.Oid = pp.oid")
        named_cond = "(io.Oid IS NOT NULL OR so.Oid IS NOT NULL)"
        is_pipe = "sp.oid IS NOT NULL"
        type_expr = ("CASE WHEN sp.oid IS NOT NULL THEN 'pipe' "
                     "WHEN io.Oid IS NOT NULL THEN 'instrument' "
                     "WHEN so.Oid IS NOT NULL THEN 'specialty' "
                     "ELSE 'com' END")
        return type_join, named_cond, is_pipe, type_expr

    if is_ora:
        # CASE 两分支也要统一 NVARCHAR2，混 VARCHAR2 同样 ORA-12704
        suffix_expr = ("CASE WHEN nh.oid IS NOT NULL "
                       "THEN TO_NCHAR('|') || nh.ItemName "
                       "ELSE TO_NCHAR('') END")
    else:
        suffix_expr = ("CASE WHEN nh.oid IS NOT NULL "
                       "THEN '|' + nh.ItemName ELSE '' END")

    # 管道等级：MSSQL 走 RDB 视图；Oracle 视图慢改 MDB 关系直查，
    # SpecName 白名单对齐视图 INNER JOIN 的丢弃语义。
    # Oracle 不能把 IN 子查询放进 GROUP BY（ORA-00979，SELECT/GROUP BY
    # 表达式逐字相同也报），故白名单改 LEFT JOIN JDPipeSpec 等价实现
    # （SpecName 唯一无扇出；TO_NCHAR 两侧统一防 ORA-12704）
    if is_ora:
        spec_join = f"""LEFT JOIN {mdb}.CoRERelationOrigin r1
       ON r1.oid = f.oidOrigin AND {_rel(db, _REL_RUNUSESSPEC, "r1")}
LEFT JOIN {mdb}.CoRERelationOrigin r2
       ON r2.oidTarget = r1.oidTarget AND {_rel(db, _REL_NAME_CARRIER, "r2")}
LEFT JOIN JDPipeSpec sd ON TO_NCHAR(sd.SpecName) = TO_NCHAR(r2.RelationName)"""
        spec_expr = "TO_NCHAR(sd.SpecName)"
        pipe_oid_expr = "MIN(RAWTOHEX(g.oidOrigin))"
    else:
        spec_join = """LEFT JOIN XPathRunUsesSpec xs ON xs.oidOrigin = f.oidOrigin
LEFT JOIN JDPipeSpec sd ON sd.oid = xs.oidDestination"""
        spec_expr = "sd.SpecName"
        pipe_oid_expr = "MIN(CONVERT(VARCHAR(36), g.oidOrigin))"

    # 单位重量守卫：仪表件 DryWeight 存 1.6e308 哨兵值，直接 SUM 会浮点溢出
    weight_expr = (f"SUM(CASE WHEN w.DryWeight > 1E30 OR w.DryWeight < -1E30 "
                   f"THEN NULL ELSE w.DryWeight END)")

    def _build_sql(typing: str) -> str:
        type_join, named_cond, is_pipe, type_expr = _type_frags(typing)
        name_join = f"""LEFT JOIN JNamedItem nh ON nh.oid = pp.oid
      AND {named_cond}"""
        return f"""{scope_cte}SELECT h.ItemName AS lineName,
       {partno_expr} {concat} {suffix_expr} AS partNo,
       {spec_expr} AS specName,
       MIN(rn.ItemName) AS runName,
       {pipe_oid_expr} AS pipeOid,
       MAX({type_expr}) AS typeCode,
       SUM(CASE WHEN {is_pipe}
                THEN {nvl}(sp.CutLength, 0) ELSE 1 END) AS qty,
       {weight_expr} AS weight,
       pd.OperatingMaxPressure AS opP, pd.OperatingMaxTemp AS opT,
       pd.DesignMaxPressure AS deP, pd.DesignMaxTemp AS deT,
       ri.InsulationPurpose AS insP, ri.InsulationMaterial AS insM,
       ri.InsulationThickness AS insT
FROM JRtePipePart pp
{def_join}
{type_join if type_join else ''}
{name_join}
LEFT JOIN JRteStockPartOccur sp ON sp.oid = pp.oid
JOIN XOwnsParts f ON f.oidDestination = pp.oid
JOIN XSystemHierarchy g ON g.oidDestination = f.oidOrigin
JOIN JNamedItem h ON h.oid = g.oidOrigin
LEFT JOIN JNamedItem rn ON rn.oid = f.oidOrigin
LEFT JOIN JProcessDataInfo pd ON pd.Oid = f.oidOrigin
LEFT JOIN JRteInsulation ri ON ri.Oid = f.oidOrigin
LEFT JOIN JWeightCG w ON w.Oid = pp.oid
{spec_join}
LEFT JOIN JMTOInfo mto ON mto.Oid = pp.oid
WHERE {nvl}(mto.ReportingRequirements, 5) {'=' if opts.hidden_only else '<>'} 2{scope_cond}
GROUP BY h.ItemName, {partno_expr} {concat} {suffix_expr}, {spec_expr},
         pd.OperatingMaxPressure, pd.OperatingMaxTemp,
         pd.DesignMaxPressure, pd.DesignMaxTemp,
         ri.InsulationPurpose, ri.InsulationMaterial, ri.InsulationThickness"""

    rows = None
    err = None
    for typing in ("table", "prefix"):
        try:
            rows = db.query(_build_sql(typing))
            if typing == "prefix":
                log("类型表查询失败，已回退 RDB 本体表判定")
            break
        except Exception as ex:
            err = ex
    if rows is None:
        raise err
    anc = load_system_ancestors(db)
    ins_p_texts, ins_m_texts = _insulation_texts(db)

    def _f(v) -> Optional[float]:
        return float(v) if v is not None else None

    # 过程数据单位换算：库内温度开尔文、压力帕斯卡（多库实测 x.15 尾数
    # 即 273.15 换算指纹），输出统一摄氏度 / 兆帕
    def _t_c(v) -> Optional[float]:
        return _f(v) - 273.15 if v is not None else None

    def _p_m(v) -> Optional[float]:
        return _f(v) / 1_000_000.0 if v is not None else None

    def _txt(texts: Dict[int, str], v) -> str:
        # 保温目的/材料：-1/NULL=未设；翻译不到回退 id 文本
        if v is None or int(v) <= 0:
            return ""
        i = int(v)
        return texts.get(i, str(i))

    out: List[MatRow] = []
    for (line, partno, spec, run_name, pipe_oid, cls, qty, wt,
         op_p, op_t, de_p, de_t, ins_p, ins_m, ins_t) in rows:
        unit, area = anc.get(str(pipe_oid or ""), ("", ""))
        out.append(MatRow(
            line=str(line or ""), partno=str(partno or ""), qty=float(qty or 0),
            cls=str(cls or ""), spec=str(spec or ""),
            run_name=str(run_name or ""),
            op_p=_p_m(op_p), op_t=_t_c(op_t),
            de_p=_p_m(de_p), de_t=_t_c(de_t),
            ins_purpose=_txt(ins_p_texts, ins_p),
            ins_material=_txt(ins_m_texts, ins_m),
            ins_thk=(float(ins_t) * 1000.0) if ins_t is not None else None,
            weight=_f(wt),
            unit_name=unit, area_name=area))
    log(f"采集完成：{len(out)} 个 (管线, PartNo, run 属性) 聚合行")
    return out


def load_system_ancestors(db: Db) -> Dict[str, Tuple[str, str]]:
    """管线 oid → (单元名, 区域名)：自 JUnitSystem/JAreaSystem 种子沿
    XSystemHierarchy 向下递归，每节点取最近的单元/区域祖先（树结构下与
    VB 自管线向上首遇等价；向下展开一趟同时算出全部节点，深度上限 32
    防脏数据成环）。查不到类型表的库返回空 dict（列留空不阻断）。
    """
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    conv = ("RAWTOHEX(d.node)" if is_ora else "CONVERT(VARCHAR(36), d.node)")
    # 种子拆独立 CTE：Oracle 递归 WITH 的 UNION ALL 只允许两分支
    # （一个锚点+一个递归，ORA-32041），两种子并列会整体报错被下面
    # except 吞成空 dict（单元/区域列全空）；MSSQL 无此限制，两后端通用
    sql = f"""WITH seeds (root, kind) AS (
    SELECT s.Oid, 'U' FROM JUnitSystem s
    UNION ALL
    SELECT a.Oid, 'A' FROM JAreaSystem a
),
down (root, kind, node, depth) AS (
    SELECT root, kind, root, 0 FROM seeds
    UNION ALL
    SELECT d.root, d.kind, x.oidDestination, d.depth + 1
      FROM down d JOIN XSystemHierarchy x ON x.oidOrigin = d.node
     WHERE d.depth < 32
)
SELECT {conv} AS node, d.kind, j.ItemName AS name, d.depth
FROM down d JOIN JNamedItem j ON j.oid = d.root"""
    try:
        rows = db.query(sql)
    except Exception:
        return {}
    best: Dict[str, Dict[str, Tuple[int, str]]] = {}
    for node, kind, name, depth in rows:
        cur = best.setdefault(str(node), {})
        d = int(depth)
        if kind not in cur or d < cur[kind][0]:
            cur[kind] = (d, str(name or ""))
    return {n: (v.get("U", (0, ""))[1], v.get("A", (0, ""))[1])
            for n, v in best.items()}


def _insulation_texts(db: Db) -> Tuple[Dict[int, str], Dict[int, str]]:
    """保温目的/材料代码表 id → 文本（RDB 小表直查；Oracle 无 CL_ 表时回退空）。"""
    def _load(tbl: str) -> Dict[int, str]:
        try:
            return {int(r[0]): str(r[1] or "") for r in db.query(
                f"SELECT ValueID, ShortStringValue FROM {tbl}")}
        except Exception:
            return {}
    return _load("CL_InsulationPurpose"), _load("CL_InsulationMaterial")


# ------------------------------------------------------------- 目录属性 ----

# CommodityClass 文本 → 排序值（对齐 VB frmReportPipePart 的 PartSortIndex）
_CLASS_SORT = {"Straight sections": 1, "Direction Change Fittings": 2,
               "Branch Fittings": 3, "In-Line Fittings": 4,
               "End Fittings": 5, "Valves": 6}


def load_catalog_attrs(db: Db, cdb_name: str, schema_db: str,
                       partnos: Set[str],
                       on_log: Callable[[str], None] = None) -> Dict[str, dict]:
    """按 PartNumber 批量查 CDB 目录属性，返回 {partno: {...}}。

    属性：commodity_code / long_desc / primary / secondary / sch1 / sch2 /
    industry_code(CC码) / commodity_type(部件类型) / class_sort(排序值)。
    sch1/sch2/部件类型/大类为代码表 id，经 Schema CodelistValueView 翻译，
    细节见 skill「PartNumber 链」；取不到记录的不返回，调用方回退。
    """
    log = on_log or (lambda m: None)
    if not partnos:
        return {}
    # 带名称后缀的键取 | 前段（CLB Split(PartNo,"|")(0)）
    keys = sorted({p.split("|")[0] for p in partnos if p.split("|")[0]})

    # 代码表翻译（Schema 库视图）；schema 库名缺失时不翻译、回退 id
    if schema_db:
        clv = _t(db, schema_db, 'CodelistValueView')
        sch_join = (f"LEFT JOIN {clv} cv1 "
                    f"ON cv1.TableName = 'ScheduleThickness' "
                    f"AND cv1.ValueID = A.FirstSizeSchedule "
                    f"LEFT JOIN {clv} cv2 "
                    f"ON cv2.TableName = 'ScheduleThickness' "
                    f"AND cv2.ValueID = A.SecondSizeSchedule "
                    f"LEFT JOIN {clv} ct "
                    f"ON ct.TableName = 'PipingCommodityType' "
                    f"AND ct.ValueID = A.CommodityType "
                    f"LEFT JOIN {clv} cc "
                    f"ON cc.TableName = 'PipingCommodityClass' "
                    f"AND cc.ValueID = A.CommodityClass")
        sch_cols = ("cv1.ShortStringValue AS Sch1Text, "
                    "cv2.ShortStringValue AS Sch2Text, "
                    "ct.ShortStringValue AS TypeText, "
                    "cc.ShortStringValue AS ClassText")
    else:
        sch_join = ""
        sch_cols = ("NULL AS Sch1Text, NULL AS Sch2Text, "
                    "NULL AS TypeText, NULL AS ClassText")

    def _sch(text, raw) -> str:
        t = str(text or "").strip()
        if t:
            return t
        return str(raw) if raw is not None else ""

    out: Dict[str, dict] = {}
    B = 200  # IN 分批（MSSQL 参数上限稳妥值）
    for i in range(0, len(keys), B):
        lst = ", ".join(f"'{safe_sql(k)}'" for k in keys[i:i + B])
        sql = (f"SELECT A.PartNumber, M.ContractorCommodityCode, "
               f"M.LongMaterialDescription, A.PrimarySize, A.SecondarySize, "
               f"A.FirstSizeSchedule, A.SecondSizeSchedule, {sch_cols}, "
               f"A.IndustryCommodityCode, A.CommodityType "
               f"FROM {_t(db, cdb_name, 'REFDATPipeComponent')} A "
               f"LEFT JOIN {_t(db, cdb_name, 'CORErelations')} CR "
               f"ON CR.IDorigin = A.oid "
               f"LEFT JOIN {_t(db, cdb_name, 'REFDATCOMMODITYMATLCTRLDATA')} M "
               f"ON M.oid = CR.IDdestination "
               f"{sch_join} "
               f"WHERE A.PartNumber IN ({lst})")
        for (pn, code, desc, p1, p2, s1, s2, t1, t2, ttype, tclass,
             ind_code, type_id) in db.query(sql):
            pn = str(pn or "")
            if pn in out:
                # 同名多行时优先保留描述齐全的（对齐 CLB GetPart 唯一性）
                if out[pn].get("long_desc") and not desc:
                    continue
            try:
                p1f, p2f = float(p1 or 0), float(p2 or 0)
            except (TypeError, ValueError):
                p1f = p2f = 0.0
            type_txt = _sch(ttype, type_id if type_id is not None else "")
            out[pn] = {"commodity_code": str(code or pn),
                       "long_desc": str(desc or "") if desc else "",
                       "primary": p1f, "secondary": p2f,
                       "sch1": _sch(t1, s1), "sch2": _sch(t2, s2),
                       "industry_code": str(ind_code or ""),
                       "commodity_type": type_txt,
                       "class_sort": _CLASS_SORT.get(str(tclass or ""), 7)}
    log(f"CDB 目录属性：{len(out)}/{len(keys)} 个 PartNumber 查到")
    return out


# --------------------------------------------------------- 螺栓/垫片报表 ----
# 对齐 frmReportBolt/frmReportGasket（隐含件明细，逐套/逐片一行不聚合）。
# 链路：XImpliedMatingParts 连接两侧 → 隐含件 → 目录 def（REFDATBoltPart /
# REFDATGasketPart）；实例属性 JRteBolt/JRteGasket；业主采购码 JDProcurementDataInfo
# 无索引不能 JOIN 主 SQL，Python 端整表小读映射。细节见项目 skill matpy-summary。


@dataclass
class BoltRow:
    """螺栓明细行（一套螺栓一行）。"""
    oid: str
    diameter: str        # 螺栓直径（公制 M14；VB 默认公制档）
    length_mm: float     # 长度 mm（RoundedLength×1000）
    qty: int             # 每套数量（BoltQuantity）
    cc_code: str
    long_desc: str
    owner_code: str      # 业主编码（SizedCommodityCode）
    part_pns: str        # 两侧部件码（去重逗号拼）
    run_names: str       # 两侧管道名（去重）
    line_names: str      # 两侧管线名（去重）
    unit_name: str = ""
    area_name: str = ""


@dataclass
class GasketRow:
    """垫片明细行（一片一行）。"""
    oid: str
    nd: float            # 通径（目录 NominalDiameterFrom，空回退部件 PrimarySize）
    cc_code: str
    long_desc: str
    owner_code: str      # 业主采购码
    part_pns: str
    run_names: str
    line_names: str
    unit_name: str = ""
    area_name: str = ""


def _implied_tail(db: Db, mdb: str) -> Tuple[str, str, str]:
    """螺栓/垫片共用尾段：连接两侧部件 → 部件码/管道/管线/管线 oid。

    返回 (tail_from, partno_expr, pipe_oid_expr)，引用别名：g(连接关系)、
    cp(连接部件关系)。
    """
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    if is_ora:
        partno = "NVL(TO_NCHAR(A.PartNumber), TO_NCHAR(cn.RelationName))"
        pipe_oid = "RAWTOHEX(hh.oidOrigin)"
    else:
        partno = "ISNULL(A.PartNumber, cn.RelationName)"
        pipe_oid = "CONVERT(VARCHAR(36), hh.oidOrigin)"
    tail = f"""JOIN XRelConnectionAndPartOcc cp ON cp.OidOrigin = g.OidOrigin
JOIN {_t(db, mdb, 'CoRERelationOrigin')} d
   ON d.oid = cp.OidDestination AND {_rel(db, _REL_PART2DEF, "d")}
LEFT JOIN {_t(db, mdb, 'REFDATPipeComponent')} A
   ON A.oid = d.oidTarget
LEFT JOIN {_t(db, mdb, 'CoRERelationOrigin')} cn
   ON cn.oidTarget = d.oidTarget AND {_rel(db, _REL_NAME_CARRIER, "cn")}
JOIN XOwnsParts op ON op.oidDestination = cp.OidDestination
JOIN JNamedItem rn ON rn.oid = op.oidOrigin
JOIN XSystemHierarchy hh ON hh.oidDestination = op.oidOrigin
JOIN JNamedItem h ON h.oid = hh.oidOrigin"""
    return tail, partno, pipe_oid


def _merge_sides(groups: dict, oid: str) -> Tuple[str, str]:
    """两侧值去重逗号拼（同侧重名只留一个，对齐 VB 的去重拼接）。"""
    pns, runs, lines = [], [], []
    for pn, run, line, _pipe_oid in groups.get(oid, ()):
        for seq, v in ((pns, pn), (runs, run), (lines, line)):
            if v and v not in seq:
                seq.append(v)
    return ",".join(pns), ",".join(runs) + "|" + ",".join(lines)


def extract_bolts(db: Db, plant: dict, opts: MatOptions,
                  on_log: Callable[[str], None] = None) -> List[BoltRow]:
    """提取螺栓明细（排除 No-MTO），范围按连接部件所属 run 的管线过滤。"""
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    nvl = "NVL" if is_ora else "ISNULL"
    mdb, cdb = plant["model"], plant["catalog"]
    oid_out = "RAWTOHEX(bs.Oid)" if is_ora else "CONVERT(VARCHAR(36), bs.Oid)"
    # volume：借宿主连接点判定（98.5% 有宿主盒，无宿主盒保留）
    scope_cte, scope_cond = _scope_sql(db, opts, "op.oidOrigin", mdb=mdb,
                                       vol_col="bs.Oid", vol_mode="host")
    tail, partno_expr, pipe_oid_expr = _implied_tail(db, mdb)

    sql = f"""{scope_cte}SELECT {oid_out} AS objOid,
       b.Diameter, b.RoundedLength, b.BoltQuantity,
       bp.IndustryCommodityCode, m.LongMaterialDescription,
       ju.SizedCommodityCode,
       {partno_expr} AS partPn, rn.ItemName AS runName,
       h.ItemName AS lineName, {pipe_oid_expr} AS pipeOid
FROM JRteBoltSet bs
JOIN JRteBolt b ON b.Oid = bs.Oid
LEFT JOIN JUASizedBoltData ju ON ju.Oid = bs.Oid
JOIN XGeneratesConnectionItems g ON g.OidDestination = bs.Oid
JOIN XImpliedMatingParts im ON im.OidOrigin = bs.Oid
JOIN {_t(db, cdb, 'REFDATBoltPart')} bp ON bp.oid = im.OidDestination
LEFT JOIN {_t(db, cdb, 'CORErelations')} cr ON cr.IDorigin = bp.oid
LEFT JOIN {_t(db, cdb, 'REFDATCOMMODITYMATLCTRLDATA')} m ON m.oid = cr.IDdestination
{tail}
WHERE {nvl}(b.BoltReportingClassification, 5) {'=' if opts.hidden_only else '<>'} 2{scope_cond}"""
    rows = db.query(sql)
    anc = load_system_ancestors(db)

    sides: Dict[str, list] = {}
    head: Dict[str, tuple] = {}
    for (oid, dia, rlen, bqty, cc, ldesc, owner,
         pn, run, line, pipe_oid) in rows:
        head[str(oid)] = (dia, rlen, bqty, cc, ldesc, owner)
        sides.setdefault(str(oid), []).append(
            (str(pn or ""), str(run or ""), str(line or ""),
             str(pipe_oid or "")))

    out: List[BoltRow] = []
    for oid, (dia, rlen, bqty, cc, ldesc, owner) in head.items():
        pns, runs_lines = _merge_sides(sides, oid)
        runs, lines = runs_lines.split("|", 1)
        unit, area = _anc_of_sides(sides, oid, anc)
        d_mm = round(float(dia or 0) * 1000)
        out.append(BoltRow(
            oid=oid, diameter=f"M{d_mm}",
            length_mm=round(float(rlen or 0) * 1000, 1),
            qty=int(bqty or 0), cc_code=str(cc or ""),
            long_desc=str(ldesc or ""), owner_code=str(owner or ""),
            part_pns=pns, run_names=runs, line_names=lines,
            unit_name=unit, area_name=area))
    out.sort(key=lambda r: (r.line_names, r.run_names, r.cc_code))
    log(f"螺栓明细：{len(out)} 套（"
        f"{'仅 No-MTO' if opts.hidden_only else 'No-MTO 已排除'}）")
    return out


def extract_gaskets(db: Db, plant: dict, opts: MatOptions,
                    on_log: Callable[[str], None] = None) -> List[GasketRow]:
    """提取垫片明细（排除 No-MTO）：XImpliedMatingParts 直连目录。"""
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    nvl = "NVL" if is_ora else "ISNULL"
    mdb, cdb = plant["model"], plant["catalog"]
    oid_out = "RAWTOHEX(gs.Oid)" if is_ora else "CONVERT(VARCHAR(36), gs.Oid)"
    def_out = ("RAWTOHEX(im.OidDestination)" if is_ora
               else "CONVERT(VARCHAR(36), im.OidDestination)")
    # volume：借宿主连接点判定（100% 有宿主盒）
    scope_cte, scope_cond = _scope_sql(db, opts, "op.oidOrigin", mdb=mdb,
                                       vol_col="gs.Oid", vol_mode="host")
    tail, partno_expr, pipe_oid_expr = _implied_tail(db, mdb)

    # 注意：业主采购码所在的 CDB JDProcurementDataInfo 无索引，JOIN 进主
    # SQL 会坏计划挂死（实测 20s 超时），改 Python 端整表小读取后映射
    sql = f"""{scope_cte}SELECT {oid_out} AS objOid,
       {def_out} AS defOid,
       gp.IndustryCommodityCode, gp.NominalDiameterFrom,
       m.LongMaterialDescription,
       gs.GasketSizedCommodityCode, A.PrimarySize AS partNpd,
       {partno_expr} AS partPn, rn.ItemName AS runName,
       h.ItemName AS lineName, {pipe_oid_expr} AS pipeOid
FROM JRteGasket gs
JOIN XGeneratesConnectionItems g ON g.OidDestination = gs.Oid
JOIN XImpliedMatingParts im ON im.OidOrigin = gs.Oid
JOIN {_t(db, cdb, 'REFDATGasketPart')} gp ON gp.oid = im.OidDestination
LEFT JOIN {_t(db, cdb, 'CORErelations')} cr ON cr.IDorigin = gp.oid
LEFT JOIN {_t(db, cdb, 'REFDATCOMMODITYMATLCTRLDATA')} m ON m.oid = cr.IDdestination
{tail}
WHERE {nvl}(gs.GasketReportingClassification, 5) {'=' if opts.hidden_only else '<>'} 2{scope_cond}"""
    rows = db.query(sql)
    anc = load_system_ancestors(db)
    proc_codes: Dict[str, str] = {}
    try:
        proc_codes = {
            str(o): str(c or "")
            for o, c in db.query(
                f"SELECT oid, ProcurementClientCommodityCode "
                f"FROM {_t(db, cdb, 'JDProcurementDataInfo')}")}
    except Exception:
        proc_codes = {}

    sides: Dict[str, list] = {}
    head: Dict[str, tuple] = {}
    for (oid, def_oid, cc, ndf, ldesc, sized, part_npd,
         pn, run, line, pipe_oid) in rows:
        head[str(oid)] = (str(def_oid or ""), cc, ndf, ldesc, sized, part_npd)
        sides.setdefault(str(oid), []).append(
            (str(pn or ""), str(run or ""), str(line or ""),
             str(pipe_oid or "")))

    def _f(v) -> Optional[float]:
        try:
            return float(v) if v is not None else None
        except (TypeError, ValueError):
            return None

    out: List[GasketRow] = []
    for oid, (def_oid, cc, ndf, ldesc, sized, part_npd) in head.items():
        pns, runs_lines = _merge_sides(sides, oid)
        runs, lines = runs_lines.split("|", 1)
        unit, area = _anc_of_sides(sides, oid, anc)
        # 通径：目录 NominalDiameterFrom，空则回退连接部件 PrimarySize
        nd = _f(ndf)
        if nd is None:
            nd = _f(part_npd)
        out.append(GasketRow(
            oid=oid, nd=nd if nd is not None else 0.0,
            cc_code=str(cc or ""), long_desc=str(ldesc or ""),
            owner_code=proc_codes.get(def_oid, "") or str(sized or ""),
            part_pns=pns, run_names=runs, line_names=lines,
            unit_name=unit, area_name=area))
    out.sort(key=lambda r: (r.line_names, r.run_names, r.cc_code))
    log(f"垫片明细：{len(out)} 片（"
        f"{'仅 No-MTO' if opts.hidden_only else 'No-MTO 已排除'}）")
    return out


@dataclass
class ClampRow:
    """管夹明细行（一只一行，机械连接卡箍，非支吊架）。"""
    oid: str
    npd: float            # 通径（连接部件 PrimarySize）
    spec: str             # 管道等级名（run → spec）
    cc_code: str          # 商品代码（等级 filter 的 ContractorCommodityCode）
    clamp_option: str     # 型号名（CL_ClampOption 短名）
    part_pns: str
    run_names: str
    line_names: str
    unit_name: str = ""
    area_name: str = ""


def extract_clamps(db: Db, plant: dict, opts: MatOptions,
                   on_log: Callable[[str], None] = None) -> List[ClampRow]:
    """提取管夹明细：与螺栓/垫片不同，无实例级目录 def——XImpliedMatingParts
    目的地是 MDB 内置轻量对象（无物料）；物料在等级层：run → spec →
    XSpecDefinesClamp → REFDATClampSelectionFilter（CDB，按通径段匹配）。
    材料范围走 JMTOInfo 通用判定表（管夹实例全部挂行，与材料表同口径）。
    """
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    nvl = "NVL" if is_ora else "ISNULL"
    mdb, cdb = plant["model"], plant["catalog"]
    oid_out = "RAWTOHEX(m.Oid)" if is_ora else "CONVERT(VARCHAR(36), m.Oid)"
    # volume：管夹自身盒（1781/1781 点状）判定，无盒保留
    scope_cte, scope_cond = _scope_sql(db, opts, "op.oidOrigin", mdb=mdb,
                                       vol_col="m.Oid", vol_mode="self")
    tail, partno_expr, pipe_oid_expr = _implied_tail(db, mdb)
    # run → spec 名（MSSQL 走 RDB 视图；Oracle 视图慢改关系两跳直查，
    # 同 extract_materials——A53D19B6 目标是命名载体，须再经 5280312B
    # 拿 SpecName，不能直接当 spec oid 用）
    if is_ora:
        spec_join = (f"LEFT JOIN {mdb}.CoRERelationOrigin r1\n   "
                     f"ON r1.oid = op.oidOrigin "
                     f"AND {_rel(db, _REL_RUNUSESSPEC, 'r1')}\n   "
                     f"LEFT JOIN {mdb}.CoRERelationOrigin r2\n   "
                     f"ON r2.oidTarget = r1.oidTarget "
                     f"AND {_rel(db, _REL_NAME_CARRIER, 'r2')}")
        spec_out = "TO_NCHAR(r2.RelationName)"
    else:
        spec_join = ("LEFT JOIN XPathRunUsesSpec xsc\n   "
                     "ON xsc.oidOrigin = op.oidOrigin\n   "
                     "LEFT JOIN JDPipeSpec sd ON sd.oid = xsc.oidDestination")
        spec_out = "sd.SpecName"

    sql = f"""{scope_cte}SELECT {oid_out} AS objOid,
       A.PrimarySize AS partNpd, m.ClampOption, {spec_out} AS specName,
       {partno_expr} AS partPn, rn.ItemName AS runName,
       h.ItemName AS lineName, {pipe_oid_expr} AS pipeOid
FROM JRteMechImpliedItem m
JOIN XGeneratesConnectionItems g ON g.OidDestination = m.Oid
LEFT JOIN JMTOInfo mt ON mt.Oid = m.Oid
{tail}
{spec_join}
WHERE {nvl}(mt.ReportingRequirements, 5) {'=' if opts.hidden_only else '<>'} 2{scope_cond}"""
    rows = db.query(sql)
    anc = load_system_ancestors(db)

    # CDB 小读：等级名 → 管夹 filter 定义（XSpecDefinesClamp 挂的是
    # REFDATPipeSpec.oid，与 Oracle 关系链的命名载体不同源，按 SpecName
    # 对齐；缺表回退空）+ 型号枚举
    spec_defs: Dict[str, list] = {}
    try:
        for (name, f, t, cc, opt, prio) in db.query(
                f"SELECT p.SpecName, f.NominalDiameterFrom, "
                f"f.NominalDiameterTo, f.ContractorCommodityCode, "
                f"f.ClampOption, f.Priority "
                f"FROM {_t(db, cdb, 'XSpecDefinesClamp')} s "
                f"JOIN {_t(db, cdb, 'REFDATPipeSpec')} p "
                f"ON p.oid = s.OidOrigin "
                f"JOIN {_t(db, cdb, 'REFDATClampSelectionFilter')} f "
                f"ON f.oid = s.OidDestination"):
            spec_defs.setdefault(str(name or ""), []).append(
                (f, t, cc, opt, prio))
    except Exception:
        spec_defs = {}
    options: Dict[int, str] = {}
    try:
        options = {int(r[0]): str(r[1] or "") for r in db.query(
            "SELECT ValueID, ShortStringValue FROM CL_ClampOption")}
    except Exception:
        options = {}

    def _f(v) -> Optional[float]:
        try:
            return float(v) if v is not None else None
        except (TypeError, ValueError):
            return None

    def _pick(flts, npd):
        """按通径段匹配 filter；npd 缺失或无命中取整表，多条按 Priority 升序。"""
        if not flts:
            return None
        hit = [x for x in flts
               if npd is None or (x[0] or 0) <= npd <= (x[1] or 1e9)]
        if not hit:
            hit = flts
        return min(hit, key=lambda x: (x[4] or 0,))

    sides: Dict[str, list] = {}
    head: Dict[str, tuple] = {}
    for (oid, npd, opt_id, spec_name,
         pn, run, line, pipe_oid) in rows:
        head[str(oid)] = (_f(npd), opt_id, str(spec_name or ""))
        sides.setdefault(str(oid), []).append(
            (str(pn or ""), str(run or ""), str(line or ""),
             str(pipe_oid or "")))

    # npd 兜底：命名载体 def 无 PrimarySize，按部件码批查目录件表
    # （load_catalog_attrs 同款 IN 分批；PrimarySize 列 PartNumber 可查到）
    npd_by_pn: Dict[str, float] = {}
    miss = {pn.split("|")[0]
            for oid, (npd, _o, _s) in head.items() if npd is None
            for pn, *_rest in sides.get(oid, ()) if pn.split("|")[0]}
    miss.discard("")
    B = 200
    keys = sorted(miss)
    for i in range(0, len(keys), B):
        lst = ", ".join(f"'{safe_sql(k)}'" for k in keys[i:i + B])
        try:
            npd_by_pn.update({
                str(pn or ""): float(sz or 0) for pn, sz in db.query(
                    f"SELECT PartNumber, PrimarySize "
                    f"FROM {_t(db, cdb, 'REFDATPipeComponent')} "
                    f"WHERE PartNumber IN ({lst})")})
        except Exception:
            break

    out: List[ClampRow] = []
    for oid, (npd, opt_id, spec_name) in head.items():
        pns, runs_lines = _merge_sides(sides, oid)
        runs, lines = runs_lines.split("|", 1)
        unit, area = _anc_of_sides(sides, oid, anc)
        if npd is None:  # 取两侧首个查到通径的部件（卡箍等径）
            npd = next((v for pn, *_r in sides.get(oid, ())
                        for v in [npd_by_pn.get(pn.split("|")[0])]
                        if v), None)
        flt = _pick(spec_defs.get(spec_name, []), npd)
        cc = flt[2] if flt else ""
        if flt and opt_id is None:
            opt_id = flt[3]
        out.append(ClampRow(
            oid=oid, npd=npd if npd is not None else 0.0,
            spec=spec_name,
            cc_code=str(cc or ""),
            clamp_option=options.get(int(opt_id or 0), str(opt_id or "")),
            part_pns=pns, run_names=runs, line_names=lines,
            unit_name=unit, area_name=area))
    out.sort(key=lambda r: (r.line_names, r.run_names, r.cc_code))
    hit = sum(1 for r in out if r.cc_code)
    log(f"管夹明细：{len(out)} 只（{hit} 只匹配到等级定义，"
        f"{'仅 No-MTO' if opts.hidden_only else 'No-MTO 已排除'}）")
    return out


def _anc_of_sides(sides: dict, oid: str,
                  anc: Dict[str, Tuple[str, str]]) -> Tuple[str, str]:
    """两侧管线 oid 中首个有祖先映射的 (单元, 区域)。"""
    for _pn, _run, _line, pipe_oid in sides.get(oid, []):
        got = anc.get(pipe_oid)
        if got:
            return got
    return ("", "")


# ------------------------------------------------------ 仪表/管道报表 ----
# 仪表（对齐 frmReportInstrument）：位号/归属/部件码/坐标，无 No-MTO 过滤
# （VB 原版不滤）；管道（对齐 frmReportPipingRun）：run 级过程数据/保温
# 一览。链路口径细节见 skill「仪表明细」「run 明细」。


@dataclass
class InstrumentRow:
    """仪表明细行（逐台一行）。"""
    oid: str
    tag: str             # 位号（JNamedItem 名）
    run_name: str
    line_name: str
    part_code: str       # 部件码（def IndustryCommodityCode）
    x: Optional[float] = None   # 坐标 mm
    y: Optional[float] = None
    z: Optional[float] = None
    unit_name: str = ""
    area_name: str = ""


@dataclass
class RunRow:
    """管道（run）属性行。"""
    oid: str
    name: str
    line_name: str
    spec: str
    op_t: Optional[float] = None      # ℃
    op_p: Optional[float] = None      # MPa
    de_t: Optional[float] = None      # ℃
    de_p: Optional[float] = None      # MPa
    is_insulated: Optional[bool] = None
    ins_purpose: str = ""
    ins_material: str = ""
    ins_temp: Optional[float] = None    # 保温温度（℃，库内开尔文换算）
    ins_thk: Optional[float] = None     # 保温厚度 mm
    unit_name: str = ""
    area_name: str = ""


def extract_instruments(db: Db, plant: dict, opts: MatOptions,
                        on_log: Callable[[str], None] = None
                        ) -> List[InstrumentRow]:
    """提取仪表明细（对齐 VB：无 No-MTO 过滤）。"""
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    mdb = plant["model"]
    oid_out = "RAWTOHEX(pp.oid)" if is_ora else "CONVERT(VARCHAR(36), pp.oid)"
    pipe_oid_out = ("RAWTOHEX(g.oidOrigin)" if is_ora
                    else "CONVERT(VARCHAR(36), g.oidOrigin)")
    scope_cte, scope_cond = _scope_sql(db, opts, "f.oidOrigin")
    if is_ora:
        code_expr = ("NVL(TO_NCHAR(A.IndustryCommodityCode), "
                     "TO_NCHAR(cn.RelationName))")
    else:
        code_expr = "ISNULL(A.IndustryCommodityCode, cn.RelationName)"
    # 类型判定：默认 JOIN MDB 类型表，查询失败回退 RDB 本体表
    type_join = (f"JOIN {_t(db, mdb, 'ROUTEPipeInstrumentOcc')} io "
                 f"ON io.Oid = pp.oid")

    def _build(typing: str) -> str:
        head = (type_join if typing == "table"
                else "JOIN JRteInstrument io ON io.Oid = pp.oid")
        return f"""{scope_cte}SELECT {oid_out} AS objOid,
       n.ItemName AS tag, rn.ItemName AS runName, h.ItemName AS lineName,
       {pipe_oid_out} AS pipeOid,
       {code_expr} AS partCode,
       MIN(pf.LocationX) AS locX, MIN(pf.LocationY) AS locY,
       MIN(pf.LocationZ) AS locZ
FROM JRtePipePart pp
{head}
JOIN JNamedItem n ON n.oid = pp.oid
JOIN {_t(db, mdb, 'CoRERelationOrigin')} d
   ON d.oid = pp.oid AND {_rel(db, _REL_PART2DEF, "d")}
LEFT JOIN {_t(db, mdb, 'REFDATPipeComponent')} A
   ON A.oid = d.oidTarget
LEFT JOIN {_t(db, mdb, 'CoRERelationOrigin')} cn
   ON cn.oidTarget = d.oidTarget AND {_rel(db, _REL_NAME_CARRIER, "cn")}
LEFT JOIN XPathGeneratedParts xp ON xp.OidDestination = pp.oid
LEFT JOIN JRtePathFeature pf ON pf.Oid = xp.OidOrigin
JOIN XOwnsParts f ON f.oidDestination = pp.oid
JOIN XSystemHierarchy g ON g.oidDestination = f.oidOrigin
JOIN JNamedItem h ON h.oid = g.oidOrigin
LEFT JOIN JNamedItem rn ON rn.oid = f.oidOrigin
WHERE 1=1{scope_cond}
GROUP BY {oid_out}, n.ItemName, rn.ItemName, h.ItemName,
         {pipe_oid_out}, {code_expr}"""

    rows = None
    err = None
    for typing in ("table", "prefix"):
        try:
            rows = db.query(_build(typing))
            if typing == "prefix":
                log("类型表查询失败，已回退 RDB 本体表（JRteInstrument）判定")
            break
        except Exception as ex:
            err = ex
    if rows is None:
        raise err
    anc = load_system_ancestors(db)

    def _f(v) -> Optional[float]:
        return float(v) if v is not None else None

    out = []
    for oid, tag, run, line, pipe_oid, code, x, y, z in rows:
        unit, area = anc.get(str(pipe_oid or ""), ("", ""))
        out.append(InstrumentRow(
            oid=str(oid or ""), tag=str(tag or ""),
            run_name=str(run or ""), line_name=str(line or ""),
            part_code=str(code or ""), x=_f(x), y=_f(y), z=_f(z),
            unit_name=unit, area_name=area))
    out.sort(key=lambda r: (r.line_name, r.tag))
    log(f"仪表明细：{len(out)} 台")
    return out


def extract_piping_runs(db: Db, plant: dict, opts: MatOptions,
                        on_log: Callable[[str], None] = None) -> List[RunRow]:
    """提取管道（run）属性一览：过程数据/保温/等级/归属。

    run 集 = XOwnsParts 属主且 JOIN JRtePipeRun 限定 PipeRun（裸
    XOwnsParts 会混入风管/桥架）；默认 JOIN MDB 类型表，查询失败
    回退 RDB 本体表（成员相同）。
    """
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    mdb = plant["model"]
    oid_out = ("RAWTOHEX(f.oidOrigin)" if is_ora
               else "CONVERT(VARCHAR(36), f.oidOrigin)")
    pipe_oid_out = ("RAWTOHEX(hh.oidOrigin)" if is_ora
                    else "CONVERT(VARCHAR(36), hh.oidOrigin)")
    scope_cte, scope_cond = _scope_sql(db, opts, "f.oidOrigin")
    if is_ora:
        # Oracle RDB 的 JProcessDataInfo/JRteInsulation 是 UNION-ALL 巨视图
        # （COREBaseClass ⋈ 属性表 ×N 分支），JOIN 进主 SQL 会被优化器整体
        # 物化（BPA 20.5K run 实测 819s）；改两段式——主 SQL 只取 run 骨架，
        # spec 走 weld 同款关系直查（完整 GUID 等值可走索引），两个属性
        # 视图全量直读成字典（实测 1.2s/10.3s），Python 端拼装（合计 ~14s）
        attr_cols = ""
        attr_join = ""
    else:
        attr_cols = (""", sd.SpecName AS specName,
       pd.OperatingMaxTemp, pd.OperatingMaxPressure,
       pd.DesignMaxTemp, pd.DesignMaxPressure,
       ri.IsInsulated, ri.InsulationPurpose, ri.InsulationMaterial,
       ri.InsulationTemperature, ri.InsulationThickness""")
        attr_join = """LEFT JOIN JProcessDataInfo pd ON pd.Oid = f.oidOrigin
LEFT JOIN JRteInsulation ri ON ri.Oid = f.oidOrigin
LEFT JOIN XPathRunUsesSpec xs ON xs.oidOrigin = f.oidOrigin
LEFT JOIN JDPipeSpec sd ON sd.oid = xs.oidDestination
"""

    def _build(typing: str) -> str:
        # 类型判定：表版 JOIN MDB 类型表，回退 JOIN RDB 本体表（成员相同）
        type_join = (f"JOIN {_t(db, mdb, 'JRtePipeRun')} pr "
                     f"ON pr.Oid = f.oidOrigin" if typing == "table"
                     else "JOIN JRtePipeRun pr ON pr.Oid = f.oidOrigin")
        return f"""{scope_cte}SELECT DISTINCT {oid_out} AS runOid,
       rn.ItemName AS runName, h.ItemName AS lineName,
       {pipe_oid_out} AS pipeOid{attr_cols}
FROM XOwnsParts f
JOIN JNamedItem rn ON rn.oid = f.oidOrigin
{type_join}
JOIN XSystemHierarchy hh ON hh.oidDestination = f.oidOrigin
JOIN JNamedItem h ON h.oid = hh.oidOrigin
{attr_join}WHERE 1=1{scope_cond}"""

    rows = None
    err = None
    for typing in ("table", "prefix"):
        try:
            rows = db.query(_build(typing))
            if typing == "prefix":
                log("类型表查询失败，已回退 RDB 本体表（JRtePipeRun）判定")
            break
        except Exception as ex:
            err = ex
    if rows is None:
        raise err

    spec_d: Dict[str, str] = {}
    pd_d: Dict[str, tuple] = {}
    ri_d: Dict[str, tuple] = {}
    if is_ora:
        spec_names = {str(n) for (n,) in
                      db.query("SELECT SpecName FROM JDPipeSpec")}
        pairs = db.query(
            f"SELECT r1.oid, r2.RelationName "
            f"FROM {mdb}.CoRERelationOrigin r1 "
            f"LEFT JOIN {mdb}.CoRERelationOrigin r2 "
            f"ON (r1.oidTarget = r2.oidTarget "
            f"AND r2.RelationType = HEXTORAW('{_GUID_NAME_CARRIER}')) "
            f"WHERE r1.RelationType = HEXTORAW('{_GUID_RUNUSESSPEC}')")
        spec_d = {str(k).lower(): str(v) for k, v in pairs
                  if v is not None and str(v) in spec_names}
        pd_d = {str(r[0]).lower(): tuple(r[1:]) for r in db.query(
            "SELECT Oid, OperatingMaxPressure, OperatingMaxTemp, "
            "DesignMaxPressure, DesignMaxTemp FROM JProcessDataInfo")}
        ri_d = {str(r[0]).lower(): tuple(r[1:]) for r in db.query(
            "SELECT Oid, IsInsulated, InsulationPurpose, InsulationMaterial, "
            "InsulationTemperature, InsulationThickness FROM JRteInsulation")}
        log(f"Oracle 两段式字典：spec {len(spec_d)} / 过程数据 {len(pd_d)}"
            f" / 保温 {len(ri_d)} 行")

    anc = load_system_ancestors(db)
    ins_p_texts, ins_m_texts = _insulation_texts(db)

    def _txt(texts: Dict[int, str], v) -> str:
        if v is None or int(v) <= 0:
            return ""
        i = int(v)
        return texts.get(i, str(i))

    def _f(v) -> Optional[float]:
        return float(v) if v is not None else None

    # 单位换算同材料报表：温度开尔文→摄氏度，压力帕斯卡→兆帕
    def _t_c(v) -> Optional[float]:
        return _f(v) - 273.15 if v is not None else None

    def _p_m(v) -> Optional[float]:
        return _f(v) / 1_000_000.0 if v is not None else None

    def _mk(oid, name, line, pipe_oid, spec, op_t, op_p, de_t, de_p,
            insul, ins_p, ins_m, ins_temp, ins_thk, anc, p_texts, m_texts):
        unit, area = anc.get(str(pipe_oid or ""), ("", ""))
        return RunRow(
            oid=str(oid or ""), name=str(name or ""),
            line_name=str(line or ""), spec=str(spec or ""),
            op_t=_t_c(op_t), op_p=_p_m(op_p),
            de_t=_t_c(de_t), de_p=_p_m(de_p),
            is_insulated=_b(insul),
            ins_purpose=_txt(p_texts, ins_p),
            ins_material=_txt(m_texts, ins_m),
            ins_temp=_t_c(ins_temp),
            ins_thk=(_f(ins_thk) * 1000.0) if ins_thk is not None else None,
            unit_name=unit, area_name=area)

    def _b(v):
        # MSSQL BIT → bool；Oracle NUMBER → 0/1 int；偶见 'True' 文本全兼容
        if v is None:
            return None
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)):
            return v != 0
        return str(v).strip().lower() == "true"

    out = []
    if is_ora:
        for (oid, name, line, pipe_oid) in rows:
            k = str(oid or "").lower()
            op_p, op_t, de_p, de_t = pd_d.get(k, (None,) * 4)
            insul, ins_p, ins_m, ins_temp, ins_thk = ri_d.get(k, (None,) * 5)
            out.append(_mk(oid, name, line, pipe_oid, spec_d.get(k, ""),
                           op_t, op_p, de_t, de_p,
                           insul, ins_p, ins_m, ins_temp, ins_thk, anc,
                           ins_p_texts, ins_m_texts))
    else:
        for (oid, name, line, pipe_oid, spec, op_t, op_p, de_t, de_p,
             insul, ins_p, ins_m, ins_temp, ins_thk) in rows:
            out.append(_mk(oid, name, line, pipe_oid, spec,
                           op_t, op_p, de_t, de_p,
                           insul, ins_p, ins_m, ins_temp, ins_thk, anc,
                           ins_p_texts, ins_m_texts))
    out.sort(key=lambda r: (r.line_name, r.name))
    log(f"管道属性行：{len(out)} 条")
    return out


@dataclass
class PidRow:
    """PID/应力号行（设备 + 管线逐条一行）。"""
    oid: str
    name: str            # 模型名称（JNamedItem）
    pid_no: str          # HQSH_PIDNumber
    stress_no: str       # HQSH_StressNumber（仅管线有）
    model_type: str      # "设备" / "管线"


def extract_pid_numbers(db: Db, plant: dict, opts: MatOptions,
                        on_log: Callable[[str], None] = None) -> List[PidRow]:
    """提取设备/管线的 PID 号与应力分析题号（对齐 VB frmReportPIDStressNo：
    JHQSHEquipment ∪ JHQSHPipeline，列 oid/模型名称/PID号/应力分析题号/模型类型）。

    范围过滤仅对管线生效：管线 100% 挂在 XSystemHierarchy（P231282
    10001/10001），勾选范围按 sysScope 过滤；设备不在系统层次树
    （实测 0 台挂接，归属关系缺失），勾选范围时设备仍全量输出并
    日志提示，避免静默丢失。
    """
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    oid_out = "RAWTOHEX(p.Oid)" if is_ora else "CONVERT(VARCHAR(36), p.Oid)"
    equip_oid_out = ("RAWTOHEX(e.Oid)" if is_ora
                     else "CONVERT(VARCHAR(36), e.Oid)")
    # 设备无应力号列，UNION 对齐补 NULL（Oracle 需显式 CAST 定型）
    null_char = ("CAST(NULL AS NVARCHAR2(100))" if is_ora
                 else "CAST(NULL AS NVARCHAR(100))")
    scope_cte, scope_cond = _scope_sql(db, opts, "p.Oid")

    # Oracle 站点 HQSH 定制部署参差（BPA：管线仅 StressNumber、设备表
    # 整个不存在；P181249/P201189 等三列齐全）：查 ALL_TAB_COLUMNS 探测，
    # 缺列补 NULL、缺表跳过该分支、全缺则空返回（MSSQL 站点恒齐全）
    pl_pid = pl_stress = eq_pid = True
    if is_ora:
        # 数据字典对象名恒为大写存储，字面量必须 UPPER（混合大小写在
        # BINARY 比较下不匹配，探测会误报未部署）
        owner = str(plant.get("reports") or "").upper()
        have = {(str(t).upper(), str(c).upper()) for t, c in db.query(
            "SELECT TABLE_NAME, COLUMN_NAME FROM ALL_TAB_COLUMNS "
            f"WHERE OWNER = '{owner}' "
            "AND TABLE_NAME IN ('JHQSHPIPELINE', 'JHQSHEQUIPMENT') "
            "AND COLUMN_NAME IN ('HQSH_PIDNUMBER', 'HQSH_STRESSNUMBER')")}
        pl_tbl = any(t == "JHQSHPIPELINE" for t, _c in have)
        eq_tbl = any(t == "JHQSHEQUIPMENT" for t, _c in have)
        pl_pid = ("JHQSHPIPELINE", "HQSH_PIDNUMBER") in have
        pl_stress = ("JHQSHPIPELINE", "HQSH_STRESSNUMBER") in have
        eq_pid = ("JHQSHEQUIPMENT", "HQSH_PIDNUMBER") in have
        if not ((pl_tbl and (pl_pid or pl_stress)) or (eq_tbl and eq_pid)):
            log("当前项目库未部署 HQSH PID/应力号定制列，无数据可统计")
            return []
        if not (pl_tbl and (pl_pid or pl_stress)):
            log("提示：JHQSHPipeline 缺 HQSH 定制列，管线分支跳过")
        if not (eq_tbl and eq_pid):
            log("提示：JHQSHEquipment 缺失或无 HQSH_PIDNumber，设备分支跳过")

    parts = []
    if pl_pid or pl_stress:
        parts.append(f"""{scope_cte}SELECT {oid_out} AS objOid, n.ItemName AS name,
       {('p.HQSH_PIDNumber' if pl_pid else null_char)} AS pidNo,
       {('p.HQSH_StressNumber' if pl_stress else null_char)} AS stressNo,
       N'管线' AS modelType
FROM JHQSHPipeline p
JOIN JNamedItem n ON n.oid = p.Oid
WHERE 1=1{scope_cond}""")
    if eq_pid:
        parts.append(f"""SELECT {equip_oid_out}, n.ItemName, e.HQSH_PIDNumber,
       {null_char}, N'设备'
FROM JHQSHEquipment e
JOIN JNamedItem n ON n.oid = e.Oid""")
    rows = db.query("\nUNION ALL\n".join(parts))
    out = [PidRow(oid=str(oid or ""), name=str(nm or ""),
                  pid_no=str(pid or "") if pid is not None else "",
                  stress_no=str(st or "") if st is not None else "",
                  model_type=str(mt or ""))
           for oid, nm, pid, st, mt in rows]
    out.sort(key=lambda r: (r.model_type, r.name))
    log(f"PID/应力号：管线 {sum(1 for r in out if r.model_type == '管线')}"
        f" + 设备 {sum(1 for r in out if r.model_type == '设备')}"
        f" = {len(out)} 条")
    if opts.sys_oids:
        log("提示：设备不在系统层次树，范围过滤仅对管线生效，设备全量输出")
    return out


@dataclass
class CorrelationRow:
    """PID 关联状态行（逐对象一行，对齐 VB frmReportCorrelate 列）。"""
    oid: str
    name: str            # 模型名称（无名字对象如管段特征取生成部件名）
    model_class: str     # 类型码双轨显示：树节点名>类名（_type_label）
    path: str            # 模型路径（沿 XSystemHierarchy 祖先串）
    basis: str           # 关联PID（CorrelationBasis 翻译：Correlate/Ignore）
    status: str          # 关联状态（CorrelationStatus 翻译）
    approval: str        # 关联批准（CorrelationApprovalStatus 翻译）
    comparison: str      # 属性比较（细节需 COM API 不可复刻；未关联填提示）


# PID 关联对象类别（原版 frmReportCorrelate 过滤器提示的 7 类）→ 本体
# 类视图（互斥，一对象一本体类），勾选类别 UNION 出 oid 白名单。设备表
# 为 HQSH 定制，缺表按 JSmartEquipment → JEquipment 降级。管口双表：
# JPipeNozzle=管道管口本体（无关联属性行→未关联PID），
# JEquipmentComponent=设备端口（有属性行）。
PID_KIND_LABELS = {
    "equip": "设备",
    "nozzle": "管口",
    "pipeline": "管线",
    "run": "管道 Run",
    "feature": "管段特征",
    "instrument": "仪表件",
    "specialty": "特殊件",
}
_PID_KIND_TABLES: Dict[str, Tuple[str, ...]] = {
    "equip": ("JHQSHEquipment", "JSmartEquipment", "JEquipment"),
    "nozzle": ("JPipeNozzle", "JEquipmentComponent"),
    "pipeline": ("JPipelineSystem",),
    "run": ("JRtePipeRun",),
    "feature": ("JRtePipeAlongLegPathFeat", "JRtePipeEndPathFeat",
                "JRtePipeTurnPathFeat"),
    "instrument": ("JRteInstrument",),      # 部件实例挂 run（XOwnsParts）
    "specialty": ("JRteSpecialtyOccur",),   # 同上（P231282 100% 挂树）
}


def _pid_kind_sets(db: Db, is_ora: bool, kinds: frozenset,
                   log: Callable[[str], None]) -> Dict[str, Set[str]]:
    """勾选类别 → {kind: 该类别本体表中的 oid 集合}。

    equip 表组为降级链，只取第一张可用的；其余类别表组内并集，表
    均不可用时该类别得空集（日志提示，对象将被白名单排除）。
    """
    oid_e = "RAWTOHEX(Oid)" if is_ora else "CONVERT(VARCHAR(36), Oid)"
    sets: Dict[str, Set[str]] = {}
    for kind in PID_KIND_LABELS:
        if kind not in kinds:
            continue
        got: Set[str] = set()
        for tbl in _PID_KIND_TABLES[kind]:
            try:
                got |= {str(r[0]) for r in db.query(f"SELECT {oid_e} FROM {tbl}")}
            except Exception:
                if kind == "equip":
                    continue  # 设备表降级链：HQSH 缺则 Smart、再通用
                log(f"类别 {PID_KIND_LABELS[kind]}：类表 {tbl} 不可用")
                continue
            if kind == "equip":
                break     # 降级链命中即止
        sets[kind] = got
    return sets


# --------------------------------------------- 对象类型↔类别/类名映射 ----
# 全库内闭环，零文件依赖（数据链/双后端表名异名/边表零 GUID 语义/退役
# 数据源等细节见 skill「对象类型映射层」）：Schema 树四表 + RDB ClassId
# 登记表（_rpt_table 分流）桥接，ClassId = oid 前 8 位 hex 的整数值。
_CLASS_DICTS: Dict[str, Dict[int, str]] = {}
_NODE_LABELS: Dict[str, Dict[int, str]] = {}
_CATEGORY_DICTS: Dict[str, Dict[str, frozenset]] = {}
_TREE_CACHE: Dict[str, tuple] = {}


def _rpt_table(db: Db) -> str:
    """ClassId↔CLSID 登记表名（双后端异名，见上方注释）。"""
    if getattr(db, "provider", "mssql") == "oracle":
        return "RPTFIRSTCLASSCLASSDEF"
    return "REPORTFirstClassClassDef"


def _load_class_names(db: Db, plant: dict,
                      log: Callable[[str], None]) -> Dict[int, str]:
    """对象类型码 -> 类显示名（Schema 直读 + RDB 桥接，按项目缓存）。

    名字优先级：Schema 显示名（"Pipe Instrument"）> Schema 技術名
    （CPRtePipeInstrumentOcc）> RDB Username；Schema 不可用时整表
    RDB 兜底，再不行空 dict（模型类列回退 oid 前缀）。
    """
    code = (plant or {}).get("code") or "?"
    if code in _CLASS_DICTS:
        return _CLASS_DICTS[code]
    schema = (plant or {}).get("schema") or ""
    d: Dict[int, str] = {}
    if schema:
        try:
            rows = db.query(
                f"SELECT r.class_id, r.Username, ci.UserName, ci.Name "
                f"FROM {_rpt_table(db)} r LEFT JOIN "
                f"{_t(db, schema, 'ClassInfoView')} ci ON ci.CLSID = r.BOClsid")
            d = {int(cid): str(un or nm or rn or "")
                 for cid, rn, un, nm in rows}
            log(f"对象类名表：{len(d)} 个类型码（Schema ClassInfoView）")
        except Exception as ex:
            log(f"Schema 类名表不可用：{ex}")
    if not d:
        try:
            d = {int(cid): str(n or "") for cid, n in db.query(
                f"SELECT class_id, Username FROM {_rpt_table(db)}")}
            log(f"对象类名表：{len(d)} 个类型码（RDB 兜底）")
        except Exception as ex:
            log(f"对象类名表不可用，模型类列回退 oid 前缀：{ex}")
    _CLASS_DICTS[code] = d
    return d


def _schema_tree(db: Db, plant: dict,
                 log: Callable[[str], None]) -> tuple:
    """Schema 类型树三元组（按项目缓存）：

    nodes: 节点oid -> (Name, UserName)；edges: 节点oid -> [孩子oid]；
    reps: 节点oid -> set(ClassId)（节点代表的类中已桥接到 ClassId 者）。
    树为官方全量（与过滤器是否用过无关）；任一查询失败返回已得部分。
    """
    code = (plant or {}).get("code") or "?"
    if code in _TREE_CACHE:
        return _TREE_CACHE[code]
    schema = (plant or {}).get("schema") or ""
    nodes: Dict[str, tuple] = {}
    edges: Dict[str, list] = {}
    reps: Dict[str, set] = {}
    if schema:
        is_ora = getattr(db, "provider", "mssql") == "oracle"
        oid_x = (lambda c: f"RAWTOHEX({c})") if is_ora \
            else (lambda c: f"CONVERT(VARCHAR(36), {c})")

        def gid(v) -> str:
            """oid 键归一：去横线大写 32 位 hex（MSSQL CONVERT 带
            横线、Oracle RAWTOHEX 无横线，统一同一键空间）。"""
            return str(v).replace("-", "").strip().upper()

        try:
            nodes = {gid(o): (str(nm or ""), str(un or "")) for o, nm, un in
                     db.query(f"SELECT {oid_x('n.oid')}, n.Name, n.UserName "
                              f"FROM {_t(db, schema, 'NodeInfoView')} n")}
            zero = "0" * 32
            for po, ch in db.query(
                    f"SELECT {oid_x('e.oidOrg')}, {oid_x('e.oidDst')} "
                    f"FROM {_t(db, schema, 'JNode_HasChild_JNodes')} e"):
                po, ch = gid(po), gid(ch)
                # (全零→子)＝根标记（子即顶级节点，不入边，让根检测
                # 「未作为孩子出现」自然命中）；(父→全零)＝死行
                if po == zero or ch == zero:
                    continue
                edges.setdefault(po, []).append(ch)
            for no, cid in db.query(
                    f"SELECT {oid_x('r.oidOrg')}, f.class_id "
                    f"FROM {_t(db, schema, 'JNode_Represents_JClasses')} r "
                    f"JOIN {_t(db, schema, 'ClassInfoView')} ci "
                    f"ON ci.oid = r.oidDst "
                    f"LEFT JOIN {_rpt_table(db)} f "
                    f"ON f.BOClsid = ci.CLSID"):
                if cid is not None:
                    reps.setdefault(gid(no), set()).add(int(cid))
            log(f"Schema 类型树：{len(nodes)} 节点 / {len(edges)} 父节点 / "
                f"{sum(len(v) for v in reps.values())} 条节点-类映射")
        except Exception as ex:
            log(f"Schema 类型树不可用：{ex}")
    tree = (nodes, edges, reps)
    _TREE_CACHE[code] = tree
    return tree


def _load_node_labels(db: Db, plant: dict,
                      log: Callable[[str], None]) -> Dict[int, str]:
    """对象类型码 -> 最深类别节点显示名（按项目缓存）。

    同一类可挂多个节点（CPMPipeOccur 同时挂 Pipes 叶与 Piping 根），
    取树最深者 = 最具体类别（"Pipes" 而非 "Piping"）；防脏数据环。
    """
    code = (plant or {}).get("code") or "?"
    if code in _NODE_LABELS:
        return _NODE_LABELS[code]
    nodes, edges, reps = _schema_tree(db, plant, log)
    depth: Dict[str, int] = {}
    parented = {ch for kids in edges.values() for ch in kids}
    stack = [(o, 0) for o in nodes if o not in parented]
    while stack:
        cur, dp = stack.pop()
        if depth.get(cur, -1) >= dp:
            continue
        depth[cur] = dp
        for ch in edges.get(cur, ()):
            if ch in nodes:
                stack.append((ch, dp + 1))
    best: Dict[int, tuple] = {}          # class_id -> (depth, label)
    for oid, cids in reps.items():
        nm, un = nodes.get(oid, ("", ""))
        label = un or nm
        if not label:
            continue
        for cid in cids:
            if cid not in best or best[cid][0] < depth.get(oid, 0):
                best[cid] = (depth.get(oid, 0), label)
    d = {cid: lbl for cid, (_, lbl) in best.items()}
    _NODE_LABELS[code] = d
    return d


def load_type_categories(db: Db, plant: dict,
                         log: Callable[[str], None]) -> Dict[str, frozenset]:
    """类型树类别(节点路径, 形如 Piping\\PipingParts\\Pipes) -> ClassId
    集合（按项目缓存，Schema 树直读）。

    路径段取节点 Name（技術名，与 MDB 过滤器 strGroupPath 同格式）；
    类别 ClassId 集 = 自身 + 全部后代节点所代表类的并集（"Piping" 根
    即管道域全部 ClassId，含 80074 PipingGeneric）。同一节点挂多父时
    生成多条路径（与 SP3D 树的 DAG 形态一致）；防脏数据环。
    """
    code = (plant or {}).get("code") or "?"
    if code in _CATEGORY_DICTS:
        return _CATEGORY_DICTS[code]
    nodes, edges, reps = _schema_tree(db, plant, log)
    d: Dict[str, frozenset] = {}

    def walk(oid: str, path: str, anc: frozenset) -> set:
        if oid not in nodes or oid in anc:
            return set()
        p = f"{path}\\{nodes[oid][0]}" if path else nodes[oid][0]
        acc = set(reps.get(oid, ()))
        for ch in edges.get(oid, ()):
            acc |= walk(ch, p, anc | {oid})
        d[p] = d.get(p, frozenset()) | frozenset(acc)
        return acc

    parented = {ch for kids in edges.values() for ch in kids}
    for oid in nodes:
        if oid not in parented:
            walk(oid, "", frozenset())
    n_hit = sum(1 for v in d.values() if v)
    log(f"类型树类别：{len(d)} 条路径，其中 {n_hit} 条映射到 ClassId")
    _CATEGORY_DICTS[code] = d
    return d


def _cls_label(class_names: Dict[int, str], oid: str) -> str:
    """模型类显示：已登记直接显示类名；未登记类型码回退裸前缀。"""
    try:
        return class_names.get(int(oid[:8], 16), "") or oid[:8]
    except ValueError:
        return oid[:8]


def _type_label(node_labels: Dict[int, str], class_names: Dict[int, str],
                oid: str) -> str:
    """模型类显示（双轨）：最深类别节点名（"Piping Instruments"，对齐
    过滤器类型树）> 类显示名（"Pipe Instrument"）> 裸 oid 前缀。"""
    try:
        cid = int(oid[:8], 16)
    except ValueError:
        cid = None
    return (node_labels.get(cid, "") if cid is not None else "") \
        or _cls_label(class_names, oid)


def extract_pid_correlations(db: Db, plant: dict, opts: MatOptions,
                             on_log: Callable[[str], None] = None
                             ) -> List[CorrelationRow]:
    """提取 PID 关联状态总表（对齐 VB frmReportCorrelate 基础版）。

    挂接对象/类别正选并入/树锚点爬链范围过滤等口径细节见 skill
    「PID 关联」；属性比较列依赖 COM API 不可复刻，Not Correlated
    填「未关联PID」。
    """
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    oid_out = "RAWTOHEX(c.Oid)" if is_ora else "CONVERT(VARCHAR(36), c.Oid)"
    feat_out = ("RAWTOHEX(gp.oidOrigin)" if is_ora
                else "CONVERT(VARCHAR(36), gp.oidOrigin)")
    run_out = ("RAWTOHEX(x.oidOrigin)" if is_ora
               else "CONVERT(VARCHAR(36), x.oidOrigin)")
    part_out = ("RAWTOHEX(x.oidDestination)" if is_ora
                else "CONVERT(VARCHAR(36), x.oidDestination)")
    gpart_out = ("RAWTOHEX(gp.oidDestination)" if is_ora
                 else "CONVERT(VARCHAR(36), gp.oidDestination)")
    po_out = ("RAWTOHEX(x.oidOrigin)" if is_ora
              else "CONVERT(VARCHAR(36), x.oidOrigin)")
    ch_out = ("RAWTOHEX(x.oidDestination)" if is_ora
              else "CONVERT(VARCHAR(36), x.oidDestination)")

    # 1) 关联主表 + 对象名
    rows = db.query(f"""SELECT {oid_out} AS objOid, n.ItemName AS name,
       c.CorrelationBasis, c.CorrelationStatus, c.CorrelationApprovalStatus
FROM JEFWCorrelation c
LEFT JOIN JNamedItem n ON n.oid = c.Oid""")

    # 2) 代码翻译（_CL 视图带 Short/Long 两列；查失败用内置默认）
    def _load_codes(col: str, default: Dict) -> Dict:
        try:
            return {str(k): str(v) for k, v in db.query(
                f"SELECT DISTINCT {col}, {col}_ShortValue "
                f"FROM JEFWCorrelation_CL") if k is not None}
        except Exception:
            return dict(default)

    basis_txt = _load_codes("CorrelationBasis",
                            {"0": "Correlate", "1": "Ignore"})
    status_txt = _load_codes("CorrelationStatus",
                             {"0": "Not Correlated", "4": "Ignored"})
    appr_txt = _load_codes("CorrelationApprovalStatus", {})

    # 3) 管段特征 → 生成部件（边表用 LEFT JOIN 取名：普通部件不在
    #    JNamedItem，INNER JOIN 会让边表残缺、锚点失效）
    feat_part: Dict[str, List[str]] = {}     # 特征 → 生成部件 oid
    feat_names: Dict[str, List[str]] = {}    # 特征 → 有名部件名（VB 语义）
    for f, p, pn in db.query(f"""SELECT {feat_out}, {gpart_out}, pn.ItemName
FROM XPathGeneratedParts gp
JOIN JEFWCorrelation c ON c.Oid = gp.oidOrigin
LEFT JOIN JNamedItem pn ON pn.oid = gp.oidDestination"""):
        feat_part.setdefault(str(f), []).append(str(p))
        if pn:
            feat_names.setdefault(str(f), []).append(str(pn))

    # 4) 部件 → run（XOwnsParts: origin=run, destination=part）
    part_run: Dict[str, str] = {}
    for r, p in db.query(f"""SELECT {run_out}, {part_out} FROM XOwnsParts x"""):
        part_run.setdefault(str(p), str(r))

    # 4.5) 端口/管口属主（锚点爬链用）：设备端口经 XHasEqpComponents
    #      挂设备本体；管口经 XDistribPorts 挂属主（设备/管件/设备端口）
    #      ——设备本体不挂 run，管口属主管件再经 XOwnsParts 挂 run
    owner_of: Dict[str, str] = {}
    for p, o in db.query(f"""SELECT {ch_out}, {po_out}
FROM XHasEqpComponents x
JOIN JEquipmentComponent ec ON ec.Oid = x.oidDestination"""):
        owner_of.setdefault(str(p), str(o))
    for p, o in db.query(f"""SELECT {ch_out}, {po_out}
FROM XDistribPorts x
JOIN JPipeNozzle nz ON nz.Oid = x.oidDestination"""):
        owner_of.setdefault(str(p), str(o))

    # 5) 系统树边 + 两侧节点名（parent/children/名称）。设备本体→设备
    #    系统的最后一跳在 XHasEqpAsChild（XSystemHierarchy 视图不含设
    #    备级边，VB ISystemChild 两种关系都走），一并并入树图
    parent: Dict[str, str] = {}
    children: Dict[str, List[str]] = {}
    names: Dict[str, str] = {}

    def _absorb(tbl: str) -> None:
        for po, ch, pn, cn in db.query(f"""SELECT {po_out}, {ch_out},
           np.ItemName, nc.ItemName
FROM {tbl} x
LEFT JOIN JNamedItem np ON np.oid = x.oidOrigin
LEFT JOIN JNamedItem nc ON nc.oid = x.oidDestination"""):
            parent.setdefault(str(ch), str(po))
            children.setdefault(str(po), []).append(str(ch))
            if pn:
                names.setdefault(str(po), str(pn))
            if cn:
                names.setdefault(str(ch), str(cn))

    _absorb("XSystemHierarchy")
    _absorb("XHasEqpAsChild")

    # 6) 对象类别白名单（正选勾选；kind_sets 为 None=未启用过滤）
    kind_sets: Optional[Dict[str, Set[str]]] = None
    allowed: Optional[Set[str]] = None
    if opts.pid_kinds:
        kind_sets = _pid_kind_sets(db, is_ora, opts.pid_kinds, log)
        allowed = set()
        for s in kind_sets.values():
            allowed |= s

    # 6.5) 勾选过滤时并入「支持关联但无属性行」的对象（管口本体/
    #      仪表件/特殊件）：以部件/本体身份挂树，原版按类级接口枚举、
    #      状态默认 Not Correlated→未关联PID
    rows = list(rows)
    n_extra: Dict[str, int] = {}
    if allowed is not None:
        ext_oid = ("RAWTOHEX(t.Oid)" if is_ora
                   else "CONVERT(VARCHAR(36), t.Oid)")
        for kind, tbl in (("nozzle", "JPipeNozzle"),
                          ("instrument", "JRteInstrument"),
                          ("specialty", "JRteSpecialtyOccur")):
            if kind not in (opts.pid_kinds or ()):
                continue
            try:
                ext = db.query(f"""SELECT {ext_oid}, n.ItemName
FROM {tbl} t
LEFT JOIN JEFWCorrelation c ON c.Oid = t.Oid
LEFT JOIN JNamedItem n ON n.oid = t.Oid
WHERE c.Oid IS NULL""")
            except Exception:
                log(f"类别 {PID_KIND_LABELS[kind]}：并入表 {tbl} 不可用")
                continue
            n_extra[kind] = len(ext)
            rows.extend((r[0], r[1], None, None, None) for r in ext)

    # 7) 勾选范围 → 树上子孙集合
    scope_oids: Set[str] = set()
    if opts.sys_oids:
        stack = [o for o in (str(x) for x in opts.sys_oids)]
        scope_oids.update(stack)
        while stack:
            cur = stack.pop()
            for c in children.get(cur, ()):
                if c not in scope_oids:
                    scope_oids.add(c)
                    stack.append(c)

    # 7) 树锚点：沿属主边爬到树上节点（自身在树 / 部件或管口→run/属主
    #    / 特征经 生成部件→run）与路径（爬到根）。仪表件、特殊件、管件
    #    经 XOwnsParts 挂 run；设备端口经 XHasEqpComponents、管口经
    #    XDistribPorts 挂属主（设备/管件）
    def _anchor(oid: str) -> Optional[str]:
        cur, seen = oid, set()
        while cur and cur not in seen:
            seen.add(cur)
            if cur in parent:
                return cur
            cur = part_run.get(cur) or owner_of.get(cur)
        for p in feat_part.get(oid, ()):
            r = part_run.get(p)
            if r:
                return r
        return None

    def _path(oid: str) -> str:
        segs = []
        cur, seen = oid, set()
        while cur and cur not in seen:
            seen.add(cur)
            nm = names.get(cur)
            if nm:
                segs.append(nm)
            cur = parent.get(cur)
        return "\\".join(reversed(segs))

    anchor_memo: Dict[str, Optional[str]] = {}
    out: List[CorrelationRow] = []
    n_equip_kept = n_no_anchor = n_kind_drop = 0
    # 设备判定集合：勾选过滤时取白名单内设备集合，未启用时也按设备
    # 本体表降级链取（链全缺为空集，无锚设备不再保留）
    equip_set = (kind_sets.get("equip") if kind_sets is not None else
                 _pid_kind_sets(db, is_ora, frozenset(("equip",)),
                                log).get("equip"))
    # 模型类翻译：类型码(=oid 首段 int) -> 节点名/类名双轨（替代裸前缀）
    class_names = _load_class_names(db, plant, log)
    node_labels = _load_node_labels(db, plant, log)
    for oid, nm, b, s, a in rows:
        oid = str(oid or "")
        if allowed is not None and oid not in allowed:
            n_kind_drop += 1  # 非所选类别（白名单外，含无链路杂项）
            continue
        if oid not in anchor_memo:
            anchor_memo[oid] = _anchor(oid)
        anchor = anchor_memo[oid]
        # 设备判定：设备本体表降级链成员（能走到这里的必属勾选类别）
        is_equip = oid in (equip_set or ())
        if scope_oids:
            if anchor and anchor in scope_oids:
                pass  # 范围内
            elif not anchor and is_equip:
                n_equip_kept += 1  # 无挂树边设备兜底保留
            else:
                n_no_anchor += 1  # 无系统归属链或范围外：排除
                continue
        elif not anchor and not is_equip:
            # 游离对象排除（原版沿模型结构枚举取不到，日志计数）
            n_no_anchor += 1
            continue
        # 名称链：对象名 → 生成部件名（VB：p1;p2;...）→ 所属 run 名
        #（普通部件/特征无 JNamedItem 行，API 自动名 SQL 拿不到）
        name = str(nm or "")
        if not name:
            name = ";".join(feat_names.get(oid, ()))
        if not name and anchor:
            name = names.get(anchor, "")
        # 无属性行（并入的管口本体/仪表/特殊件）：原版类级接口取
        # 属性得默认值，状态 Not Correlated→未关联PID
        b_s = basis_txt.get(str(b), "") if b is not None else ""
        s_s = (status_txt.get(str(s), "") if s is not None
               else "Not Correlated")
        a_s = appr_txt.get(str(a), "") if a is not None else ""
        out.append(CorrelationRow(
            oid=oid, name=name, model_class=_type_label(node_labels, class_names, oid),
            path=_path(anchor) if anchor else name,
            basis=b_s, status=s_s, approval=a_s,
            comparison="未关联PID" if (s is None or s == 0) else ""))
    n_nc = sum(1 for r in out if r.comparison)
    log(f"PID 关联：{len(out)} 行（未关联 {n_nc}）")
    if allowed is not None:
        chosen = "、".join(lbl for k, lbl in PID_KIND_LABELS.items()
                           if k in (opts.pid_kinds or ()))
        log(f"类别过滤：{chosen}（排除 {n_kind_drop} 个非所选类别对象）")
    if n_extra:
        det = "、".join(f"{PID_KIND_LABELS[k]} {n}" for k, n in n_extra.items())
        log(f"并入无属性行对象（按未关联PID 计）：{det}")
    if opts.sys_oids:
        extra = (f"，{n_equip_kept} 台无挂树边设备兜底保留"
                 if n_equip_kept else "")
        log(f"范围过滤：{n_no_anchor} 个无归属链对象被排除{extra}")
    elif n_no_anchor:
        log(f"整项目范围：{n_no_anchor} 个无归属链游离对象被排除"
            f"（原版过滤器取不到此类对象）")
    out.sort(key=lambda r: (r.path, r.name, r.oid))
    return out


@dataclass
class DwgRow:
    """ISO 图纸行（每图一行，带终版版次）。"""
    oid: str
    name: str
    path: str             # DrawingFolderHierarchy 路径
    rev_minor: str        # 终版版次
    rev_desc: str         # 终版描述
    rev_by: str           # 终版设计人
    chk_by: str           # 终版校核人
    app_by: str           # 终版审核人
    rev_date: str         # 终版日期（库存字符串如 2026.01.13）
    pages: str = ""       # 页数（ISO 输出 XML 的 TotalSheet）
    doc_ade3: str = ""    # 内部：上游文档对象 oid（页数提取键，不导出）


# sheet ← 文档对象（DrawingDocument）无名关系；文档对象 -xml-> 图纸内容
_TOTAL_SHEET_PAT = None  # 惰性编译（re 仅此处用）

# 图纸目录树（DrawingFolderHierarchy）关系 GUID；运行时优先查 RDB 函数
# dbo.REPORTGetHierarchyRelations（REPORTGetHierarchyPathByChildOid 同源），
# 查不到（如 Oracle）回退此默认值（P231282 实测）
_DWG_FOLDER_REL = "8E645EFF-904C-4896-9D75-20047C9F2379"


def load_drawing_tree(db: Db, plant: dict) -> Tuple[list, list]:
    """加载图纸目录树（DrawingFolderHierarchy，Deliverable 文件夹层级）。

    层次存 MDB.CoreRelations（注意与其他层次相反：child=idorigin、
    parent=iddestination，路径函数 REPORTGetHierarchyPathByChildOid 的
    DrawingFolderHierarchy 分支即此实现）。SP3D FilterSelector（模型
    过滤器）选不到该树，故 ISO 图纸报表单独提供目录勾选。
    返回 (roots, edges)：roots=嵌套节点 [{"name","oid","children"}]，
    edges=[(child_oid, parent_oid)]（extract 端展开子树用）。
    """
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    mdb = plant["model"]
    guids = [_DWG_FOLDER_REL]
    if not is_ora:
        try:
            rows = db.query("SELECT RelationGUID FROM "
                            "dbo.REPORTGetHierarchyRelations("
                            "'DrawingFolderHierarchy')")
            if rows:
                guids = [str(r[0]) for r in rows]
        except Exception:
            pass
    gl = ", ".join(_gl(db, g) for g in guids)
    child_out = ("RAWTOHEX(c.idorigin)" if is_ora
                 else "CONVERT(VARCHAR(36), c.idorigin)")
    parent_out = ("RAWTOHEX(c.iddestination)" if is_ora
                  else "CONVERT(VARCHAR(36), c.iddestination)")
    rows = db.query(f"""SELECT {child_out}, {parent_out}, n.ItemName
FROM {_t(db, mdb, 'CoreRelations')} c
LEFT JOIN JNamedItem n ON n.oid = c.idorigin
WHERE c.RelationType IN ({gl})""")
    edges = [(str(a), str(b)) for a, b, _n in rows]
    names = {str(a): str(n or "") for a, _b, n in rows}
    # parent 名称也要（根节点只出现在 parent 侧）
    for p_oid in {b for _a, b in edges}:
        if p_oid not in names:
            names[p_oid] = ""
    if is_ora:
        extra = db.query(f"""SELECT RAWTOHEX(n.oid), n.ItemName
FROM {_t(db, mdb, 'JNamedItem')} n
WHERE n.oid IN ({', '.join(_gl(db, o) for o in names if not names[o])})""")
    else:
        extra = db.query(f"""SELECT CONVERT(VARCHAR(36), n.oid), n.ItemName
FROM JNamedItem n
WHERE n.oid IN ({', '.join(_gl(db, o) for o in names if not names[o])})""")
    for oid, nm in extra:
        names[str(oid)] = str(nm or "")
    children: Dict[str, List[str]] = {}
    for c, p in edges:
        children.setdefault(p, []).append(c)
    child_set = {c for c, _p in edges}
    # 根 = 从未当过 child 的节点（一般就是项目根）
    roots_oid = [p for p in children if p not in child_set]

    def _node(oid: str) -> dict:
        return {"name": names.get(oid, ""), "oid": oid,
                "children": [_node(c) for c in children.get(oid, ())]}

    roots = [_node(o) for o in roots_oid]
    return roots, edges


def _dwg_total_sheets(db: Db, plant: dict, ade3_oids: List[str],
                      log: Callable[[str], None]) -> Dict[str, str]:
    """批量提取图纸页数：文档对象 -xml-> MDB.DRAWNGDocumentData.DataBlob
    （ZIP，Deflate64 压缩），解压得 ISO 输出 XML，取 HEADERTotalSheet 节点。

    P231282 实测 32710 个 blob 共 264MB：每批 2000 个约 0.5s（拉+解压），
    全量 <10s；1995/2000 命中（少量图纸 XML 无该节点，如实留空）。
    blob 经 db 层转 hex 字符串，bytes.fromhex 还原。缺
    zipfile-deflate64 包时整列留空（日志提示，不影响主报表）。
    """
    if not ade3_oids:
        return {}
    try:
        import io
        import re
        import zipfile
        import zipfile_deflate64  # noqa: F401（monkey-patch 支持 Deflate64）
    except ImportError:
        log("提示：缺少 zipfile-deflate64 包，页数列留空"
            "（pip install zipfile-deflate64）")
        return {}
    global _TOTAL_SHEET_PAT
    if _TOTAL_SHEET_PAT is None:
        _TOTAL_SHEET_PAT = re.compile(r'name="TotalSheet" data="(\d+)"')
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    mdb = plant["model"]
    ade3_out = "RAWTOHEX(c.oid)" if is_ora else "CONVERT(VARCHAR(36), c.oid)"
    # Oracle 直取 BLOB（fetch_lobs=False 返回 bytes，db 层再转 hex 文本）；
    # CONVERT(VARBINARY(MAX)) 是 T-SQL 语法，Oracle 报 ORA-00904
    blob_expr = "dd.DataBlob" if is_ora else "CONVERT(VARBINARY(MAX), dd.DataBlob)"
    # Oracle IN 列表上限 1000，保守取 500；MSSQL 用 2000
    step = 500 if is_ora else 2000
    uniq = list(dict.fromkeys(ade3_oids))
    out: Dict[str, str] = {}
    for i in range(0, len(uniq), step):
        chunk = uniq[i: i + step]
        sql = f"""SELECT {ade3_out},
       {blob_expr}
FROM {_t(db, mdb, 'CORERelationOrigin')} c
JOIN {_t(db, mdb, 'DRAWNGDocumentData')} dd ON dd.oid = c.oidTarget
WHERE c.RelationName = 'xml' AND c.oid IN ({_oid_list(db, chunk)})"""
        for ade3, blob in db.query(sql):
            data = (bytes.fromhex(blob) if isinstance(blob, str)
                    else bytes(blob))
            try:
                zf = zipfile.ZipFile(io.BytesIO(data))
                txt = zf.read(zf.namelist()[0]).decode(
                    "utf-8", errors="replace")
                m = _TOTAL_SHEET_PAT.search(txt)
                if m:
                    out[str(ade3 or "")] = m.group(1)
            except Exception:
                pass  # 个别坏 blob/解码失败：页数留空，不中断
    return out


def extract_iso_drawings(db: Db, plant: dict, opts: MatOptions,
                         on_log: Callable[[str], None] = None) -> List[DwgRow]:
    """提取 ISO 图纸目录与终版版次（对齐 VB frmReportISODrawing 意图：
    只列 ISO 图）。判定规则/版次链/路径库函数/页数与 RevDesc 解码
    细节见 skill「ISO 图纸」；图纸不在系统层次树，范围过滤恒输出
    全部（勾选时日志提示）。
    """
    log = on_log or (lambda m: None)
    is_ora = getattr(db, "provider", "mssql") == "oracle"
    mdb = plant["model"]
    if is_ora:
        oid_out = "RAWTOHEX(d.Oid)"
        ade3_out = "RAWTOHEX(dc.oid)"
        snap_out = "RAWTOHEX(s.OidDestination)"
        path_expr = ("RPTHierarchyPathByChildOid(s.OidDestination, "
                     "'DrawingFolderHierarchy')")
        desc_expr = "r.RevDesc"
        order_tail = ("r.RevVersion DESC NULLS LAST, "
                      "r.RevMinorNumber DESC NULLS LAST")
        # RevDesc 是 BLOB，DISTINCT 含 LOB 列报 ORA-00932；ROW_NUMBER 每
        # d.Oid 分区恰出一行 rn=1，最终行集不受去重影响，Oracle 免 DISTINCT
        sel_head = "SELECT"
    else:
        oid_out = "CONVERT(VARCHAR(36), d.Oid)"
        ade3_out = "CONVERT(VARCHAR(36), dc.oid)"
        snap_out = "CONVERT(VARCHAR(36), s.OidDestination)"
        path_expr = ("dbo.REPORTGetHierarchyPathByChildOid(s.OidDestination, "
                     "'DrawingFolderHierarchy')")
        desc_expr = "CAST(r.RevDesc AS NVARCHAR(200))"
        # MSSQL 的 DESC 默认 NULL 在最后，无需 NULLS LAST
        order_tail = "r.RevVersion DESC, r.RevMinorNumber DESC"
        sel_head = "SELECT DISTINCT"

    sql = f"""SELECT * FROM (
{sel_head}
{oid_out} AS dwgOid, ni.ItemName AS Document,
{path_expr} AS Path,
r.RevMinorNumber AS RevMark, r.RevDate,
{desc_expr} AS RevDesc, r.RevRevBy, r.RevChkBy, r.RevAppBy,
{ade3_out} AS docOid,
{snap_out} AS snapOid,
ROW_NUMBER() OVER (PARTITION BY d.Oid ORDER BY {order_tail}) AS rn
FROM JDDwgSheet d
JOIN JNamedItem ni ON ni.Oid = d.Oid
LEFT JOIN XSnapInHasSheets s ON s.OidOrigin = d.Oid
LEFT JOIN XSheetHasProperty sp ON sp.oidDestination = d.Oid
LEFT JOIN XPropertyHasChildren pc ON pc.oidDestination = sp.oidOrigin
LEFT JOIN XPropertiesHasRevisions xr ON xr.oidDestination = pc.oidOrigin
LEFT JOIN JDDwgRevision r ON r.Oid = xr.oidOrigin AND r.RevRecordType = 1
JOIN {_t(db, mdb, 'CORERelationOrigin')} dc
     ON dc.oidTarget = d.Oid AND {_rel(db, _REL_SHEET2DOC, "dc")}
JOIN {_t(db, mdb, 'CORERelationOrigin')} xc
     ON xc.oid = dc.oid AND xc.RelationName = 'xml'
) t WHERE rn = 1
ORDER BY t.Path, t.Document"""
    rows = db.query(sql)

    def _desc(v) -> str:
        if v is None:
            return ""
        if isinstance(v, (bytes, bytearray)):
            return v.decode("utf-16-le", errors="replace").strip("\x00 ")
        if is_ora and isinstance(v, str):
            # Oracle BLOB 经 db 转换层（_conv_oracle bytes→hex）成 hex 文本，
            # 先还原字节再按 UTF-16LE 解码；非 hex（理论不出现）原样返回
            try:
                return (bytes.fromhex(v).decode("utf-16-le", errors="replace")
                        .strip("\x00 "))
            except ValueError:
                return v
        return str(v)

    def _s(v) -> str:
        return "" if v is None else str(v)

    # 图纸目录树过滤：勾选节点沿 CoreRelations 展开子树，sheet 的
    # SnapIn 落在子树内才保留（目录树叶子就是 SnapIn/模板节点）
    if opts.dwg_folder_oids:
        _roots, edges = load_drawing_tree(db, plant)
        kids: Dict[str, List[str]] = {}
        for c, p in edges:
            kids.setdefault(p, []).append(c)
        scope: Set[str] = set()
        stack = [str(o) for o in opts.dwg_folder_oids]
        scope.update(stack)
        while stack:
            cur = stack.pop()
            for c in kids.get(cur, ()):
                if c not in scope:
                    scope.add(c)
                    stack.append(c)
        rows = [r for r in rows if str(r[10] or "") in scope]
        log(f"目录树过滤：{len(opts.dwg_folder_oids)} 个勾选节点，"
            f"子树 {len(scope)} 节点")

    out: List[DwgRow] = []
    ade3_list: List[str] = []
    for (oid, name, path, minor, date, desc, rby, cby, aby, ade3, snap,
         _rn) in rows:
        if minor is None and date is None and rby is None:
            # 无任何正式版次：对齐 VB 填 "Null"
            minor = date = rby = cby = aby = "Null"
        if ade3:
            ade3_list.append(str(ade3))
        out.append(DwgRow(
            oid=str(oid or ""), name=_s(name), path=_s(path),
            rev_minor=_s(minor), rev_desc=_desc(desc),
            rev_by=_s(rby), chk_by=_s(cby), app_by=_s(aby),
            rev_date=_s(date), doc_ade3=str(ade3 or "")))
    # 页数：批量解压图纸内容 XML 取 TotalSheet
    pages = _dwg_total_sheets(db, plant, ade3_list, log)
    for r in out:
        r.pages = pages.get(r.doc_ade3, "")
    n_rev = sum(1 for r in out if r.rev_minor != "Null")
    n_pg = sum(1 for r in out if r.pages)
    log(f"ISO 图纸：{len(out)} 张（有正式版次 {n_rev}，有页数 {n_pg}；"
        f"无出图内容的平面图/PID 壳图已排除）")
    if opts.sys_oids:
        log("提示：图纸不在系统层次树，忽略范围过滤，输出全部图纸")
    out.sort(key=lambda r: (r.path, r.name))
    return out


# ----------------------------------------------------------- 用户过滤器 ----
# SP3D 用户过滤器存于 MDB 的 CMNAPPCP*Filter 表族，组合关系在
# CORERelationOrigin（关系 GUID 语义与存储细节见 skill「用户过滤器」）。
_FILTER_COMPOUND = "38461280-82EC-11D4-9D3E-00104B34EBF6"   # 过滤器/组合节点下行
_FILTER_HIER = "1E5DA248-2999-11D5-8C51-0010B59403C7"       # 层次条件→引用节点
_FILTER_HIER_PLANES = "1E5DA245-2999-11D5-8C51-0010B59403C7"  # Planes 体积条件→平面集对象
_FILTER_FOLDER = "3846127C-82EC-11D4-9D3E-00104B34EBF6"     # 文件夹→过滤器
_FILTER_FOLDERSUB = "38461284-82EC-11D4-9D3E-00104B34EBF6"  # 文件夹→子文件夹
_ADAPTER_SYSTEM = "GSCADHierarchyHelper.SystemAdapter"       # 系统层次条件标识
_ADAPTER_SPACE = "GSCADHierarchyHelper.SpaceAdapter"         # 体积(Volume)条件标识


def load_user_filters(db: Db, win_user: Optional[str] = None):
    """从 MDB 解析过滤器树，返回 (tree, my_tree) 两棵根节点列表。

    tree = 「Plant Filters」（Category 1 公共区）；my_tree = 当前
    Windows 用户的「My Filters」（Category 0 且创建者本人），已剥掉
    单一根文件夹壳；win_user 为 None 或查不到时 my_tree 为空。
    节点 {"type": "folder"/"filter", "name", "children"/"kind",
    "oids", "sys_names", "vol_oids", "vol_names"}；kind：
    "sys"=含系统层次条件（oids 为引用的系统节点，可用于统计）；
    "vol"=仅含体积条件（vol_oids 为引用的体积对象，材料表按体积盒
    相交统计）；"sys"+vol_oids 非空=两者混合（统计时取交集）；
    "obj"/"none" 无 oids，界面置灰不可选；"sql"=SQL 过滤器
    （CMNAPPCPSqlFilter，不在 UserFilter 登记表、经文件夹关系直接挂
    树；sql_text/sql_result_col 供选择时在 RDB 执行，返回对象 oid
    直接作为统计范围）。
    体积条件 = CMNAPPCPNamedSpFilter(SpaceAdapter)，经 _HIER 关系引用
    体积对象（JNamedItem 有名；盒子在 MDB CORESpatialIndex）。Named
    Space 条件（CMNAPPCPHierarchyFilter+SpaceAdapter）引用的是体积本
    身而非其中内容，不支持，仅 Volume 条件按体积处理。
    """
    hier_oids = {str(r[0]) for r in db.query(
        "SELECT oid FROM CMNAPPCPHierarchyFilter "
        f"WHERE bstrAdapterProgID = '{_ADAPTER_SYSTEM}'")}
    # 体积(Volume)条件两种录入方式（存储结构见 skill）：
    # Named Spaces → NamedSpFilter(SpaceAdapter) + 1E5DA248 引用空间对象
    #   （体积盒逐对象进 volScope，多体积并集）；
    # Planes → RefPlanesFilter（全表即条件，无 ProgID 列）+ 1E5DA245 引用
    #   平面对象组（ElevPlane/GPX/GPY 等零厚度平面，单平面当盒会失真，
    #   须组内 MIN/MAX 聚合成围合盒，条件粒度一盒）
    vol_conds = {str(r[0]) for r in db.query(
        "SELECT oid FROM CMNAPPCPNamedSpFilter "
        f"WHERE bstrAdapterProgID = '{_ADAPTER_SPACE}'")}
    planes_conds: Set[str] = set()
    try:
        planes_conds = {str(r[0]) for r in db.query(
            "SELECT oid FROM CMNAPPCPRefPlanesFilter")}
    except Exception:
        pass  # 老库无此表（未见 Planes 条件），忽略
    obj_oids = {str(r[0]) for r in db.query("SELECT oid FROM CMNAPPCPObjTypeFilter")}
    compound: Dict[str, List[str]] = {}
    for o, t in db.query(
            "SELECT oid, oidTarget FROM CORERelationOrigin "
            f"WHERE RelationType = {_gl(db, _FILTER_COMPOUND)}"):
        compound.setdefault(str(o), []).append(str(t))
    refs: Dict[str, List[str]] = {}
    for o, t in db.query(
            "SELECT oid, oidTarget FROM CORERelationOrigin "
            f"WHERE RelationType IN ({_gl(db, _FILTER_HIER)}, "
            f"{_gl(db, _FILTER_HIER_PLANES)})"):
        refs.setdefault(str(o), []).append(str(t))
    contain: Dict[str, List[str]] = {}
    parent_of: Dict[str, List[str]] = {}
    for rt in (_FILTER_FOLDER, _FILTER_FOLDERSUB):
        for o, t in db.query(
                "SELECT oid, oidTarget FROM CORERelationOrigin "
                f"WHERE RelationType = {_gl(db, rt)}"):
            contain.setdefault(str(o), []).append(str(t))
            parent_of.setdefault(str(t), []).append(str(o))

    def _collect(creator_uid: Optional[str]):
        # creator_uid 为 None → 公共区（Category 1）；否则个人区
        # （Category 0 且创建者 = 该用户 SPF GUID）
        if creator_uid:
            join = "JOIN COREBaseClass c ON c.oid = f.oid"
            where = (f"WHERE f.Category = 0 AND "
                     f"c.UIDCreator = {_gl(db, creator_uid)}")
        else:
            join, where = "", "WHERE f.Category = 1"
        folders = {str(o): str(n or "") for o, n in db.query(
            "SELECT f.oid, f.bstrName FROM CMNAPPCPFilterFolder f "
            f"{join} {where}")}
        filters: Dict[str, dict] = {}
        for o, name in db.query(
                "SELECT f.oid, f.FilterName FROM CMNAPPCPUserFilter f "
                f"{join} {where}"):
            key, name = str(o), str(name or "")
            # Sym_ 前缀 = SP3D 自动生成的符号过滤器（不入文件夹、仅对象
            # 类型条件、不能用于统计），SP3D 界面不显示，跳过
            if not name or name.upper().startswith("SYM_"):
                continue
            oids: Set[str] = set()
            vols: Set[str] = set()
            pboxes: List[List[float]] = []
            pgroups: List[List[str]] = []
            kind = "none"
            stack, seen = [key], set()
            while stack:
                cur = stack.pop()
                if cur in seen:
                    continue
                seen.add(cur)
                if cur in hier_oids:
                    kind = "sys"
                    oids.update(refs.get(cur, []))
                elif cur in planes_conds:
                    # Planes 条件：组内平面聚合围合盒（一条件一盒）；
                    # 名字渲染延后到 filter_node（name_of 彼时才构建）
                    po = refs.get(cur, [])
                    if po:
                        row = db.query(
                            "SELECT MIN(xmin), MIN(ymin), MIN(zmin), "
                            "MAX(xmax), MAX(ymax), MAX(zmax) "
                            f"FROM CORESpatialIndex WHERE oid IN ({_oid_list(db, po)})")
                        if row and row[0][0] is not None:
                            pboxes.append([float(x) for x in row[0]])
                            pgroups.append(po)
                    if kind == "none":
                        kind = "vol"
                elif cur in vol_conds:
                    vols.update(refs.get(cur, []))
                    if kind == "none":
                        kind = "vol"
                elif cur in obj_oids and kind == "none":
                    kind = "obj"
                stack.extend(compound.get(cur, []))
            filters[key] = {"name": name, "kind": kind, "oids": sorted(oids),
                            "vol_oids": sorted(vols),
                            "vol_boxes": pboxes, "vol_pgroups": pgroups}
        # SQL 过滤器（CMNAPPCPSqlFilter）：不在 UserFilter 登记表，自带
        # Name/Category；与 UserFilter 同一 oid 空间，混入 fs 即经现有
        # 文件夹关系（3846127C target 已收录）自动挂树
        for o, name, qt, rc in db.query(
                "SELECT s.oid, s.Name, s.QueryText, s.ResultColumnName "
                f"FROM CMNAPPCPSqlFilter s {join.replace('f.', 's.')} "
                + where.replace("f.", "s.")):
            key, name = str(o), str(name or "")
            if not name or name.upper().startswith("SYM_"):
                continue
            filters[key] = {"name": name, "kind": "sql", "oids": [],
                            "vol_oids": [],
                            "sql_text": str(qt or ""),
                            "sql_result_col": str(rc or "")}
        return folders, filters

    # Windows 账号 → 该用户的 SPF GUID（个人过滤器归属键）
    my_uid = None
    if win_user:
        rows = db.query(
            "SELECT c.UIDCreator FROM CMNAPPCPGenericUserRoot r "
            "JOIN COREBaseClass c ON c.oid = r.oid "
            f"WHERE r.strUserID = '{safe_sql(win_user)}'")
        if rows:
            my_uid = str(rows[0][0])

    folders, filters = _collect(None)
    my_folders, my_filters = _collect(my_uid) if my_uid else ({}, {})

    # 引用系统节点的名字（显示用）：与名字表 JOIN 一次取回
    name_of: Dict[str, str] = {}
    for o, n in db.query(
            "SELECT r.oidTarget, j.ItemName FROM "
            "(SELECT DISTINCT oidTarget FROM CORERelationOrigin "
            f"WHERE RelationType IN ({_gl(db, _FILTER_HIER)}, "
            f"{_gl(db, _FILTER_HIER_PLANES)})) r "
            "LEFT JOIN JNamedItem j ON j.oid = r.oidTarget"):
        name_of[str(o)] = str(n or o)

    def filter_node(fs: Dict[str, dict], key: str) -> dict:
        f = fs[key]
        vo = f.get("vol_oids", [])
        vb = f.get("vol_boxes") or []
        node = {"type": "filter", "name": f["name"], "kind": f["kind"],
                "oids": f["oids"],
                "sys_names": [name_of.get(x, x) for x in f["oids"]],
                "vol_oids": vo,
                "vol_names": [name_of.get(x, x) for x in vo],
                "vol_boxes": vb,
                "vol_pnames": ["、".join(name_of.get(x, x[:8])
                                         for x in g[:4])
                               + ("…" if len(g) > 4 else "")
                               for g in f.get("vol_pgroups") or []]}
        if f["kind"] == "sql":
            node["sql_text"] = f["sql_text"]
            node["sql_result_col"] = f["sql_result_col"]
        return node

    def folder_node(folders: Dict[str, str], fs: Dict[str, dict],
                    key: str, path: tuple) -> dict:
        # path 为当前祖先链，防文件夹关系成环时死递归
        kids = contain.get(key, [])
        children = [folder_node(folders, fs, k, path + (k,)) for k in sorted(
            (x for x in kids if x in folders and x not in path),
            key=lambda x: folders[x].lower())]
        children += [filter_node(fs, k) for k in sorted(
            (x for x in kids if x in fs),
            key=lambda x: fs[x]["name"].lower())]
        return {"type": "folder", "name": folders[key], "children": children}

    def build(folders: Dict[str, str], fs: Dict[str, dict]) -> list:
        # 官方界面只渲染文件夹体系，未入文件夹的过滤器一律不显示
        roots = [k for k in folders
                 if not any(p in folders for p in parent_of.get(k, []))]
        return [folder_node(folders, fs, k, (k,))
                for k in sorted(roots, key=lambda x: folders[x].lower())]

    tree = build(folders, filters)

    my_tree = build(my_folders, my_filters)
    if len(my_tree) == 1 and my_tree[0]["type"] == "folder":
        # 剥掉单一的 My Filters 根文件夹壳，子项直接挂「My Filters」区块
        my_tree = my_tree[0]["children"]

    return tree, my_tree
