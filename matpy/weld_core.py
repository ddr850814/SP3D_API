# -*- coding: utf-8 -*-
"""核心统计逻辑：焊口提取 + 管子聚合（Form1.vb 的 SQL 直连路径逐行翻译）。

不依赖 SP3D API，双后端（RDB/CDB/Schema 三库）：
- SQL Server：跨库三段名 [db].dbo.T、GUID 字面量 '{uuid}'；
- Oracle：跨 schema 二段名 db.T、GUID 字面量 HEXTORAW('大端 hex')、
  RAW 列前缀匹配走 RAWTOHEX、ISNULL 换 NVL——经 _t/_gl/nvl 按
  连接的 provider 分流，业务查询主体两种后端共用同一份 SQL 骨架。
"""
from dataclasses import dataclass, field
import re
from typing import Callable, Dict, List, Optional, Set, Tuple

from db import Db

# ------------------------------------------------------------- 单位换算 ----

# 标称口径 mm -> 标称 NPS 英寸（SP3D UnitType.Npd 域的标准名义尺寸表）。
# 对应原版 m_UOM.ConvertUnitToUnit(UnitType.Npd, NPD_MILLIMETER, NPD_INCH)，
# 不是简单 ÷25.4：50mm 应得 2"，而不是 1.968503937…。
_NPD_MM_TO_INCH = [
    (6.0, 0.125), (8.0, 0.25), (10.0, 0.375), (15.0, 0.5), (20.0, 0.75),
    (25.0, 1.0), (32.0, 1.25), (40.0, 1.5), (50.0, 2.0), (65.0, 2.5),
    (80.0, 3.0), (90.0, 3.5), (100.0, 4.0), (125.0, 5.0), (150.0, 6.0),
    (200.0, 8.0), (250.0, 10.0), (300.0, 12.0), (350.0, 14.0), (400.0, 16.0),
    (450.0, 18.0), (500.0, 20.0), (550.0, 22.0), (600.0, 24.0),
    (650.0, 26.0), (700.0, 28.0), (750.0, 30.0), (800.0, 32.0), (850.0, 34.0),
    (900.0, 36.0), (950.0, 38.0), (1000.0, 40.0), (1050.0, 42.0),
    (1100.0, 44.0), (1150.0, 46.0), (1200.0, 48.0),
]

# 运行时从 CDB JNPDEquivalenceRule 加载的 mm -> 英寸 精确映射（库优先）。
# 加载失败或查无数据时保持空，_mm_to_nominal_inch 回退硬编码表。
_NPD_MM_TO_INCH_DB: Dict[float, float] = {}


def load_npd_equivalence(db: Db, cdb_name: str) -> Dict[float, float]:
    """从 CDB 管道等级库读 JNPDEquivalenceRule：mm -> 标称 NPS 英寸。

    表结构（SP3D 标准）：oid / NPDInPrimaryUnitsOfMeasure / PrimaryUnitsOfMeasure /
    NPDInSecondaryUnitsOfMeasure / SecondaryUnitsOfMeasure，一行含 in 与 mm 两侧值。
    返回 {mm: inch}；表不存在/查询失败返回空 dict（调用方回退硬编码表）。
    """
    d: Dict[float, float] = {}
    sql = (f"SELECT NPDInPrimaryUnitsOfMeasure, PrimaryUnitsOfMeasure, "
           f"NPDInSecondaryUnitsOfMeasure, SecondaryUnitsOfMeasure "
           f"FROM {_t(db, cdb_name, 'JNPDEquivalenceRule')}")
    for v1, u1, v2, u2 in db.query(sql):
        if v1 is None or v2 is None:
            continue
        try:
            f1, f2 = float(v1), float(v2)
        except (TypeError, ValueError):
            continue
        if (u1 or "").lower() == "mm" and (u2 or "").lower() == "in":
            d[round(f1, 2)] = f2
        elif (u1 or "").lower() == "in" and (u2 or "").lower() == "mm":
            d[round(f2, 2)] = f1
    return d


def _mm_to_nominal_inch(mm: float) -> float:
    """标称 mm -> 最近标称 NPS 英寸（对应 UOM 的 Npd 域换算）。

    库映射（JNPDEquivalenceRule）最近值优先（含 115→4.5 等硬编码没有的非标口径——
    115 是项目 mm 侧 NPD 公称档位，与 DN/ASME 外径无关，DN 系列没有 4.5" 档）；
    库表未加载时回退硬编码最近值。
    """
    if _NPD_MM_TO_INCH_DB:
        best, best_d = 0.0, float("inf")
        for mm_i, inch in _NPD_MM_TO_INCH_DB.items():
            d = abs(mm - mm_i)
            if d < best_d:
                best, best_d = inch, d
        return best
    best, best_d = 0.0, float("inf")
    for mm_i, inch in _NPD_MM_TO_INCH:
        d = abs(mm - mm_i)
        if d < best_d:
            best, best_d = inch, d
    return best


def convert_npd_value(npd: float, npd_unit_type: str) -> float:
    """NPD 单位 mm -> 标称 NPS 英寸；否则原值。对应 GetConvertedNPD。"""
    if npd_unit_type == "mm":
        return _mm_to_nominal_inch(npd)
    return npd


def wall_thickness_from_meters(thickness_m: float) -> float:
    """壁厚 m -> mm；<10 按 10 计，否则保留两位。对应 WallThicknessFromMeters。"""
    mm = thickness_m * 1000.0
    return 10.0 if mm < 10 else round(mm, 2)


def safe_sql(value: str) -> str:
    """转义单引号，避免拼接 SQL 报错/注入。"""
    return (value or "").replace("'", "''")


def _t(db: Db, dbname: str, table: str) -> str:
    """跨库表名：MSSQL 三段名 [db].dbo.T；ORACLE 同实例跨 schema 直接 db.T。"""
    if getattr(db, "provider", "mssql") == "oracle":
        return f"{dbname}.{table}"
    return f"[{dbname}].dbo.{table}"


def _gl(db: Db, guid: str) -> str:
    """GUID 字面量：MSSQL '{...}'；ORACLE HEXTORAW('大端 hex')（SP3D 视图定义同款）。"""
    if getattr(db, "provider", "mssql") == "oracle":
        return f"HEXTORAW('{guid.replace('-', '').upper()}')"
    return f"'{guid}'"


# ------------------------------------------------------------- 数据载体 ----

@dataclass
class PortAttr:
    sch: int = 0
    wt: float = 0.0
    npd: float = 0.0
    unit: str = ""


@dataclass(frozen=True)
class WeldInfo:
    """焊口/管子行的合并键，对应 WeldInfo 类（Equals/GetHashCode 全字段）。

    pipeline_oid 参与分组但不导出：旧版按管线对象 ObjectID 分组，同名管线
    拆成多行；导出时 pipeline_id（管线名）与旧版一致。
    """
    pipeline_id: str
    npd: float
    spec: str = ""
    materials_group: str = ""
    schedule_thickness: str = ""
    wall_thickness: float = 0.0
    pipeline_oid: str = ""


@dataclass
class StatsOptions:
    """一次统计的选项（对应 Form1 控件状态）。"""
    clr: Dict[str, float] = field(default_factory=dict)      # 规格 -> 切割长度
    mg: Dict[str, str] = field(default_factory=dict)         # 规格 -> 材料组
    met: int = 0                                            # RadDropDownList2.SelectedIndex
    thk: bool = False                                       # RadCheckBox1（忽略壁厚）
    sys_names: List[str] = field(default_factory=list)      # 勾选系统显示名（仅用于界面/日志）
    sys_oids: List[str] = field(default_factory=list)       # 勾选系统节点 oid（范围过滤用：
                                                            # 同名系统不止一个，按名字会误收
                                                            # 其他分支，必须按 oid，对齐旧版
                                                            # AddHierarchy 勾选具体节点对象）
    vol_oids: List[str] = field(default_factory=list)       # 范围体积对象 oid（Volume 过滤器，
                                                            # 焊口自身点状盒中心点判定，
                                                            # 无盒保留不过滤）
    vol_boxes: List[List[float]] = field(default_factory=list)  # Planes 条件平面组
                                                            # 围合盒（Named 体积之外
                                                            # 的并集来源，语义同上）
    skip_pipes: bool = False                                # 忽略直管焊缝（勾选后
                                                            # aggregate_pipes 整体跳过，
                                                            # 等价原版规格表全部删行）


# ------------------------------------------------------------- 字典加载 ----

PairDict = Dict[str, List[str]]          # key(hex oid) -> [value oid hex]
SingleDict = Dict[str, str]              # key oid hex -> value oid hex


def load_pair_dict(db: Db, sql: str) -> PairDict:
    d: PairDict = {}
    for row in db.query(sql):
        k, v = row[0], row[1]
        if k is None or v is None:
            continue
        d.setdefault(str(k), []).append(str(v))
    return d


def load_single_dict(db: Db, sql: str) -> SingleDict:
    d: SingleDict = {}
    for row in db.query(sql):
        k, v = row[0], row[1]
        if k is None or v is None:
            continue
        d[str(k)] = str(v)
    return d


def load_single_str(db: Db, sql: str) -> Dict[str, str]:
    d: Dict[str, str] = {}
    for row in db.query(sql):
        k = row[0]
        if k is None:
            continue
        d[str(k)] = str(row[1])
    return d


def load_port_attrs(db: Db) -> Dict[str, PortAttr]:
    d: Dict[str, PortAttr] = {}
    for row in db.query(
        "SELECT oid, Npd, NpdUnitType, ScheduleThickness, WallThicknessOrGrooveSetback FROM JDPipePort"
    ):
        d[str(row[0])] = PortAttr(
            sch=int(row[3] or 0),
            wt=float(row[4] or 0.0),
            npd=float(row[1] or 0.0),
            unit=str(row[2] or ""),
        )
    return d


def load_sched_text(db: Db, schema_db: str) -> Dict[int, str]:
    d: Dict[int, str] = {}
    sql = (f"SELECT ValueID, ShortStringValue "
           f"FROM {_t(db, schema_db, 'CodelistValueView')} "
           f"WHERE TableName = 'ScheduleThickness'")
    for row in db.query(sql):
        d[int(row[0])] = str(row[1])
    return d


# ------------------------------------------------------------- 辅助函数 ----

def first_other(d: PairDict, key: str, exclude: str) -> Optional[str]:
    """列表中第一个 != exclude 的元素（对应 FirstOther）。"""
    lst = d.get(key)
    if lst:
        for g in lst:
            if g != exclude:
                return g
    return None


# MDB CORERelationOrigin 关系类型（代理端口 → catalog 定义端口链路）
_MDB_REL_CPNAME = "5280312B-E69C-11D1-A966-080036069A02"   # CP 点命名（RelationName）
_MDB_REL_PART2DEF = "B0E39EC4"    # 部件实例(00013885) -> 目录定义(00000008)
_MDB_REL_PART2PORT = "BFE914B5"   # 部件实例 -> 端口


def load_cp_names(model_db: Db) -> Dict[str, str]:
    """MDB CORERelationOrigin(5280312B)：对象 oid -> CP 点名称（RelationName）。

    同一命名关系给两类对象命名（旧 API IJDPipePort 属性解析的关键）：
    - 代理端口（CReferenceProxy）：PNoz1/PNoz2（= 目录定义端口 PortIndex）；
    - 部件的目录定义（00000008）：定义名 = CDB PartNumber
      （如 PBREC575DS-40SS-40S2515）。
    """
    d: Dict[str, str] = {}
    for o, name in model_db.query(
        "SELECT oidTarget, RelationName FROM CORERelationOrigin "
        f"WHERE RelationType = {_gl(model_db, _MDB_REL_CPNAME)} "
        "AND RelationName IS NOT NULL"
    ):
        if o is None or name is None:
            continue
        d[str(o)] = str(name)
    return d


def load_mdb_partnumber(model_db: Db) -> Dict[str, str]:
    """MDB REFDATPipeComponent：目录定义 oid -> PartNumber。

    阀门/消音器（0000EA74 前缀，如 CSInlineSilencer4）的目录定义在 MDB
    CORERelationOrigin(5280312B) 中没有命名记录（无 PartNumber 名），
    但 MDB 自身 REFDATPipeComponent 表按 oid 直接存有 PartNumber。
    用户确认：component 与 part 通过 PartNumber 连接（而非 oid）。
    """
    d: Dict[str, str] = {}
    for o, pn in model_db.query(
        "SELECT oid, PartNumber FROM REFDATPipeComponent "
        "WHERE PartNumber IS NOT NULL"
    ):
        if o is None or pn is None:
            continue
        d[str(o)] = str(pn)
    return d


def load_part_defs(model_db: Db) -> Dict[str, str]:
    """MDB 关系链（B0E39EC4）：部件实例 oid -> 目录定义 oid。

    管道目录件判定（阀门/仪表/设备接口件的对侧端口替换，配合
    load_pipe_defs）依赖它，与代理端口无关，总是加载。
    """
    # RAW 列不能 LIKE：Oracle 侧先 RAWTOHEX 转大端 hex 再前缀匹配
    rel_col = ("RAWTOHEX(RelationType)"
               if getattr(model_db, "provider", "mssql") == "oracle" else "RelationType")
    d: Dict[str, str] = {}
    for a, b in model_db.query(
        "SELECT oid, oidTarget FROM CORERelationOrigin "
        f"WHERE {rel_col} LIKE '{_MDB_REL_PART2DEF}%'"
    ):
        if a is None or b is None:
            continue
        d[str(a)] = str(b)
    return d


def load_pipe_defs(model_db: Db) -> frozenset:
    """MDB 管道目录件定义 oid 集合（REFDATPipeComponent 全表）。

    部件定义 oid 落在此集合 = 管道目录件（阀门/仪表/设备接口件），
    替代按定义 oid 前缀 0000EA74 的判定（表本身即该类的完整成员集）。
    """
    return frozenset(str(r[0]) for r in
                     model_db.query("SELECT oid FROM REFDATPipeComponent"))


def load_port_parts(model_db: Db) -> Dict[str, str]:
    """MDB 关系链（BFE914B5）：代理/真实端口 oid -> 部件实例 oid。

    仅代理端口穿透需要；伴随 JReferenceProxy/JDPipePort/JRtePipePart
    三张全表 oid 集合（合计 ~150 万行），由调用方按需加载。

    BFE914B5：部件 ↔ 端口（双向出现）；端口 = JReferenceProxy ∪
    JDPipePort（仅认 Proxy 会漏真实端口）；部件 = JRtePipePart 成员
    （管道件全类存储，桥架等非管道件自然排除；实测与旧前缀法
    管道件映射 0 差异）。
    """
    proxies = {str(r[0]) for r in model_db.query("SELECT oid FROM JReferenceProxy")}
    jdp = {str(r[0]) for r in model_db.query("SELECT oid FROM JDPipePort")}
    parts = {str(r[0]) for r in model_db.query("SELECT oid FROM JRtePipePart")}
    rel_col = ("RAWTOHEX(RelationType)"
               if getattr(model_db, "provider", "mssql") == "oracle" else "RelationType")
    d_port_part: Dict[str, str] = {}
    for a, b in model_db.query(
        "SELECT oid, oidTarget FROM CORERelationOrigin "
        f"WHERE {rel_col} LIKE '{_MDB_REL_PART2PORT}%'"
    ):
        if a is None or b is None:
            continue
        ka, kb = str(a), str(b)
        if ka in proxies and kb not in proxies and kb in parts:
            d_port_part[ka] = kb
        elif kb in proxies and ka not in proxies and ka in parts:
            d_port_part[kb] = ka
        elif ka in jdp and kb not in jdp and kb in parts:
            d_port_part[ka] = kb
        elif kb in jdp and ka not in jdp and ka in parts:
            d_port_part[kb] = ka
    return d_port_part


def load_catalog_port_attr(db: Db, cdb_name: str) -> Dict[Tuple[str, int], PortAttr]:
    """CDB 目录定义端口属性：(PartNumber, PortIndex) -> PortAttr(Npd+表号+壁厚)。

    链路：REFDATPipeComponent --CORErelations--> REFDATPipePort
    （PortIndex/Npd/表号）--同 oid--> JCatalogPipePort（数值壁厚 m）。
    MDB 无这些表，须在 RDB 连接上跨库查 CDB。
    """
    d: Dict[Tuple[str, int], PortAttr] = {}
    sql = (f"SELECT A.PartNumber, x.PortIndex, x.Npd, x.NpdUnitType, "
           f"x.ScheduleThickness, j.WallThicknessOrGrooveSetback "
           f"FROM {_t(db, cdb_name, 'REFDATPipeComponent')} A "
           f"JOIN {_t(db, cdb_name, 'CORErelations')} B ON A.oid = B.IDorigin "
           f"JOIN {_t(db, cdb_name, 'REFDATPipePort')} x ON B.IDdestination = x.oid "
           f"LEFT JOIN {_t(db, cdb_name, 'JCatalogPipePort')} j ON j.oid = x.oid "
           f"WHERE A.PartNumber IS NOT NULL AND x.PortIndex IS NOT NULL AND x.Npd > 0")
    for pn, pidx, npd, unit, sch, wt in db.query(sql):
        try:
            v = float(npd)
            i = int(pidx)
        except (TypeError, ValueError):
            continue
        d[(str(pn), i)] = PortAttr(
            sch=int(sch or 0),
            wt=float(wt or 0.0),
            npd=v,
            unit=(unit or "").lower(),
        )
    return d


def resolve_port_attr_catalog(port_oid: str, d_cp_names: Dict[str, str],
                              d_port_part_mdb: Dict[str, str],
                              d_part_def: Dict[str, str],
                              d_cat: Dict[Tuple[str, int], PortAttr],
                              d_mdb_pn: Optional[Dict[str, str]] = None) -> Optional[PortAttr]:
    """代理端口判径主路径：port -> 部件 -> 目录定义 -> CDB 端口属性。

    代理端口不落 JDPipePort 属性表，这里用 MDB 关系复现 SP3D 内部解析：
    5280312B CP 名（PNoz1/2 = PortIndex 1/2）→ 部件 → 目录定义（定义名
    = PartNumber，阀门等无命名的回退 d_mdb_pn）→
    REFDATPipePartAndPortDataView 取 Npd/表号。链上任一环缺失返回 None。
    """
    part = d_port_part_mdb.get(port_oid)
    if part is None:
        return None
    def_oid = d_part_def.get(part)
    if def_oid is None:
        return None
    pn = d_cp_names.get(def_oid) or (d_mdb_pn or {}).get(def_oid)
    if not pn:
        return None
    # CP 名尾数字 = 目录定义端口 PortIndex（SP3D 命名惯例）：
    # 三通/支管台 PNoz1/PNoz2/PNoz3（PNoz3=支管口）、阀门 Nozzle1/Nozzle2。
    m = re.search(r"(\d+)$", d_cp_names.get(port_oid) or "")
    if m is None:
        return None
    return d_cat.get((pn, int(m.group(1))))


# ------------------------------------------------------------- 焊口提取 ----

def _oid_list(db: Db, oids: List[str]) -> str:
    # oid 字面量列表：Oracle 下 oid 列是 RAW(16)（大端原序，db.query 已按
    # RAWTOHEX 同款输出 plain hex），HEXTORAW 直出即物理字节；MSSQL '{oid}' 原样
    return ", ".join(_gl(db, o) for o in dict.fromkeys(oids))


# PathRunUsesSpec 关系类型与其「名字载体」关系类型（SP3D 全局 GUID）。
# 视图 ~100s → MDB 关系直查 ~2s，替换依据与对齐细节见 skill「管道等级」。
_A53D_PATHRUNUSESSPEC = "HEXTORAW('A53D19B621DE4563B7BC03DE5D632E29')"
_5280_NAME_CARRIER = "HEXTORAW('5280312BE69C11D1A966080036069A02')"


def _run_spec_pairs_sql(mdb_schema: str) -> str:
    """run oid -> 规格名 对：直查 MDB.CoRERelationOrigin 自 JOIN。"""
    return (f"SELECT r1.oid, r2.RelationName "
            f"FROM {mdb_schema}.CoRERelationOrigin r1 "
            f"LEFT JOIN {mdb_schema}.CoRERelationOrigin r2 "
            f"ON (r1.oidTarget = r2.oidTarget "
            f"AND r2.RelationType = {_5280_NAME_CARRIER}) "
            f"WHERE r1.RelationType = {_A53D_PATHRUNUSESSPEC}")


def _load_run_spec(db: Db, mdb_schema: Optional[str]) -> Dict[str, str]:
    """d_run_spec：run oid -> SpecName。Oracle+MDB 直查（快 50 倍），
    其余后端/降级走原 XPathRunUsesSpec 视图。"""
    if mdb_schema and getattr(db, "provider", "mssql") == "oracle":
        spec_names = {str(n) for (n,) in db.query("SELECT SpecName FROM JDPipeSpec")}
        return {str(k): str(v) for k, v in
                db.query(_run_spec_pairs_sql(safe_sql(mdb_schema)))
                if v is not None and str(v) in spec_names}
    return load_single_str(
        db, "SELECT xs.oidOrigin, sd.SpecName FROM XPathRunUsesSpec xs "
            "JOIN JDPipeSpec sd ON sd.oid = xs.oidDestination")


def extract_welds_via_sql(db: Db, schema_db: str, cdb_name: str,
                          opts: StatsOptions,
                          on_log: Callable[[str], None] = None,
                          model_db: Optional[Db] = None,
                          mdb_schema: Optional[str] = None) -> Dict[WeldInfo, int]:
    """焊口阶段 SQL 化：Q1 基础行 + Q2-Q6 关系字典 + 逐焊口归属（对应 ExtractWeldsViaSql）。

    model_db：可选 MDB 连接，供读不到属性的代理端口（CReferenceProxy）
    做规格穿透，失败自动降级不影响主流程。SQL 路径无产出时返回空 dict。
    mdb_schema：MDB schema 名（Oracle 专用优化：绕开 XPathRunUsesSpec 慢视图）。
    """
    log = on_log or (lambda m: None)
    nvl = "NVL" if getattr(db, "provider", "mssql") == "oracle" else "ISNULL"

    # 范围过滤对齐原版 API：AddHierarchy 在焊口集合上按层次归属过滤，实验证明
    # 等价于支管归属切换后的 final_run（而非焊口所在 run0）在勾选范围内。
    # 恒取范围内（IN 语义，无范围反选）；循环内按 final_run 判定。
    scope_oids: Optional[Set[str]] = None
    if opts.sys_oids:
        # 范围按勾选节点 oid 匹配（同名系统不止一个，按 ItemName 会误收其他分支）
        scope_q = f"""WITH sysScope (oid) AS (
SELECT h.oid FROM JNamedItem h WHERE h.oid IN ({_oid_list(db, opts.sys_oids)})
UNION ALL
SELECT x.oidDestination FROM XSystemHierarchy x JOIN sysScope s ON x.oidOrigin = s.oid
)
SELECT oid FROM sysScope"""
        scope_oids = {str(r[0]) for r in db.query(scope_q)}
        log(f"系统范围：勾选 {len(opts.sys_oids)} 个系统，展开子孙共 {len(scope_oids)} 个 oid")

    def _in_scope(run_oid: Optional[str]) -> bool:
        # 恒取范围内（IN 语义）；无范围时全收
        if scope_oids is None:
            return True
        return run_oid is not None and str(run_oid) in scope_oids

    # 焊口级 volume 过滤（q1 行源过滤，后续归属逻辑零改动）：焊口自身
    # 盒（点状，中心点即位置）落在任一勾选体积盒内才保留；无盒子保留
    # 不过滤（找不到位置不硬丢）。与系统范围一样恒取范围内（IN 语义）。
    vol_cond = ""
    vol_cte = ""
    if opts.vol_oids or opts.vol_boxes:
        si = _t(db, mdb_schema or "", "CORESpatialIndex") if mdb_schema else ""
        # db 是 RDB 连接，CORESpatialIndex 只在 MDB（RDB 是空壳），
        # 拿不到 MDB 库名时无法跨库——退化为不过滤并记日志
        if si:
            parts = []
            if opts.vol_oids:
                parts.append(
                    f"    SELECT v.xmin, v.ymin, v.zmin, v.xmax, v.ymax, v.zmax\n"
                    f"    FROM {si} v WHERE v.oid IN ({_oid_list(db, opts.vol_oids)})")
            if opts.vol_boxes:
                is_ora = getattr(db, "provider", "mssql") == "oracle"
                cols = ("xmin", "ymin", "zmin", "xmax", "ymax", "zmax")
                rows = []
                for b in opts.vol_boxes:
                    nums = ", ".join(
                        f"{repr(float(x)) if 'e' not in repr(float(x)) else format(float(x), '.10f')} AS {c}"
                        for x, c in zip(b, cols))
                    rows.append(f"    SELECT {nums} FROM dual" if is_ora
                                else f"    SELECT {nums}")
                joiner = "\n    UNION ALL\n" if is_ora else "\n  UNION ALL\n"
                parts.append("(" + joiner.join(rows) + "\n    )")
            vol_cte = ("WITH volScope (xmin, ymin, zmin, xmax, ymax, zmax) AS (\n"
                       + "\n  UNION ALL\n".join(parts) + "\n)\n")
            vol_cond = (f" AND (w.Oid IN (SELECT s.oid FROM {si} s "
                        "JOIN volScope v ON s.xmin + s.xmax BETWEEN 2*v.xmin AND 2*v.xmax "
                        "AND s.ymin + s.ymax BETWEEN 2*v.ymin AND 2*v.ymax "
                        "AND s.zmin + s.zmax BETWEEN 2*v.zmin AND 2*v.zmax)"
                        f" OR w.Oid NOT IN (SELECT oid FROM {si}))")
        else:
            log("volume 范围未生效：缺少 MDB 库名，无法跨库取空间索引")

    q1_sel = f"""{vol_cte}SELECT g.OidOrigin AS conn0, oi.oidOrigin AS partOid, bp.oidDestination AS port0,
       op.oidOrigin AS run0, wp.PortIndex AS pidx,
       bro.BranchReportingOwnership AS broV,
       CASE WHEN bex.part IS NOT NULL THEN 1 ELSE 0 END AS isBr,
       w.Oid AS weldoid
FROM JRteWeld w
JOIN XGeneratesConnectionItems g ON g.OidDestination = w.Oid
JOIN XOwnsImpliedItems oi ON oi.oidDestination = w.Oid
JOIN XFlowPorts fp ON fp.OidOrigin = g.OidOrigin
JOIN XDistribPorts bp ON bp.oidOrigin = oi.oidOrigin AND bp.oidDestination = fp.OidDestination
LEFT JOIN JDPipePort wp ON wp.oid = bp.oidDestination
JOIN XOwnsParts op ON op.oidDestination = oi.oidOrigin
LEFT JOIN JRteBranchReportingOwnership bro ON bro.Oid = oi.oidOrigin
LEFT JOIN (SELECT DISTINCT x.OidDestination AS part FROM XPathGeneratedParts x JOIN JRtePipeBranchPathFeat bf ON bf.Oid = x.OidOrigin) bex ON bex.part = oi.oidOrigin
LEFT JOIN JMTOInfo mto ON mto.Oid = w.Oid
WHERE {nvl}(mto.ReportingRequirements, 5) <> 2{vol_cond}"""
    q1 = q1_sel
    rows = db.query(q1)
    if not rows:
        log("RDB 焊口基础查询返回 0 行（该厂 RDB 镜像不完整？），无法统计焊口")
        return {}

    # Q2-Q6 关系字典
    d_part_ports = load_pair_dict(db, "SELECT oidOrigin, oidDestination FROM XDistribPorts")
    d_conn_parts = load_pair_dict(db, "SELECT OidOrigin, OidDestination FROM XRelConnectionAndPartOcc")
    # d_part_conns = part -> connections，从 d_conn_parts 反推（同表的反向映射）
    d_part_conns: PairDict = {}
    for _k, _vals in d_conn_parts.items():
        for _v in _vals:
            d_part_conns.setdefault(_v, []).append(_k)
    d_conn_ports = load_pair_dict(db, "SELECT OidOrigin, OidDestination FROM XFlowPorts")
    d_part_run = load_single_dict(db, "SELECT oidDestination, oidOrigin FROM XOwnsParts")
    d_run_spec = _load_run_spec(db, mdb_schema)
    d_run_pipe = load_single_str(db, "SELECT hh.oidDestination, pl.ItemName FROM XSystemHierarchy hh JOIN JNamedItem pl ON pl.oid = hh.oidOrigin")
    # run -> 所属管线对象 oid（同名管线按 oid 拆开，对齐旧版 ObjectID 分组）
    d_run_pipe_oid = load_single_str(db, "SELECT hh.oidDestination, hh.oidOrigin FROM XSystemHierarchy hh")
    d_port_attr = load_port_attrs(db)
    d_sched = load_sched_text(db, schema_db)
    # d_port_part = port -> parts，从 d_part_ports 反推（同表 XDistribPorts 的反向映射）
    d_port_part: PairDict = {}
    for _k, _vals in d_part_ports.items():
        for _v in _vals:
            d_port_part.setdefault(_v, []).append(_k)

    # CDB 管道等级 NPD 等价表（mm -> 标称 NPS 英寸，库优先）。加载失败保持空字典，
    # _mm_to_nominal_inch 自动回退硬编码表，不影响主流程。
    global _NPD_MM_TO_INCH_DB
    try:
        _NPD_MM_TO_INCH_DB = load_npd_equivalence(db, cdb_name)
        if _NPD_MM_TO_INCH_DB:
            log(f"CDB NPD 等价表加载完成：{len(_NPD_MM_TO_INCH_DB)} 条 mm↔in")
        else:
            log("CDB NPD 等价表无数据，口径换算回退硬编码表")
    except Exception as ex:
        _NPD_MM_TO_INCH_DB = {}
        log(f"CDB NPD 等价表加载失败，回退硬编码表：{ex}")

    # 代理端口判径链：MDB 部件链 + CDB 目录端口属性。加载失败降级为空。
    d_cp_names: Dict[str, str] = {}
    d_port_part_mdb: Dict[str, str] = {}
    d_part_def_mdb: Dict[str, str] = {}
    d_pipe_defs_mdb: frozenset = frozenset()
    d_mdb_pn: Dict[str, str] = {}
    if model_db is not None:
        try:
            # 惰性判定：B0E3 部件->定义链 + 管道目录件集合（对侧端口
            # 替换判定）总是需要；
            # 代理端口穿透链（CP 名 + BFE9 + 三张全表 oid + PartNumber，
            # 合计 ~150 万行）仅在 q1 存在「端口缺属性或缺 PortIndex」时加载
            d_part_def_mdb = load_part_defs(model_db)
            d_pipe_defs_mdb = load_pipe_defs(model_db)
            need_attr = any(
                (pa := d_port_attr.get(str(r[2]))) is None or pa.npd <= 0
                for r in rows)
            need_pidx = any(r[4] is None for r in rows)
            if need_attr or need_pidx:
                d_cp_names = load_cp_names(model_db)
                d_port_part_mdb = load_port_parts(model_db)
                d_mdb_pn = load_mdb_partnumber(model_db)
            else:
                log("RDB 端口属性与 PortIndex 完整，跳过 MDB 代理端口链"
                    "（省 ~150 万行查询）")
            log(f"MDB 链加载完成：CP 点名 {len(d_cp_names)} 条，"
                f"端口->部件 {len(d_port_part_mdb)} 条，部件->定义 {len(d_part_def_mdb)} 条，"
                f"MDB 目录 PartNumber {len(d_mdb_pn)} 条")
        except Exception as ex:
            d_cp_names = {}
            d_port_part_mdb = {}
            log(f"MDB 代理端口链加载失败，跳过：{ex}")
    d_cat_npd: Optional[Dict[Tuple[str, int], PortAttr]] = None
    if d_port_part_mdb:
        # 同属代理穿透链（resolve_port_attr_catalog 专用），跟随惰性
        try:
            d_cat_npd = load_catalog_port_attr(db, cdb_name)
            log(f"CDB 目录端口属性加载完成：{len(d_cat_npd)} 条")
        except Exception as ex:
            d_cat_npd = None
            log(f"CDB 目录端口属性加载失败：{ex}")

    welds: Dict[WeldInfo, int] = {}

    # ---- 支管台双焊缝分组（Weldolet：同一件产生 2 个 Connection/焊缝）----
    # 两条焊缝一条连主管（port2=主管管道端口，口径大）、一条连支管（port2=支管端口，
    # 口径小），即「双口径」。业务规则（用户确认）：
    #   BRO=5  → 两条都算在主管上；
    #   BRO≠5  → 两条都不算在主管上（归支管侧）；
    #   且「和主管连接的那道焊缝」口径/表号/壁厚按「和支管连接的那道」取。
    br_pairs: Dict[str, List[tuple]] = {}
    for r in rows:
        if int(r[6] or 0) == 1:
            br_pairs.setdefault(str(r[1]), []).append(r)
    br_pairs = {p: g for p, g in br_pairs.items() if len(g) == 2}

    def _port_ijd(port_oid: str) -> Optional[PortAttr]:
        """旧 API oPort.GetPropertyValue("IJDPipePort", ...) 的等价解析。

        真实端口：JDPipePort 且 Npd>0 即有效；代理端口走 MDB 链 + CDB
        （= SP3D 对 ProxyPort 的自动解析）。都取不到返回 None
        （旧 API 属性缺失抛异常 → 该焊口被跳过）。
        """
        pa = d_port_attr.get(port_oid)
        if pa is not None and pa.npd > 0:
            return pa
        if d_cat_npd is not None:
            ca = resolve_port_attr_catalog(port_oid, d_cp_names,
                                           d_port_part_mdb, d_part_def_mdb,
                                           d_cat_npd, d_mdb_pn)
            if ca is not None:
                return ca
        return None

    def _resolve_weld(r):
        """逐焊口归属 + 虚拟端口属性解析（原主循环逻辑抽出）。"""
        conn0 = str(r[0]); part_oid = str(r[1]); port0 = str(r[2]); run0 = str(r[3])
        # 代理端口（CReferenceProxy）在 JDPipePort 无记录 → pidx=NULL；
        # 用 MDB CP 点名尾数字恢复 PortIndex（PNoz1/PNoz2/PNoz3、Nozzle1/2，
        # 与原版 API IJDPipePort.PortIndex 语义一致）。
        pidx = int(r[4]) if r[4] is not None else None
        if pidx is None and d_cp_names:
            m = re.search(r"(\d+)$", d_cp_names.get(port0) or "")
            if m:
                pidx = int(m.group(1))
        if pidx is None:
            pidx = 0
        is_br = int(r[6] or 0) == 1
        bro_v = int(r[5]) if r[5] is not None else None
        other_port = first_other(d_part_ports, part_oid, port0)
        other_conn = first_other(d_part_conns, part_oid, conn0)
        branch_on = (is_br and bro_v is not None and pidx in (1, 2)
                     and other_port is not None and other_conn is not None)
        swap_on = branch_on and pidx == 1

        # run 归属：支管时取选中连接对侧件的 run
        if branch_on:
            use_main_conn = (bro_v == 5) if pidx == 1 else (bro_v != 5)
            run_conn = conn0 if use_main_conn else other_conn
            side_part = first_other(d_conn_parts, run_conn, part_oid)
            if side_part is None or side_part not in d_part_run:
                return None  # API 会因取关系失败跳过
            final_run = d_part_run[side_part]
        else:
            final_run = run0

        final_port = other_port if swap_on else port0
        thk_conn = other_conn if swap_on else conn0

        # 阀门/仪表/设备接口件（管道目录件，定义 oid 在 REFDATPipeComponent
        # 表中）：其端口无 JDPipePort 属性（P200029 实测 8/8 全无），
        # np=_port_ijd 读不到 → 照旧版 VB 行为（DistribConnectionPorts 对此类件
        # 只返回管道侧端口）取连接对侧管道端口作为 oPort，焊口处两侧口径本就一致；
        # thk_conn 保持 conn0（VB 的 oConnection 未变，仅 oPort 换为对侧端口）。
        d0 = d_part_def_mdb.get(part_oid)
        if d0 is not None and d0 in d_pipe_defs_mdb:
            alt = first_other(d_conn_ports, conn0, port0)
            if alt is not None:
                final_port = alt

        # 虚拟端口属性解析：np = oPort 的 IJDPipePort 属性（NPD/表号/壁厚同源，
        # 照抄旧 API：oPort 取不到属性则整道焊口跳过，不做 run 启发式兜底）
        port2 = first_other(d_conn_ports, thk_conn, final_port)
        np = _port_ijd(final_port)
        if np is None:
            return None
        return dict(part_oid=part_oid, bro_v=bro_v, run0=run0,
                    final_run=final_run, final_port=final_port,
                    thk_conn=thk_conn, port2=port2, np=np, pidx=pidx)

    def _npd_inch(pa):
        """口径统一比较（公称 inch 档位）：mm 侧走项目公称对照表而非
        ×25.4——NPD 取值是项目自定义档位（如 115mm≡4.5"），线性换算
        会把同档判成不同档。npd≤0 返回 0，保持空属性焊口的跳过语义。"""
        if pa.npd <= 0:
            return 0.0
        if pa.unit == "mm":
            return _mm_to_nominal_inch(pa.npd)
        return pa.npd

    def _run_of_port2(w):
        """port2（连接对侧管道端口）所属 run；取不到返回 None。"""
        p2 = w["port2"]
        if p2 is None:
            return None
        parts = d_port_part.get(p2)
        if not parts:
            return None
        for part in parts:
            run = d_part_run.get(part)
            if run is not None:
                return run
        return None

    def _build_info(final_run, final_port, thk_conn, port2, np):
        """构建 WeldInfo（管线/口径/表号/壁厚）。缺管线映射返回 None。"""
        if final_run not in d_run_spec or final_run not in d_run_pipe:
            return None
        spec_name = d_run_spec[final_run]
        info = WeldInfo(
            pipeline_id=d_run_pipe[final_run],
            npd=convert_npd_value(np.npd, np.unit),
            pipeline_oid=d_run_pipe_oid.get(final_run) or final_run,
        )
        if opts.met == 2 and spec_name in opts.mg:
            info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                            materials_group=opts.mg[spec_name],
                            pipeline_oid=info.pipeline_oid)
        elif opts.met == 1:
            info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                            spec=spec_name, pipeline_oid=info.pipeline_oid)
        if not opts.thk:
            # 表号+壁厚单通道：均取自「选定端口」pick（照抄旧 API Form1.vb
            # L99-127：ScheduleThickness 与 WallThickness 同端口读取，
            # 不做双通道混选、不做 run 启发式兜底）。
            # 对侧端口取不到 IJDPipePort 属性（幽灵端口/非管道件对侧）时，
            # 旧 API 走单端口通道（Form1.vb L124-128 Else：oPort2=Nothing
            # 直接用 oPort 的 sch/wt），不跳过整道。
            pw2 = _port_ijd(port2) if port2 is not None else None
            if pw2 is None:
                pick = np
            elif np.sch > 1 and pw2.sch > 1:
                if np.wt > 0 and pw2.wt > 0:
                    pick = pw2 if np.wt > pw2.wt else np
                elif np.wt > 0:
                    pick = np
                elif pw2.wt > 0:
                    pick = pw2
                else:
                    pick = pw2 if np.wt < pw2.wt else np
            elif pw2.sch > 1:
                pick = pw2
            else:
                # 对侧 sch=Undefined/None（如支管台 PO* 端口）→ 直接用本侧 oPort
                pick = np
            info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                            spec=info.spec, materials_group=info.materials_group,
                            schedule_thickness=d_sched.get(pick.sch) or "undefined",
                            wall_thickness=wall_thickness_from_meters(pick.wt),
                            pipeline_oid=info.pipeline_oid)
        return info

    done_pairs: Set[str] = set()
    skip_ids: Set[int] = set()

    for r in rows:
        try:
            if id(r) in skip_ids:
                continue
            is_br = int(r[6] or 0) == 1
            part_oid = str(r[1])
            if is_br and part_oid in br_pairs and part_oid not in done_pairs:
                done_pairs.add(part_oid)
                pair = br_pairs[part_oid]
                skip_ids.add(id(pair[1]))
                w1 = _resolve_weld(pair[0])
                w2 = _resolve_weld(pair[1])
                merged = False
                if w1 is not None and w2 is not None \
                        and w1["bro_v"] is not None and w1["bro_v"] == w2["bro_v"]:
                    # PortIndex 优先判定主支：1=主管侧, 2=支管侧
                    # pidx=0 表示无法恢复（JDPipePort 无记录且 CP 名无尾数字），
                    # fallback 到 NPD 比大小（仅双口径不同时成立）。
                    hdr = br = None
                    p1, p2 = w1.get("pidx", 0), w2.get("pidx", 0)
                    if {p1, p2} == {1, 2}:
                        hdr, br = (w1, w2) if p1 == 1 else (w2, w1)
                    else:
                        m1, m2 = _npd_inch(w1["np"]), _npd_inch(w2["np"])
                        if m1 > 0 and m2 > 0 and m1 != m2:
                            hdr, br = (w1, w2) if m1 > m2 else (w2, w1)
                    if hdr is not None and br is not None:
                        # BRO=5 两条都算在主管上：主管管线 = 主管侧焊缝(PortIndex 1)连接
                        # 对侧管道端口所属 run（隐含件 part 自身可能归属支管管线 run，
                        # 如 weldolet 归 VT-11008 但焊在 VT-11005 3" 主管上，
                        # 不能再用 final_run=run0 反推）。
                        # BRO≠5 两条都算在支管侧：取支管侧焊缝的归属 run。
                        if hdr["bro_v"] == 5:
                            target_run = _run_of_port2(hdr) or hdr["final_run"]
                        else:
                            target_run = br["final_run"]
                        # 两条都算在 target 管线：主管侧那条数据按支管侧那条取
                        info = None
                        if not _in_scope(target_run):
                            merged = True  # 范围外：整组不计，避免退回 fallback 重复计数
                        else:
                            info = _build_info(target_run, br["final_port"],
                                               br["thk_conn"], br["port2"], br["np"])
                            if info is not None:
                                welds[info] = welds.get(info, 0) + 2
                                merged = True
                if not merged:
                    # 分组不成立（缺关系/口径无法判定/非双口径）：退回逐条处理
                    for w in (w1, w2):
                        if w is None:
                            continue
                        if not _in_scope(w["final_run"]):
                            continue
                        info = _build_info(w["final_run"], w["final_port"],
                                           w["thk_conn"], w["port2"], w["np"])
                        if info is not None:
                            welds[info] = welds.get(info, 0) + 1
                continue

            w = _resolve_weld(r)
            if w is None:
                continue
            if not _in_scope(w["final_run"]):
                continue
            info = _build_info(w["final_run"], w["final_port"],
                               w["thk_conn"], w["port2"], w["np"])
            if info is not None:
                welds[info] = welds.get(info, 0) + 1
        except Exception as ex:
            import traceback
            log(f"跳过焊口: {ex}\n{traceback.format_exc()}")
            continue
    return welds


# ------------------------------------------------------------- 管子聚合 ----

def aggregate_pipes(db: Db, schema_db: str, opts: StatsOptions,
                    on_log: Callable[[str], None] = None,
                    mdb_schema: Optional[str] = None) -> Dict[WeldInfo, int]:
    """管子阶段聚合 SQL（对应 RadButton1_Click 的 MSSQL 分支）。

    Oracle 走 _aggregate_pipes_ora 字典方案（XPathRunUsesSpec 视图 ~105s
    且 JOIN 进大 SQL 无解——此前试过 CoRERelationOrigin 子查询反而
    >9 分钟）；MSSQL 视图有索引、基线 31.7s，SQL 版原样保留。
    mdb_schema：MDB schema 名，Oracle 直查规格关系用（缺省回退视图）。
    """
    log = on_log or (lambda m: None)
    nvl = "NVL" if getattr(db, "provider", "mssql") == "oracle" else "ISNULL"
    ceil = "CEIL" if getattr(db, "provider", "mssql") == "oracle" else "CEILING"
    if opts.skip_pipes:
        log("已勾选忽略直管焊缝，跳过管子聚合")
        return {}
    # 切割长度必须为正数：<=0 会造成 SQL 除零/负除，直接剔除并记录
    clr = {k: v for k, v in opts.clr.items()
           if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0}
    skipped = sorted(set(opts.clr) - set(clr))
    if skipped:
        log("切割长度非正数，已跳过规格: " + ", ".join(skipped))
    if not clr:
        log("无有效切割长度，跳过管子聚合")
        return {}
    if getattr(db, "provider", "mssql") == "oracle":
        return _aggregate_pipes_ora(db, schema_db, opts, clr, log, mdb_schema)
    spec_list = ", ".join(f"'{safe_sql(k)}'" for k in clr)
    # 搜索式 CASE（简单 CASE 在 Oracle 下 NVARCHAR2 选择器混排字面量会 ORA-12704）
    cut_case = "CASE " + " ".join(
        f"WHEN sd.SpecName = '{safe_sql(k)}' THEN {_fmt_double(v)}"
        for k, v in clr.items()) + " END"
    if opts.sys_oids:
        # 有系统范围：递归 CTE 展开选中系统及其全部子孙（恒取范围内）。
        # 按勾选节点 oid 匹配——同名系统不止一个，按 ItemName 会误收其他分支。
        # 无前导分号：T-SQL 老习惯的 ";WITH" 在 Oracle 单条执行会 ORA-00911
        sys_scope_cte = f"""WITH sysScope (oid) AS (
    SELECT h.oid FROM JNamedItem h WHERE h.oid IN ({_oid_list(db, opts.sys_oids)})
    UNION ALL
    SELECT x.oidDestination FROM XSystemHierarchy x JOIN sysScope s ON x.oidOrigin = s.oid
)
"""
        scope_cond = " AND f.oidOrigin IN (SELECT oid FROM sysScope)"
    else:
        # 未选择系统范围 = 统计全部
        sys_scope_cte = ""
        scope_cond = ""

    # 注（MSSQL 路径）：不将 XPathRunUsesSpec 换成 CoRERelationOrigin 子查询
    # ——实测嵌进本大 JOIN 后优化器执行计划劣化（>9 分钟），视图版基线
    # 31.7s 可接受。Oracle 已在函数入口分流到字典方案。
    spec_join = ("JOIN XPathRunUsesSpec x ON x.oidOrigin = f.oidOrigin\n"
                 "JOIN JDPipeSpec sd ON sd.oid = x.oidDestination")

    sql = f"""{sys_scope_cte}SELECT h.ItemName AS PipelineName, sd.SpecName AS SpecName, p.Npd AS Npd, p.NpdUnitType AS NpdUnitType,
       p.WallThicknessOrGrooveSetback AS WallThicknessM, cv.ShortStringValue AS ScheduleText,
       SUM({ceil}(sp.CutLength / ({cut_case}) - 1)) AS WeldCount, h.oid AS PipelineOid
FROM JRteStockPartOccur sp
JOIN JRtePipePart pp ON pp.oid = sp.oid
JOIN XDistribPorts b ON b.oidOrigin = sp.oid
JOIN JDPipePort p ON p.oid = b.oidDestination AND p.PortIndex = 1
JOIN XOwnsParts f ON f.oidDestination = sp.oid
{spec_join}
JOIN XSystemHierarchy g ON g.oidDestination = f.oidOrigin
JOIN JNamedItem h ON h.oid = g.oidOrigin
LEFT JOIN {_t(db, schema_db, 'CodelistValueView')} cv ON cv.ValueID = p.ScheduleThickness AND cv.TableName = 'ScheduleThickness'
LEFT JOIN JMTOInfo mto ON mto.Oid = sp.oid
WHERE sd.SpecName IN ({spec_list}) AND sp.CutLength > {_fmt_double(min(clr.values()))} AND sp.CutLength > {cut_case}{scope_cond}
AND {nvl}(mto.ReportingRequirements, 5) <> 2
GROUP BY h.ItemName, sd.SpecName, p.Npd, p.NpdUnitType, p.WallThicknessOrGrooveSetback, cv.ShortStringValue, h.oid"""
# 注：sp.CutLength > min(clr) 为常量预过滤（clr 清洗后恒为正，语义为精确 CASE 的超集），
# 引导优化器走 CutLength 索引；无它 Oracle 下 WHERE 侧连接依赖 CASE 会导致执行计划劣化（分钟级）。

    welds: Dict[WeldInfo, int] = {}
    try:
        for row in db.query(sql):
            if row[2] is None or row[4] is None:
                continue
            spec_name = str(row[1])
            info = WeldInfo(
                pipeline_id=str(row[0]),
                npd=convert_npd_value(float(row[2]), str(row[3] or "")),
                pipeline_oid=str(row[7]),
            )
            if opts.met == 2 and spec_name in opts.mg:
                info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                                materials_group=opts.mg[spec_name],
                                pipeline_oid=info.pipeline_oid)
            elif opts.met == 1:
                info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                                spec=spec_name, pipeline_oid=info.pipeline_oid)
            if not opts.thk:
                info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                                spec=info.spec, materials_group=info.materials_group,
                                schedule_thickness=str(row[5] or ""),
                                wall_thickness=wall_thickness_from_meters(float(row[4])),
                                pipeline_oid=info.pipeline_oid)
            cnt = int(round(float(row[6])))
            welds[info] = welds.get(info, 0) + cnt
    except Exception as ex:
        log(f"管子聚合查询失败: {ex}")
    return welds


def _aggregate_pipes_ora(db: Db, schema_db: str, opts: StatsOptions,
                         clr: Dict[str, float], log: Callable[[str], None],
                         mdb_schema: Optional[str]) -> Dict[WeldInfo, int]:
    """Oracle 版管子聚合：spec/代码表走字典，逐直管行 Python 端聚合。

    与 MSSQL SQL 版逐项等价：spec 字典白名单 = 视图 INNER JOIN 丢弃
    语义；CutLength > cut 常量预过滤保留（超集），精确过滤 len > cut 在
    Python 端做（CEIL(len/cut-1) ≥ 1 ⟺ len > cut）；分组键与 WeldInfo
    构造同构。run oid 侧 RAWTOHEX 大写，统一 lower 后对字典键。
    """
    import math
    spec_d = _load_run_spec(db, mdb_schema)
    log(f"规格字典：{len(spec_d)} 个 run（Oracle 直查）")
    # 壁厚代码表：跨 schema JOIN 大 SQL 有物化风险，小表直读
    cv_d: Dict[str, str] = {}
    try:
        cv_d = {str(r[0]): str(r[1] or "") for r in db.query(
            f"SELECT ValueID, ShortStringValue FROM "
            f"{_t(db, schema_db, 'CodelistValueView')} "
            f"WHERE TableName = 'ScheduleThickness'")}
    except Exception:
        cv_d = {}

    scope_cte, scope_cond = "", ""
    if opts.sys_oids:
        scope_cte = f"""WITH sysScope (oid) AS (
    SELECT h.oid FROM JNamedItem h WHERE h.oid IN ({_oid_list(db, opts.sys_oids)})
    UNION ALL
    SELECT x.oidDestination FROM XSystemHierarchy x JOIN sysScope s ON x.oidOrigin = s.oid
)
"""
        scope_cond = " AND f.oidOrigin IN (SELECT oid FROM sysScope)"

    sql = f"""{scope_cte}SELECT h.ItemName AS PipelineName,
       RAWTOHEX(f.oidOrigin) AS runOid, p.Npd AS Npd,
       p.NpdUnitType AS NpdUnitType,
       p.WallThicknessOrGrooveSetback AS WallThicknessM,
       p.ScheduleThickness AS SchedId, sp.CutLength AS CutLength,
       h.oid AS PipelineOid
FROM JRteStockPartOccur sp
JOIN JRtePipePart pp ON pp.oid = sp.oid
JOIN XDistribPorts b ON b.oidOrigin = sp.oid
JOIN JDPipePort p ON p.oid = b.oidDestination AND p.PortIndex = 1
JOIN XOwnsParts f ON f.oidDestination = sp.oid
JOIN XSystemHierarchy g ON g.oidDestination = f.oidOrigin
JOIN JNamedItem h ON h.oid = g.oidOrigin
LEFT JOIN JMTOInfo mto ON mto.Oid = sp.oid
WHERE sp.CutLength > {_fmt_double(min(clr.values()))}{scope_cond}
AND NVL(mto.ReportingRequirements, 5) <> 2"""

    welds: Dict[WeldInfo, int] = {}
    n_rows = n_hit = 0
    try:
        for row in db.query(sql):
            n_rows += 1
            if row[2] is None or row[4] is None:
                continue
            spec_name = spec_d.get(str(row[1] or "").lower())
            if spec_name not in clr:
                continue
            cutlen = float(row[6] or 0)
            cut = clr[spec_name]
            if cutlen <= cut:
                continue
            cnt = math.ceil(cutlen / cut - 1)
            n_hit += 1
            sched = cv_d.get(str(row[5]), "") if row[5] is not None else ""
            info = WeldInfo(
                pipeline_id=str(row[0]),
                npd=convert_npd_value(float(row[2]), str(row[3] or "")),
                pipeline_oid=str(row[7]),
            )
            if opts.met == 2 and spec_name in opts.mg:
                info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                                materials_group=opts.mg[spec_name],
                                pipeline_oid=info.pipeline_oid)
            elif opts.met == 1:
                info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                                spec=spec_name, pipeline_oid=info.pipeline_oid)
            if not opts.thk:
                info = WeldInfo(pipeline_id=info.pipeline_id, npd=info.npd,
                                spec=info.spec,
                                materials_group=info.materials_group,
                                schedule_thickness=sched,
                                wall_thickness=wall_thickness_from_meters(
                                    float(row[4])),
                                pipeline_oid=info.pipeline_oid)
            welds[info] = welds.get(info, 0) + cnt
    except Exception as ex:
        log(f"管子聚合查询失败: {ex}")
        return welds
    log(f"管子聚合(Oracle 字典版)：直管行 {n_rows}，命中规格 {n_hit}，"
        f"聚合 {len(welds)} 组")
    return welds


def _fmt_double(v: float) -> str:
    """双精度转 SQL 字面量（避免科学计数法/千分位）。"""
    if v == int(v):
        return str(int(v))
    return f"{v:.6f}".rstrip("0").rstrip(".")


# ------------------------------------------------------------- 规格加载 ----

def load_specs(db_c: Db, schema_db: str):
    """规格表（对应 RadDropDownList1_SelectedIndexChanged 的 CDB 查询）。

    返回 (specs, groups, group_list)：
      specs      = [(SpecName, MaterialsGroup 文本)] 全量（含无组的规格）
      group_list = 材料组下拉全列表（Schema 库 CodelistValueView）
    实际 .NET 版：specs 来自 REFDATPipeSpec 与材料组 JOIN，groups 与 specs 等长；
    group_list 单独从 Schema 库取全材料组。
    """
    sql = (f"SELECT S.SpecName, C.ShortStringValue "
           f"FROM REFDATPipeSpec S "
           f"JOIN {_t(db_c, schema_db, 'CodelistValueView')} C "
           f"ON C.ValueID = S.MaterialsGroup "
           f"WHERE C.TableName = 'MaterialsGroup'")
    pairs = []
    for row in db_c.query(sql):
        pairs.append((str(row[0]), str(row[1] or "")))
    group_list = []
    for row in db_c.query(
        f"SELECT c.ShortStringValue FROM {_t(db_c, schema_db, 'CodelistValueView')} c "
        f"WHERE c.TableName = 'MaterialsGroup'"
    ):
        group_list.append(str(row[0] or ""))
    return pairs, group_list
