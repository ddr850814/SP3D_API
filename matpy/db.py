# -*- coding: utf-8 -*-
"""数据访问层：注册表读 Site 连接、DPAPI 凭据加密存储、双后端查询、Plant 枚举。

不依赖 SP3D API。双后端：SQL Server = pyodbc + ODBC Driver 18；Oracle =
oracledb thin（免客户端），每库独立连接并 ALTER SESSION SET CURRENT_SCHEMA
——等价 SQL Server 的 DATABASE= 语义。oid 列字节序（Oracle RAW(16) 大端
原序 / SQL Server bytes_le）与转换细节见 skill「SP3D 库角色与通用约定」。
后端类型自动取注册表 SiteDatabaseProviderType（MSSQL/ORACLE），认证 ini
按后端读各自的注册表值（Sql/OraDBAuthenticationFileLocation）。
"""
import base64
import ctypes
import json
import os
import uuid
from ctypes import wintypes
from pathlib import Path

import pyodbc

import s3dcrypt

# ---------------------------------------------------------------- DPAPI ----
# 用 ctypes 直接调 CryptProtectData/CryptUnprotectData，避免 pywin32 依赖
# （Python 3.14 下 pywin32 尚未跟进，且本工具只需这两个函数）。

class _DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _dpapi_protect(data: bytes) -> bytes:
    blob_in = _DATA_BLOB(len(data), ctypes.cast(ctypes.create_string_buffer(data), ctypes.POINTER(ctypes.c_byte)))
    blob_out = _DATA_BLOB()
    ok = ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out))
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def _dpapi_unprotect(data: bytes) -> bytes:
    blob_in = _DATA_BLOB(len(data), ctypes.cast(ctypes.create_string_buffer(data), ctypes.POINTER(ctypes.c_byte)))
    blob_out = _DATA_BLOB()
    ok = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out))
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def encrypt_password(plain: str) -> str:
    """DPAPI 加密密码，返回 base64 文本（仅当前 Windows 用户可解）。"""
    if not plain:
        return ""
    return base64.b64encode(_dpapi_protect(plain.encode("utf-8"))).decode("ascii")


def decrypt_password(enc: str) -> str:
    if not enc:
        return ""
    try:
        return _dpapi_unprotect(base64.b64decode(enc)).decode("utf-8")
    except Exception:
        return ""


# ------------------------------------------------------------ 注册表 Site ----

_REG_KEY = r"Software\Intergraph\Applications\Environments\CommonApp\ProjectDB"

# 连接串键（PhysicalName/Schema 形如 SERVER=shser029;DATABASE=shser029_SDB）
_REG_KEYS = {
    "PhysicalName": "site",     # Site 库（PRJMGTDatabase 所在）
    "Schema": "schema",         # Site Schema 库
}


# 后端类型注册表值（MSSQL/ORACLE）
_REG_PROVIDER_VALUE = "SiteDatabaseProviderType"


def read_site_connection_from_registry():
    r"""读注册表 ProjectDB，返回 {site, schema, provider}。

    site/schema 各为 {server, db} 或缺省；provider 为后端类型
    （MSSQL/ORACLE，大写）或缺省。

    实测键：
      HKCU\Software\Intergraph\Applications\Environments\CommonApp\ProjectDB
        PhysicalName = SERVER=S3D2016;DATABASE=SHSER022_SDB
        Schema       = SERVER=S3D2016;DATABASE=SHSER022_SDB_SCHEMA
        SiteDatabaseProviderType = ORACLE
    """
    import winreg
    result = {}
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as key:
            for value_name, slot in _REG_KEYS.items():
                try:
                    raw, _ = winreg.QueryValueEx(key, value_name)
                except FileNotFoundError:
                    continue
                parsed = _parse_conn_str(raw or "")
                if parsed:
                    result[slot] = parsed
            try:
                prov, _ = winreg.QueryValueEx(key, _REG_PROVIDER_VALUE)
                prov = (prov or "").strip().upper()
                if prov:
                    result["provider"] = prov
            except FileNotFoundError:
                pass
    except OSError:
        return result
    return result


def _parse_conn_str(s: str):
    """解析 'SERVER=shser029;DATABASE=shser029_SDB' -> {server, db}。"""
    parts = {}
    for seg in s.split(";"):
        seg = seg.strip()
        if not seg or "=" not in seg:
            continue
        k, _, v = seg.partition("=")
        parts[k.strip().upper()] = v.strip()
    if not parts.get("DATABASE"):
        return None
    return {"server": parts.get("SERVER", ""), "db": parts["DATABASE"]}


# -------------------------------------------------------- 注册表认证 ini ----
# S3D 的数据库账号密码不存注册表明文，而是：
#   注册表值 SqlDBAuthenticationFileLocation 记录 ini 文件绝对路径，
#   ini 由 ServicesAuthenticationSettings.exe 维护，
#   内容为 '<ServerName> <DatabaseUser> <EncryptedPassword>'（空格分隔三列），
#   密码用 S3dCryptServices（AES-256-CBC，见 s3dcrypt.py）加密。
# 这里实现：注册表 -> ini -> 密文 -> 解密，全程无需用户输入。

# 两后端的认证 ini 各存一个注册表值，可并存（SQL Server 站点用 Sql 值、
# Oracle 站点用 Ora 值），按当前后端取
_AUTH_REG_VALUES = {
    "MSSQL": "SqlDBAuthenticationFileLocation",
    "ORACLE": "OraDBAuthenticationFileLocation",
}


def read_auth_ini_path_from_registry(provider="MSSQL"):
    """读注册表 ProjectDB 键的认证 ini 绝对路径（按后端取 Sql/Ora 值）。"""
    import winreg
    value_name = _AUTH_REG_VALUES.get((provider or "MSSQL").upper(),
                                      _AUTH_REG_VALUES["MSSQL"])
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY) as key:
            raw, _ = winreg.QueryValueEx(key, value_name)
        return (raw or "").strip() or None
    except OSError:
        return None


def parse_auth_ini(path) -> list:
    """解析 S3D 认证 ini，返回 [{server, user, password_enc, provider}, ...]。

    格式为若干段，每段两行：'Database_Type SQL_Server'（或 Oracle）声明
    类型 + '<ServerName> <DatabaseUser> <EncryptedPassword>'（空格分隔
    三列）。跳过 '#' 注释行；条目带 provider 标记，调用方按后端过滤；
    无类型声明按 MSSQL 处理（兼容老格式）。
    """
    result = []
    provider = "MSSQL"
    try:
        lines = Path(path).read_text(encoding="utf-8", errors="ignore").splitlines()
    except OSError:
        return result
    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split()
        if len(parts) == 2 and parts[0].lower() == "database_type":
            provider = "ORACLE" if parts[1].lower().startswith("oracle") else "MSSQL"
            continue
        if len(parts) == 3:
            result.append({"server": parts[0], "user": parts[1],
                           "password_enc": parts[2], "provider": provider})
    return result


def get_credentials_from_ini(path, server=None, provider="MSSQL"):
    """解析指定认证 ini 并解密，返回 (server, username, password)；失败返回 None。

    server 非空时优先匹配服务器名对应的条目（大小写/端口/实例名容错），
    匹配不到才回退第一条，保证连接目标与凭据同属一台数据库服务器。
    仅取 provider 匹配的条目（SQL Server 与 Oracle 认证文件互相独立）。
    """
    if not path:
        return None
    want = (provider or "MSSQL").upper()
    entries = [e for e in parse_auth_ini(path)
               if e.get("provider", "MSSQL") == want]
    if not entries:
        return None

    def _pick(e):
        password = s3dcrypt.s3d_decrypt(e["password_enc"])
        if not password:
            return None
        return e["server"], e["user"], password

    if server:
        for e in entries:
            if _server_matches(server, e["server"]):
                got = _pick(e)
                if got:
                    return got
    return _pick(entries[0])


def _server_matches(a: str, b: str) -> bool:
    """服务器名归一化比较：忽略大小写、端口（,1433）、实例名（\\INST）、域名后缀。"""

    def _norm(s):
        s = (s or "").strip().lower()
        return s.split(",")[0].split("\\")[0]

    na, nb = _norm(a), _norm(b)
    if not na or not nb:
        return False
    return na == nb or na.endswith("." + nb) or nb.endswith("." + na)


def get_credentials_from_registry(server=None, provider="MSSQL"):
    """注册表 -> ini -> 解密，返回 (server, username, password) 明文；任一步失败返回 None。

    server 非空时按服务器名在 ini 中匹配对应条目；ini 路径与条目都按
    当前后端（provider）取。
    """
    return get_credentials_from_ini(
        read_auth_ini_path_from_registry(provider), server, provider)


# ---------------------------------------------------------------- 配置存储 ----

CONFIG_PATH = Path(os.environ.get("APPDATA", Path.home())) / "MatMat" / "config.json"


def load_config() -> dict:
    """读本地配置（不含明文密码）。文件损坏时返回空配置。"""
    default = {
        "server": "",
        "site_db": "",
        "schema_db": "",
        "provider": "MSSQL",  # 后端类型（MSSQL/ORACLE），始终跟随注册表
        "username": "",
        "password_enc": "",
        "last_plant": "",
        "last_dir": "",
        "by_line": False,  # 「按管线号分组」勾选状态（重启后恢复）
        "hidden_only": False,  # 「材料范围」下拉：False=排除隐藏 / True=仅隐藏
        "report": "管道材料",  # 当前报表（管道材料/螺栓/垫片/焊接量统计…）
        "met": 0,  # 焊接量分组方式：0=不分组 / 1=按材料等级 / 2=按材料组
        "thk": False,  # 焊接量「忽略壁厚」勾选状态
        "skip_pipes": False,  # 焊接量「忽略直管焊缝」勾选状态
        "plant_filters": {},  # 按项目记忆的过滤器选择 {plant: [过滤器标签, …]}
        "auth_ini": "",   # 手动选择的认证 ini 路径（注册表查不到时兜底）
        # 以下键同样会被 _on_close/_mem_* 写入，必须进白名单，
        # 否则 load 的键过滤会把它们丢弃（记忆重启失效）
        "geo_mat": "",    # 材料类报表窗口尺寸记忆 "宽x高"
        "geo_pid": "",    # PID 关联报表窗口尺寸记忆
        "geo_weld": "",   # 焊接量报表窗口尺寸记忆
        "pid_kinds": None,  # PID 关联对象类别勾选（list 或 None=全选）
        "plant_dwg_filters": {},  # 按项目记忆的图纸目录勾选 {plant: {names, oids}}
    }
    try:
        if CONFIG_PATH.exists():
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            default.update({k: v for k, v in data.items() if k in default})
    except Exception:
        pass
    # 自动从认证 ini 提取账号密码（替代手动输入）：
    # 优先注册表记录的位置；注册表查不到时用上次手动选择的 ini。
    # 服务器名与 Site 库必须同源：Site 库来自注册表时，服务器也取注册表 Site 的服务器
    # （否则会用 ini 第一行的服务器，而 Site 库却是注册表环境的，导致 42S02）。
    try:
        reg = read_site_connection_from_registry()
        site_reg = reg.get("site") or {}
        reg_server = (site_reg.get("server") or "").strip()
        reg_db = (site_reg.get("db") or "").strip()
        # 后端类型与 Site 库/服务器均以注册表（SP3D 当前环境）为准，存在即
        # 覆盖本地残留——无论是后端切换还是同后端跨站点切换都立即跟随
        if reg.get("provider"):
            default["provider"] = reg["provider"]
        if reg_db:
            default["site_db"] = reg_db
            if reg_server:
                default["server"] = reg_server
        provider = default["provider"]
        target = default.get("server") or reg_server or ""
        cred = get_credentials_from_registry(target, provider)
        if not cred and default.get("auth_ini"):
            cred = get_credentials_from_ini(default["auth_ini"], target, provider)
        if cred:
            ini_server, _user, _pwd = cred
            if _user:
                default["username"] = _user
            if _pwd:
                default["password_enc"] = encrypt_password(_pwd)
            if not default.get("server") and ini_server:
                default["server"] = ini_server
    except Exception:
        pass
    return default


def _atomic_write_json(path: Path, obj) -> None:
    """先写临时文件再原子替换：写入中途断电/杀进程不会截断原文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    os.replace(tmp, path)


def save_config(cfg: dict) -> None:
    _atomic_write_json(CONFIG_PATH, cfg)


SPEC_MEM_PATH = Path(os.environ.get("APPDATA", Path.home())) / "MatMat" / "spec_mem.json"
# 旧 weldpy 的记忆文件：MatMat 下还没有时回退读取（焊接量功能从 weldpy 迁移）
_LEGACY_SPEC_MEM_PATH = (Path(os.environ.get("APPDATA", Path.home()))
                         / "WeldWeld" / "spec_mem.json")


def load_spec_mem() -> dict:
    """读焊接量规格表记忆 {plant: {rows, deleted, filter}}。

    MatMat 下无文件时回退旧 WeldWeld 文件；首次保存起落到 MatMat。
    """
    for p in (SPEC_MEM_PATH, _LEGACY_SPEC_MEM_PATH):
        try:
            if p.exists():
                data = json.loads(p.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return data
        except (OSError, ValueError):
            pass
    return {}


def save_spec_mem(mem: dict) -> None:
    _atomic_write_json(SPEC_MEM_PATH, mem)


# -------------------------------------------------------------- pyodbc 查询 ----

DRIVER = "ODBC Driver 18 for SQL Server"


def _odbc_quote(v: str) -> str:
    """ODBC 值用花括号包裹（内部 } 双写转义），防止密码含 ;={ 等字符时连接串被截断。"""
    return "{" + v.replace("}", "}}") + "}"


def connect_string(server: str, db: str, username: str, password: str) -> str:
    """拼 pyodbc 连接串（ODBC 18 需 TrustServerCertificate，否则自签名证书报错）。

    服务器名后加 ",1433" 强制走 TCP/IP：ODBC 客户端默认尝试 Named Pipes(445)，
    若 445 不可达会报 Named Pipes Provider [53]；显式端口可跳过协议探测。
    """
    target = server
    if target and "\\" not in target and "," not in target:
        target = f"{target},1433"
    return (
        f"DRIVER={{{DRIVER}}};"
        f"SERVER={target};DATABASE={db};"
        f"UID={_odbc_quote(username)};PWD={_odbc_quote(password)};"
        f"TrustServerCertificate=yes;"
    )


def open_db(provider: str, server: str, db: str, username: str, password: str) -> "Db":
    """按后端类型建 Db：ORACLE 走 oracledb thin（db 参数即 schema），否则 pyodbc。"""
    if (provider or "").upper() == "ORACLE":
        return Db.oracle(server, db, username, password)
    return Db(connect_string(server, db, username, password))


def _oracle_config_dir():
    """tns 别名解析所需的配置目录（含 tnsnames.ora）；找不到返回 None。

    依次查 TNS_ADMIN 环境变量、本机 Oracle 客户端注册表（含 32 位
    WOW6432Node）各 KEY_* 主目录的 ORACLE_HOME\\network\\admin。
    EZConnect 描述符（含 : / ( 字符）不走别名解析，无需此目录。
    """
    env = os.environ.get("TNS_ADMIN", "").strip()
    if env and Path(env, "tnsnames.ora").exists():
        return env
    try:
        import winreg
    except ImportError:
        return None
    for sub in ("SOFTWARE\\ORACLE", "SOFTWARE\\WOW6432Node\\ORACLE"):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub) as key:
                i = 0
                while True:
                    try:
                        home_name = winreg.EnumKey(key, i)
                    except OSError:
                        break
                    i += 1
                    if not home_name.upper().startswith("KEY_"):
                        continue
                    try:
                        with winreg.OpenKey(key, home_name) as hk:
                            home, _ = winreg.QueryValueEx(hk, "ORACLE_HOME")
                    except OSError:
                        continue
                    cand = str(Path(home, "network", "admin"))
                    if Path(cand, "tnsnames.ora").exists():
                        return cand
        except OSError:
            continue
    return None


class Db:
    """一次数据库会话：持连接参数缓存，供后续 query 复用（少建连）。

    MSSQL：pyodbc 连接串（DATABASE= 选库）；ORACLE：oracledb thin 免客户端，
    连接后 ALTER SESSION SET CURRENT_SCHEMA 选 schema（等价 DATABASE= 语义）。
    """

    def __init__(self, conn_str: str):
        self.provider = "mssql"
        self._conn_str = conn_str
        self._ora = None
        self._conn = None

    @classmethod
    def oracle(cls, dsn: str, schema: str, username: str, password: str) -> "Db":
        self = cls.__new__(cls)
        self.provider = "oracle"
        self._conn_str = ""
        self._ora = (username, password, dsn, schema)
        self._conn = None
        return self

    def connect(self):
        if self._conn is None:
            if self.provider == "oracle":
                import oracledb
                oracledb.defaults.fetch_lobs = False  # CLOB 列直接取文本
                user, pwd, dsn, schema = self._ora
                kwargs = {}
                if dsn and not any(ch in dsn for ch in ":/("):
                    cfg = _oracle_config_dir()  # 纯 tns 别名需指定配置目录
                    if cfg:
                        kwargs["config_dir"] = cfg
                self._conn = oracledb.connect(user=user, password=pwd,
                                              dsn=dsn, **kwargs)
                if schema:
                    cur = self._conn.cursor()
                    cur.execute(f"ALTER SESSION SET CURRENT_SCHEMA={schema}")
                    cur.close()
            else:
                self._conn = pyodbc.connect(self._conn_str, timeout=15)
                # 查询超时（秒）：防跨库 JOIN 坏计划时无限挂起（实测 CDB
                # JDProcurementDataInfo 无索引 JOIN 可挂死；统计类查询
                # Oracle 慢视图可达 ~100s，给 300s 余量）
                self._conn.timeout = 300
        return self._conn

    def close(self):
        if self._conn is not None:
            try:
                self._conn.close()
            finally:
                self._conn = None

    def query(self, sql: str):
        """执行 SELECT，返回行列表（列名在 row.cursor_description）。

        将 uniqueidentifier 二进制统一转成小写 hex 字符串（uuid 标准显示形式），
        其余类型原样返回（decimal->float、None 保留）。
        """
        conn = self.connect()
        cur = conn.cursor()
        if self.provider == "oracle":
            # oracledb 默认 arraysize=100：万行结果要上百次网络往返，
            # RTT 高（VPN/跨机房）时取数段会成倍放大
            cur.arraysize = 10000
        cur.execute(sql)
        # MSSQL 的 uniqueidentifier 物理序是 bytes_le（重排还原 GUID 标准序）；
        # Oracle 的 RAW(16) 是大端原序（= RAWTOHEX 直出），重排会得到错字节——
        # 曾因未分流导致 oid 全部错序、字面量比较恒 0 行
        conv = _conv_oracle if self.provider == "oracle" else _conv
        rows = []
        for row in cur.fetchall():
            rows.append([conv(v) for v in row])
        cur.close()
        return rows


def _conv(v):
    if isinstance(v, bytes):
        try:
            return uuid.UUID(bytes_le=v).hex
        except Exception:
            return v.hex()
    if isinstance(v, (int, float, str, bool)) or v is None:
        return v
    return str(v)


def _conv_oracle(v):
    """Oracle 行值转换：RAW(16) 取 plain 大端 hex（SP3D GUID 字面量原序）。"""
    if isinstance(v, bytes):
        return v.hex()
    if isinstance(v, (int, float, str, bool)) or v is None:
        return v
    return str(v)


def test_connection(server: str, db: str, username: str, password: str,
                    provider: str = "MSSQL") -> str:
    """测试连接，返回空串表示成功，否则返回错误信息。"""
    try:
        dbh = open_db(provider, server, db, username, password)
        dbh.connect()
        dbh.close()
        return ""
    except Exception as ex:
        return str(ex)


# ---------------------------------------------------------------- 诊断信息 ----

def collect_diagnostics(err: str = "") -> str:
    """收集自动连接失败时的诊断信息（不含明文密码），供打包分发后远程排查。

    输出：注册表 Site/Schema、认证 ini 路径与解析结果、生效配置、
    ODBC 驱动列表、目标服务器与原始报错。
    """
    reg = read_site_connection_from_registry()
    site = reg.get("site") or {}
    schema = reg.get("schema") or {}
    cfg = load_config()
    prov = cfg.get("provider") or "MSSQL"
    ini_path = read_auth_ini_path_from_registry(prov)
    target = cfg.get("server") or site.get("server", "")
    cred_reg = get_credentials_from_registry(target, prov)
    cred_manual = get_credentials_from_ini(cfg.get("auth_ini"), target, prov) if cfg.get("auth_ini") else None

    def _fmt_cred(cred):
        if not cred:
            return "未取到"
        return f"服务器={cred[0]}, 账号={cred[1]}, 密码解密={'成功' if cred[2] else '失败/为空'}"

    try:
        drivers = ", ".join(pyodbc.drivers()) or "(无)"
    except Exception:
        drivers = "(读取失败)"
    try:
        import oracledb
        ora_drv = f"python-oracledb {getattr(oracledb, '__version__', '?')}（thin）"
    except ImportError:
        ora_drv = "(python-oracledb 未安装)"

    lines = [
        "==== 自动连接失败诊断 ====",
        "",
        "【注册表 ProjectDB】",
        f"  Site     : {site.get('server', '')} / {site.get('db', '')}",
        f"  Schema   : {schema.get('server', '')} / {schema.get('db', '')}",
        f"  认证ini路径: {ini_path or '(未读到)'}",
        "",
        "【凭据来源】",
        f"  注册表 ini : {_fmt_cred(cred_reg)}",
        f"  手动 ini   : {_fmt_cred(cred_manual)}",
        f"  生效账号   : {cfg.get('username', '') or '(空)'}",
        f"  生效密码   : {'有' if cfg.get('password_enc') else '(空)'}",
        "",
        "【生效配置】",
        f"  后端   : {prov}",
        f"  服务器 : {cfg.get('server', '') or '(空)'}",
        f"  Site库 : {cfg.get('site_db', '') or '(空)'}",
        f"  认证ini: {cfg.get('auth_ini', '') or '(未手动指定)'}",
        "",
        "【ODBC 驱动】",
        f"  {drivers}",
        "",
        "【错误信息】",
        f"  {err or '(无)'}",
        "",
        "提示：检查【凭据来源】中服务器是否与【生效配置】中服务器一致；",
        "检查【注册表 ini / 手动 ini】是否解析出账号；MSSQL 需 ODBC Driver 18，",
        "Oracle 需已安装 python-oracledb 且网络可达监听。",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------- Plant 枚举 ----


def list_plants(site_server: str, site_db: str, username: str, password: str,
                provider: str = "MSSQL"):
    """枚举 Site 库中的 Plant，返回 [{code, display, catalog, model, reports, schema}]。

    库归属取 SPF 登记关系边（XDatabaseFor{Catalog,Model,Reports} 视图：
    ProjectRoot --边--> 库行），库名改名也能对上（如 P200029 -> GD19_CDB）；
    视图缺失时回退 Catalog 行 gid = root gid + 1，仍无法归属的库行跳过。
    Oracle 站点 site_db 即 Site schema，库名段为各项目 schema 名。
    """
    is_ora = (provider or "").upper() == "ORACLE"
    pfx = "" if is_ora else "dbo."                       # Oracle 无 dbo. 前缀
    schema_col = '"SCHEMA"' if is_ora else "PMSchema"    # 库行 schema 列名两后端不同
    db = open_db(provider, site_server, site_db, username, password)
    try:
        # 项目根：PermissionGroupID -> 项目号（回退用）
        roots = {}
        try:
            for r in db.query(f"SELECT Name, PermissionGroupID FROM {pfx}PRJMGTProjectRoot"):
                if r[0] and r[1] is not None:
                    roots[r[1]] = r[0]
        except Exception:
            pass

        # 库归属（权威）：SPF 登记关系边 ProjectRoot --边--> 库行
        edge = {}  # PRJMGTDatabase.oid -> 项目号
        try:
            for view in ("XDatabaseForCatalog", "XDatabaseForModel", "XDatabaseForReports"):
                for r in db.query(
                        f"SELECT x.oidDestination, p.Name FROM {pfx}{view} x "
                        f"JOIN {pfx}PRJMGTProjectRoot p ON p.oid = x.oidOrigin"):
                    if r[0] and r[1]:
                        edge[r[0]] = r[1]
        except Exception:
            edge = {}  # 视图缺失（旧版 SDB）：整体退回 gid 推断

        rows = db.query(f"SELECT oid, Name, Path, {schema_col}, PermissionGroupID "
                        f"FROM {pfx}PRJMGTDatabase")
        # 项目号 -> 分桶
        plants = {}
        order = []  # (code, Name, dbname, schema_db)，用于确定 schema
        for row in rows:
            name = row[1] or ""
            path = _parse_conn_str(row[2] or "")
            pm_schema = _parse_conn_str(row[3] or "")
            gid = row[4] if len(row) > 4 else None
            if not path:
                continue
            code = edge.get(row[0])
            if code is None and name == "Catalog" and gid is not None and (gid - 1) in roots:
                code = roots[gid - 1]  # 回退：gid = root gid + 1
            if code is None:
                continue  # 无权威归属：跳过
            order.append((code, name, path["db"], pm_schema["db"] if pm_schema else ""))
            bucket = plants.setdefault(code, {"code": code, "display": code,
                                              "catalog": "", "model": "",
                                              "reports": "", "schema": ""})
            dbname = path["db"]
            if name == "Catalog":
                bucket["catalog"] = dbname
            elif name == "Model":
                bucket["model"] = dbname
            elif name == "Reports":
                bucket["reports"] = dbname
        # schema：Catalog 行 PMSchema -> Model 行 -> 任意行（防顺序抖动误取 _RDB_SCHEMA）
        for code, bucket in plants.items():
            bucket["schema"] = next(
                (pm for c, n, dbn, pm in order if c == code and n == "Catalog" and pm),
                next((pm for c, n, dbn, pm in order if c == code and n == "Model" and pm),
                     next((pm for c, n, dbn, pm in order if c == code and pm), "")))
        # 项目显示名：PRJMGTProjectRoot（Name = 项目号）
        proot = {}
        try:
            for r in db.query(f"SELECT Name, Description FROM {pfx}PRJMGTProjectRoot"):
                proot[r[0] or ""] = r[1] or ""
        except Exception:
            pass
        result = []
        for code, bucket in plants.items():
            if not bucket["reports"]:
                # 无 RDB 无法统计焊口/管子，排除
                continue
            desc = proot.get(code, "")
            bucket["display"] = desc or code
            result.append(bucket)
        result.sort(key=lambda p: p["code"])
        return result
    finally:
        db.close()


if __name__ == "__main__":
    # 自检：注册表读取 + DPAPI 加解密 + 注册表认证 ini 自动提取
    info = read_site_connection_from_registry()
    print("注册表 Site 连接:", info)
    enc = encrypt_password("Test123!")
    print("DPAPI 加解密一致:", decrypt_password(enc) == "Test123!")
    provider = info.get("provider") or "MSSQL"
    ini_path = read_auth_ini_path_from_registry(provider)
    print("认证 ini 路径:", ini_path)
    cred = get_credentials_from_registry(provider=provider)
    print("注册表认证提取:", cred)
