# Smart3D 范例代码文档集 · 总览

> **源目录**: `C:\Program Files (x86)\Smart3D`
> **整理日期**: 2026-07-01

---

## 文档集结构

| 编号 | 文档名称 | 覆盖范围 | 核心内容 |
|------|---------|----------|---------|
| 00 | 本文（总览） | 全部 | 架构说明、通用模式、速查索引 |
| 01 | [公共应用示例](01_公共应用示例.md) | CommonApp | 客户端服务、三维几何、数学类、草图、权限组、备份恢复、站点管理 |
| 02 | [管道布线示例](02_管道布线示例.md) | CommonRoute | Pipeline/PipeRun创建、10种布线、8种特征插入、绝缘规范 |
| 03 | [空间管理示例](03_空间管理示例.md) | CommonSpace | Area/Folder/Interference/Zone 四种空间、7种构造方式 |
| 04 | [网格系统示例](04_网格系统示例.md) | GridSystem | 7种网格元素、坐标系统创建、IPlane/ISurface接口 |
| 05 | [.NET符号开发示例](05_NET符号开发示例.md) | Programming | Graphics3D工具库、符号开发模板、管道/HVAC/电缆符号 |
| 06 | [图纸与报表示例](06_图纸与报表示例.md) | Drawings/Reports | 交付物树管理、SaveAs导出、自定义查询解释器 |
| 07 | [船舶与结构示例](07_船舶与结构示例.md) | Planning/ShipDrawings等 | 板/加强筋/型材系统、装配体、块分割、规划接头 |
| 08 | [设备与支吊架示例](08_设备与支吊架示例.md) | Equipment/Supports | 设备/端口/形状创建、6种支吊架、端口约束 |
| **09** | [**开发速查手册**](09_开发速查手册.md) | **跨模块** | **通用API签名、代码模板、命名空间引用、常见陷阱** |

---

## 架构总览

Smart3D（Hexagon/Intergraph SmartPlant 3D）的二次开发 API 称为 **SOM (Symbol Object Model)**，分为客户端层和中间层：

```
┌─────────────────────────────────────────────────┐
│              应用程序 / 命令 / 窗体               │
├──────────────────┬──────────────────────────────┤
│  CommonClient    │  CommonMiddle                │
│  (客户端层)       │  (中间层/业务层)              │
│                  │                              │
│  · 命令基类       │  · MiddleServiceProvider     │
│  · WorkingSet    │  · SiteManager / Site        │
│  · SelectSet     │  · Plant / Model / Catalog   │
│  · Preferences   │  · TransactionManager        │
│  · ValueManager  │  · BusinessObject            │
│  · TransactionMgr│  · 几何/数学类                │
├──────────────────┴──────────────────────────────┤
│              各业务模块 (均基于 CommonMiddle)      │
├──────────────────────────────────────────────────┤
│ CommonRoute │ CommonSpace │ GridSystem │ Equipment│
│ Drawings    │ Planning    │ Reports    │ Supports │
│ ShipDrawings│ StructuralAnalysis │ StructMfg     │
├──────────────────────────────────────────────────┤
│         Programming (符号开发框架)                 │
│  CustomSymbolDefinition → Graphics3D → PortHelper │
└──────────────────────────────────────────────────┘
```

### 两种开发模式

| 模式 | 说明 | 继承基类 | 适用场景 |
|------|------|---------|---------|
| **SOM 业务开发** | 通过中间层 API 操作业务对象 | — | 管道布线、空间管理、设备创建、图纸管理等 |
| **符号开发** | 自定义3D零部件的几何形状 | `CustomSymbolDefinition` | 管件（法兰/三通/阀门）、HVAC组件、电气元件等 |
| **命令开发** | 自定义交互命令 | `BaseGraphicCommand` / `BaseStepCommand` / `BaseModalCommand` | 选择对象、图形交互、模态对话框 |

---

## 通用代码模板

### 1. 获取基础服务对象（几乎所有方法的起点）

```vb
' === 中间层服务（最常用） ===
Dim oTransactionMgr As TransactionManager = MiddleServiceProvider.TransactionMgr
Dim oSiteMgr As SiteManager = MiddleServiceProvider.SiteMgr
Dim oPlant As Plant = oSiteMgr.ActiveSite.ActivePlant
Dim oModel As Model = oPlant.PlantModel
Dim oCatalog As Catalog = oPlant.PlantCatalog
Dim oRootSystem As BusinessObject = DirectCast(oModel.RootSystem, BusinessObject)

' === 客户端服务 ===
Dim oWorkingSet As WorkingSet = ClientServiceProvider.WorkingSet
Dim oSelectSet As SelectSet = ClientServiceProvider.SelectSet
Dim oPreferences As Preferences = ClientServiceProvider.Preferences
Dim oValueMgr As ValueManager = ClientServiceProvider.ValueMgr
```

### 2. 事务管理模式

```vb
' 所有写操作后必须 Commit
oTransactionMgr.Commit("操作描述")

' 回滚未提交的事务
oTransactionMgr.Abort()

' Compute 模式（批量操作前暂停自动提交）
oTransactionMgr.Compute(True)   ' 进入计算模式
' ... 执行多个操作 ...
oTransactionMgr.Compute(False)  ' 退出计算模式
oTransactionMgr.Commit("批量操作")
```

### 3. 目录查询模式

```vb
' 方式A: CatalogBaseHelper（通用零件查询）
Dim oCatHelper As New CatalogBaseHelper(oCatalog)
Dim oPart As Part = DirectCast(oCatHelper.GetPart("零件编号"), Part)
Dim oPartClass As PartClass = DirectCast(oCatHelper.GetPartClass("零件类名"), PartClass)

' 方式B: CatalogStructHelper（结构截面/材料查询）
Dim oCatStrHelper As New CatalogStructHelper(oCatalog)
Dim oCrossSection As CrossSection = oCatStrHelper.GetCrossSection("AISC-LRFD-3.1", "W", "W8x10")
Dim oMaterial As Material = oCatStrHelper.GetMaterial("Steel - Carbon", "A")
```

### 4. 属性读写模式

```vb
' 通过接口名+属性名读写 BusinessObject 属性
oBO.SetPropertyValue(value, "IJNamedItem", "Name")

Dim oPV As PropertyValue = oBO.GetPropertyValue("IJNamedItem", "Name")
Dim oPVString As PropertyValueString = DirectCast(oPV, PropertyValueString)
Dim sName As String = DirectCast(oPVString.PropValue, String)

' 用户自定义名称（快捷方式）
oBO.SetUserDefinedName("MyName")
```

### 5. 矩阵变换链

```vb
Dim m As New Matrix4X4()
m.SetIdentity()
m.Rotate(angleInRadians, axisVector)   ' 先旋转
m.Scale(scaleFactor)                   ' 再缩放
m.Translate(translationVector)         ' 最后平移

' 应用到几何对象
oGeometry.Transform(m)

' 应用到位置点（含平移）
Dim newPos As Position = m.Transform(oldPos)

' 应用到向量（仅旋转+缩放，不含平移）
Dim newVec As Vector = m.Transform(oldVec)
```

---

## 示例总览表

| # | 模块 | 示例名称 | 语言 | 核心功能 |
|---|------|---------|------|---------|
| 1 | CommonApp | 客户端服务 | VB.NET | WorkingSet、SelectSet、Preferences、ValueManager、事务管理、图形命令 |
| 2 | CommonApp | 三维几何体(7种) | VB.NET | Arc3d、Circle3d、Cone3d、Line3d、Plane3d、Sphere3d、Torus3d |
| 3 | CommonApp | 数学类(4种) | VB.NET | Matrix4x4、Position、RangeBox、Vector |
| 4 | CommonApp | 三维草图 | VB.NET | Sketch3D — 点、线/弧/椭圆弧特征、倒角/弯折特征 |
| 5 | CommonApp | 权限组管理 | VB.NET | 从Excel批量创建权限组和用户访问规则 |
| 6 | CommonApp | 工厂备份/恢复 | VB.NET | BackupPlant、RestoreSite/RestorePlant、MigrateDatabase |
| 7 | CommonApp | 站点管理器 | C#/VB | 站点连接、工厂打开、连接信息对话框 |
| 8 | CommonApp | 符号几何辅助 | VB.NET | 模态命令启动符号几何编辑对话框 |
| 9 | CommonRoute | 管道段创建 | VB.NET | Pipeline/PipeRun创建、10种布线方式 |
| 10 | CommonRoute | 特征插入(8种) | VB.NET | ShortCode/Part/Tag方式插入，Split特征 |
| 11 | CommonRoute | 管道特征/零件 | VB.NET | 绝缘规范操作、BasePart转换 |
| 12 | CommonSpace | 区域/干涉/分区 | VB.NET | 7种构造方式（两点/四点/平面偏移/基元/合并/边界/路径） |
| 13 | CommonSpace | 空间文件夹 | VB.NET | 多级层级管理、GetAllSpaces递归查询 |
| 14 | GridSystem | 网格轴/平面 | VB.NET | 5种轴类型、AddPlane/AddPlanes批量创建 |
| 15 | GridSystem | 网格圆柱/弧/线 | VB.NET | 交叉元素生成、IPlane/ISurface接口方法 |
| 16 | Programming | Graphics3D工具库 | C# | ~150个静态方法：曲线/曲面/实体创建 |
| 17 | Programming | 管道符号(5种) | C# | Tee(33种PDB)、Flange、Clamp(22种PDB)、Plug、OPLever |
| 18 | Programming | HVAC符号(2种) | C# | HTee（矩形/圆形/扁椭圆）、Wye（7种变体） |
| 19 | Programming | 电缆规则 | C# | NEC标准电缆填充率计算 |
| 20 | Drawings | 图纸管理应用 | C# | 交付物树、属性/版本/发布、SaveAs导出 |
| 21 | Reports | 查询解释器(3种) | C# | 标签属性查询、权限组路径/位置查询 |
| 22 | Planning | 规划工作流 | VB.NET | 板/加强筋/型材系统、装配体、块分割/合并、接头 |
| 23 | ShipDrawings | 自定义属性/规则 | C# | Level层级计算、端切符号XML规则匹配 |
| 24 | Equipment | 设备创建(4种) | VB.NET | PartNumber/Part/PartClass/ParentSystem方式 |
| 25 | Equipment | 端口创建(6种) | VB.NET | Pipe/HVAC/Cable/CableTray/Conduit/Foundation端口 |
| 26 | HangersAndSupports | 支吊架(6种) | VB.NET | Pipe/Conduit/CableTray/Duct/Combined/Design支撑 |
