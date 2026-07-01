# 05 · .NET 符号开发示例 (Programming)

> **路径**: `C:\Program Files (x86)\Smart3D\Programming\ExampleCode\`

---

## 符号开发架构总览

```
CustomSymbolDefinition (基类)
  ├── [特性] CacheOption, VariableOutputs, SymbolVersion
  ├── [输入] InputCatalogPart, InputDouble
  ├── [方面] Aspect + SymbolOutput → AspectDefinition
  └── ConstructOutputs() → switch(PDB) → RunPDB***()
      ├── Graphics3D.Create*** → 几何体
      ├── PortHelper.Create*** → 端口
      └── m_Physical.Outputs.Add("name", geometry)
```

所有自定义符号均继承自 `CustomSymbolDefinition` 基类，通过特性（Attribute）声明缓存策略、版本、输入参数与方面输出，并在重写的 `ConstructOutputs()` 方法中根据 `PartDataBasis`（PDB）分发到具体的几何构建方法。`Graphics3D` 与 `PortHelper` 是符号开发最核心的两个工具类，前者负责创建几何体，后者负责创建端口。

---

## 一、Graphics3D — 核心几何创建工具库

> **文件**: `CommonSymbolFunctions\Graphics3D.cs`（约 5800 行，约 150 个方法）

`Graphics3D` 是 Smart3D .NET 符号开发中最核心的静态工具类，封装了底层 COM 几何对象的创建逻辑，提供面向 .NET 的简化 API。所有自定义符号的几何创建几乎都依赖此类。

### 1.1 曲线创建方法

| 方法 | 签名 | 返回类型 |
|------|------|----------|
| `CreateLine` | `(Position start, Position end)` | `Line3d` |
| `CreateLine` | `(Position root, Vector dir)`（无限长直线） | `Line3d` |
| `CreateLine` | `(Position start, Vector dir, double length)` | `Line3d` |
| `CreateArc` | `(Position start, Position mid, Position end)` | `Arc3d` |
| `CreateArc` | `(Position center, Vector normal, Position start, Position end)` | `Arc3d` |
| `CreateCircle` | `(Position center, Vector normal, double radius)` | `Circle3d` |
| `CreateCircle` | `(Position start, Position mid, Position end)` | `Circle3d` |
| `CreateBSplineCurve` | `(int order, Collection<Position> points)` | `BSplineCurve3d` |
| `CreateLineString` | `(Collection<Position> points)` | `LineString3d` |
| `CreateComplexString` | `(Collection<ICurve> curves)` | `ComplexString3d` |

### 1.2 曲面 / 实体创建方法

| 方法 | 签名 | 返回类型 |
|------|------|----------|
| `CreateProjection` | `(ICurve curve, Vector direction, double length, bool isCapped)` | `Projection3d` |
| `CreateRuledSurface` | `(ICurve baseCurve, ICurve topCurve, bool isCapped)` | `Ruled3d` |
| `CreateRevolution` | `(ICurve curve, Vector axis, Position center, double sweepAngle, bool isCapped)` | `Revolution3d` |
| `CreateCone` | `(Position baseCenter, Position topCenter, double radius1, double radius2, bool isCapped)` | `Cone3d` |
| `CreateSphere` | `(Position center, double radius, bool isSolid)` | `Sphere3d` |
| `CreateTorus` | `(Position center, Vector axis, Vector origin, double majorRadius, double minorRadius, bool isSolid)` | `Torus3d` |

### 1.3 高级封装方法（最常用）

| 方法 | 签名 | 返回类型 / 内部实现 |
|------|------|---------------------|
| `CreateCylinder` | `(Position start, Position end, double diameter, bool isCapped)` | `Projection3d`（内部：`CreateCircle(start, normal, diameter/2)` + `CreateProjection`） |
| `CreateBox` | `(Position start, Position end)` | `Projection3d` |
| `CreatePolygon` | `(ref SymbolGeometryHelper, double widthAcrossFlats, double depth, int sides)` | `Projection3d` |
| `CreateMaintenanceVolumeForValve` | `(PipeComponent, length1, length2, height1, height2, width1, width2, offset)` | `Projection3d` |
| `CreateMaintenanceVolumeForActuator` | `(...)` | `Projection3d` |

### 1.4 InteropHelper

`InteropHelper` 用于与底层 COM 交互，提供高级扫掠曲面创建能力。`GetProjectedSurfaces` 是最常用的方法，可沿投影路径从截面创建扫掠曲面，支持镜像、旋转、多种扫掠选项。

**方法签名与用法**：

```csharp
Collection<ISurface> surfaces = InteropHelper.GetProjectedSurfaces(
    connection, crossSection, projectionPath,
    cardinalPoint, mirrorOption, rotationAngle,
    startPositionNormal, endPositionNormal,
    sweepOptions, sweepOrientation);
```

---

## 二、符号开发完整模板（Tee.cs 为例）

本节以管道三通 `Tee.cs` 为例，展示一个完整符号类的开发骨架，涵盖类声明、输入参数、方面输出、PDB 分发与几何构建全流程。

### 2.1 类骨架（每个符号必须遵循）

```csharp
[CacheOption(CacheOptionType.Cached)]
[VariableOutputs]
[SymbolVersion("1.0.0.3")]
public class Tee : CustomSymbolDefinition
{
    public Tee() : base()
    {
        Ingr.SP3D.Content.CommonSymbolFunctionLoader.Load();  // 必须调用
    }
    // ...
}
```

**要点说明**：
- `[CacheOption]` 指定缓存策略，避免重复计算，提升性能。
- `[VariableOutputs]` 表示符号输出可变（端口数量/位置随参数变化）。
- `[SymbolVersion]` 标注符号版本号，用于升级兼容性管理。
- 构造函数中**必须**调用 `CommonSymbolFunctionLoader.Load()`，否则 `Graphics3D` 等工具类不可用。

### 2.2 输入参数定义

```csharp
[InputCatalogPart(1)]
public InputCatalogPart partFclt;  // 目录零件（必须）

[InputDouble(2, "FacetoCenter", "Face to Center", 0, true)]  // (序号, 属性名, 显示名, 默认值, 可编辑)
public InputDouble m_FacetoCenter;

[InputDouble(3, "InsulationThickness", "Insulation Thickness", 0.025)]  // 无 true = 不可编辑
public InputDouble m_InsulationThickness;
```

**要点说明**：
- `InputCatalogPart` 是符号的目录零件输入，**每个符号必须至少有一个**。
- `InputDouble` 第 5 个参数 `true` 表示该参数在属性窗口中可由用户编辑；省略则不可编辑。
- 序号（第一个参数）用于参数排序，必须唯一。

### 2.3 Aspect 和 Output 定义

```csharp
[Aspect("Physical", "Physical", AspectID.SimplePhysical)]
[SymbolOutput("PNoz1", "Nozzle 1")]
[SymbolOutput("PNoz2", "Nozzle 2")]
[SymbolOutput("PNoz3", "Nozzle 3")]
public AspectDefinition m_Physical;

[Aspect("Insulation", "Insulation", AspectID.Insulation)]
[SymbolOutput("InsulatedBody", "Insulated Body")]
public AspectDefinition m_Insulation;
```

**要点说明**：
- 一个 `AspectDefinition` 字段可声明多个 `SymbolOutput`，每个输出对应一个几何体或端口。
- `AspectID` 常见值：`SimplePhysical`（物理）、`Insulation`（保温）、`Maintenance`（维护空间）。
- 输出名（如 `"PNoz1"`）必须与 `ConstructOutputs` 中 `Outputs.Add` 使用的 key 一致。

### 2.4 ConstructOutputs 分发

```csharp
protected override void ConstructOutputs()
{
    parFacetoCenter = m_FacetoCenter.Value;
    parInsulationThickness = m_InsulationThickness.Value;

    PipeComponent pipeComponent = (PipeComponent)partFclt.Value;
    SP3DConnection connection = this.OccurrenceConnection;
    partDataBasis = pipeComponent.PartDataBasis;

    switch (partDataBasis)
    {
        case 1: case 190: case 192:
            RunPDB1_190_192(pipeComponent, connection);
            break;
        // ... 更多 case 分支
        default:
            UtilityFunctions.RaisePDBNotImplementedError(this, partDataBasis);
            break;
    }
}
```

**要点说明**：
- `ConstructOutputs()` 是符号的**核心入口**，由框架在建模时调用。
- `PartDataBasis`（PDB）是目录中定义的"零件数据基础"编号，决定符号的具体几何形态。
- 每个 PDB（或一组 PDB）对应一个 `RunPDB***` 私有方法。
- `default` 分支必须调用 `RaisePDBNotImplementedError`，避免未支持的规格静默失败。

### 2.5 完整 RunPDB 方法模板（几何 + 端口 + 保温）

```csharp
private void RunPDB3310(PipeComponent pipeComponent, SP3DConnection connection)
{
    UtilityFunctions.CheckParameterForZeroValues(this, "FacetoCenter");

    // 获取端口数据
    double[] pipeDia = new double[3];
    double[] flangeDia = new double[3];
    double[] flangeThck = new double[3];
    UtilityFunctions.GetPipingPortData(1, pipeComponent, out pipeDia[0], out flangeThck[0], out flangeDia[0], ...);

    // Physical：创建主体圆柱
    Projection3d body = Graphics3D.CreateCylinder(
        new Position(-parFacetoCenter, 0, 0),
        new Position(parFacetoCenter, 0, 0),
        pipeDia[0], true);
    m_Physical.Outputs.Add("Body", body);

    // 放置端口
    PipeNozzle port1 = PortHelper.CreatePipeNozzle(
        pipeComponent, connection, false, 1,
        new Position(-parFacetoCenter, 0, 0),
        new Vector(-1, 0, 0), true, true);
    m_Physical.Outputs.Add("PNoz1", port1);

    // Insulation：主体 + 2 倍保温厚度
    Projection3d insBody = Graphics3D.CreateCylinder(
        new Position(-parFacetoCenter, 0, 0),
        new Position(parFacetoCenter, 0, 0),
        pipeDia[0] + 2 * parInsulationThickness, true);
    m_Insulation.Outputs.Add("InsulatedBody", insBody);
}
```

**要点说明**：
- `CheckParameterForZeroValues` 防止零值参数导致几何异常。
- `GetPipingPortData` 从 `PipeComponent` 提取各端口的管径、法兰厚度、法兰直径等数据。
- 保温几何通常是物理几何外扩 `2 * 保温厚度`（直径方向）。
- 端口（`PipeNozzle`）位置与方向必须与几何端面对齐，否则连接关系会错误。

---

## 三、几何组合技术（Clamp.cs 为例）

`Clamp`（管夹）是典型的多种几何体组合符号，共 22 种 PDB 变体。其 `RunPDB1_430` 方法展示了四种几何体的组合使用。

### 3.1 RunPDB1_430 — 四种几何体组合

`Clamp` 通过以下方式组合几何：

1. **Revolution3d（回转体）** — `CreateLineString` + `CreateRevolution` 创建夹体主体。
2. **Projection3d（投影体）** — `CreateProjection` 创建上下支撑。
3. **CreateCylinder（圆柱）** — 创建螺杆与手柄。
4. **CreateBox（长方体）** — 创建保温盒体。

### 3.2 防止零尺寸几何的关键模式

当保温厚度为 0 时，直接创建几何会报错。此时应创建一个极小尺寸（如 `0.001`）的占位几何，避免零尺寸导致的底层 COM 错误。

```csharp
if (UtilityFunctions.CompareDoubleGreaterThan(parInsulationThickness, Constants.LINEAR_TOLERANCE))
{
    stPoint.Set(-clampWidth/2 - parInsulationThickness, 0, 0);
}
else
{
    stPoint.Set(-0.001, 0, 0);  // 防止零尺寸几何
    insulationDia = 0.001;
}
```

**要点说明**：
- `Constants.LINEAR_TOLERANCE` 是 Smart3D 的线性容差常量，用于浮点比较。
- 任何依赖用户输入尺寸的几何创建都应考虑零值或负值的防护。
- 使用 `UtilityFunctions.CompareDoubleGreaterThan` 等方法进行浮点比较，避免直接使用 `>` 运算符导致的精度问题。

---

## 四、HVAC 截面处理（HTee.cs 为例）

HVAC 符号与管道符号的最大区别在于截面形状多样：矩形、圆形、扁椭圆三种。`HTee.cs` 展示了如何根据截面类型统一构建几何。

### 4.1 三种截面类型分发

```csharp
ICurve curve1 = null;
if (parHVACShape == (int)CrossSectionShapeTypes.FlatOval)
    curve1 = CreteFlatOvalBranch(centerPoint, width, depth, planeOfBranch);
else if (parHVACShape == 4)  // Round（圆形）
    curve1 = Graphics3D.CreateCircle(centerPoint, new Vector(0,1,0), diameter/2);
else if (parHVACShape == 1)  // Rectangular（矩形）
    curve1 = CreateRectangularBranch(centerPoint, width, depth, planeOfBranch);

// 统一处理：用 RuledSurface 连接两个截面
Ruled3d branchBody = Graphics3D.CreateRuledSurface(curve1, curve2, true);
```

**要点说明**：
- 三种截面类型最终都转换为 `ICurve`（截面曲线），从而统一后续处理。
- `CreateRuledSurface` 在两个截面曲线之间放样，形成主体几何——这是 HVAC 符号最核心的建模思路。
- `CrossSectionShapeTypes` 枚举定义截面类型：`FlatOval`（扁椭圆）、圆形（值 4）、矩形（值 1）。

### 4.2 矩形截面构建辅助方法

`CreateRectangularBranch` 方法通过 4 条直线构建矩形截面，再用 `ComplexString3d` 组合，最后根据 `PlaneOfBranch` 用 `Matrix4X4` 旋转到正确朝向。

```csharp
// 构建矩形截面的核心思路（简化示意）：
Collection<ICurve> lines = new Collection<ICurve>();
lines.Add(Graphics3D.CreateLine(p1, p2));
lines.Add(Graphics3D.CreateLine(p2, p3));
lines.Add(Graphics3D.CreateLine(p3, p4));
lines.Add(Graphics3D.CreateLine(p4, p1));
ComplexString3d rect = Graphics3D.CreateComplexString(lines);

// 根据 PlaneOfBranch 旋转截面
Matrix4X4 rotMatrix = new Matrix4X4();
rotMatrix.Rotate(...);  // 绕指定轴旋转
// 对 rect 应用变换
```

### 4.3 HVAC 端口创建

```csharp
// 基础 HVAC 端口
HvacPort port = PortHelper.CreateHVACPort(hvacPart, connection, portNum,
    position, direction, portDirection2, width, depth, false);

// 带长度的 HVAC 端口
HvacPort port = PortHelper.CreateHVACPortWithLength(hvacPart, connection, portNum,
    position, direction, portDirection2, length, width, depth, false);
```

**要点说明**：
- HVAC 端口需要两个方向向量（`direction` 和 `portDirection2`）来定义截面朝向。
- `width` 和 `depth` 定义端口截面尺寸。
- 带长度版本用于端口本身有延伸段的情况。

---

## 五、电缆规则开发（CableFillRule.cs）

`CableFillRule` 是基于 **NEC 2008（美国国家电气规范）** 标准的电缆填充率计算规则引擎，根据电缆类型、导体截面积、电缆外径等参数计算允许的最大填充率。

### 5.1 规则类骨架

```csharp
public class CableFillRule : CableFillCalculationRule
{
    // 主入口（重写）
    public override void GetCableFillValueForFeature(
        RouteFeature cablewayFeature, ref CableFillDetails cableInfo)
    {
        RouteRun oRun = cablewayFeature.Run;
        if (oRun is Cableway)
            GetCwayFillParams(...);
        else if (oRun is ConduitRun)
            GetConduitFillParams(...);

        cableInfo.percentFill = dPercentFill;
        cableInfo.fillStatus = "Full"; // "Partial" / "OverFilled"
    }

    // 实时填充计算（重写）
    public override void CableRealTimeFill(
        CableRun oCableRun, RouteFeature oFeat, out double dRealFill)
    {
        // ...
    }
}
```

**要点说明**：
- `CableFillCalculationRule` 是电缆填充率规则的基类。
- `GetCableFillValueForFeature` 是主计算入口，根据特征类型（电缆桥架 / 导管）分发。
- `fillStatus` 取值：`"Full"`（满）、`"Partial"`（部分）、`"OverFilled"`（超填）。
- `CableRealTimeFill` 用于实时计算单根电缆的填充率，支持拖拽预览。

### 5.2 NEC 表 XML 加载模式

NEC 标准表格存储在 `NECStandardTables.xml` 中，规则引擎通过 XPath 查询读取允许截面积等参数。

```csharp
// 加载 NEC 标准 XML 表
XmlDocument necDoc = new XmlDocument();
necDoc.Load(necTablesXmlPath);

// XPath 查询示例：根据导管尺寸查询允许截面积
XmlNode node = necDoc.SelectSingleNode(
    "//ConduitTable[@Size='" + conduitSize + "']/MaxFillArea");
double allowedArea = Convert.ToDouble(node.InnerText);
```

**要点说明**：
- NEC 表区分单导体与多导体电缆的填充规则（单导体按面积，多导体按外径）。
- XML 表与规则代码解耦，便于标准升级时只更新 XML。

### 5.3 单位转换

Smart3D 内部使用 DBU（数据库单位）存储尺寸，规则计算时需转换为目标单位。

```csharp
m_UOM = MiddleServiceProvider.UOMMgr;

// DBU → 英寸
double inches = m_UOM.ConvertDBUtoUnit(UnitType.Distance, value, UnitName.DISTANCE_INCH);

// 平方英寸 → DBU
double dbu = m_UOM.ConvertUnitToDBU(UnitType.Area, value, UnitName.AREA_SQUARE_INCH);
```

**要点说明**：
- NEC 标准使用英制单位（英寸、平方英寸），而 Smart3D 内部为 DBU，必须转换。
- `UnitType` 区分 `Distance`（距离）、`Area`（面积）等量纲。
- `UnitName` 指定具体单位，如 `DISTANCE_INCH`、`AREA_SQUARE_INCH`、`DISTANCE_MM` 等。

---

## 六、符号示例总览表

| 符号 | PDB 变体数 | 核心几何技术 | 特殊功能 |
|------|-----------|-------------|---------|
| **Tee**（管道三通） | 33 | `CreateCylinder` + `CreateRuledSurface` | 3 个端口 + 保温层 |
| **Flange**（法兰） | 多种 | `DrillingTemplatePattern` | 圆 / 方 / 矩 / 三角 / 椭圆螺栓孔布置 |
| **Clamp**（管夹） | 22 | `CreateRevolution` + `CreateProjection` + `CreateCylinder` + `CreateBox` | 四种几何体组合 |
| **Plug**（管塞） | 10+ | `CreateCylinder` + `CreatePolygon` + `CreateRevolution` | 多边形截面塞体 |
| **OPLever**（杠杆操作器） | 多种 | `CreateRuledSurface` + `CreateCylinder` + `CreateMaintenanceVolume` | 执行器维护空间 |
| **HTee**（HVAC 三通） | 6 | 矩形 / 圆形 / 扁椭圆截面 | 三种截面类型统一处理 |
| **Wye**（HVAC Y 型分支） | 7 | `CreateProjection` + `CreateEllipticalArc` + `CreateRuledSurface` + `CreateCone` + Skinning | 蒙皮（Skinning）高级曲面技术 |

---

## 七、UtilityFunctions 常用方法

`UtilityFunctions` 是符号开发中的通用工具类，提供参数校验、浮点比较、端口数据提取等常用功能。

| 方法 | 签名 | 说明 |
|------|------|------|
| `CheckParameterForZeroValues` | `(this, paramName)` | 校验指定参数是否为零，为零则抛出异常 |
| `GetPipingPortData` | `(portIndex, pipeComponent, out diameter, out flangeThickness, out flangeDiameter, ...)` | 提取管道端口的管径、法兰厚度、法兰直径等数据 |
| `CompareDoubleEqualTo` | `(d1, d2)` | 判断两个浮点数是否相等（考虑容差） |
| `CompareDoubleGreaterThan` | `(d1, d2)` | 判断 `d1 > d2` |
| `CompareDoubleLessThanOrEqualTo` | `(d1, d2)` | 判断 `d1 <= d2` |
| `RaisePDBNotImplementedError` | `(this, pdb)` | 抛出"PDB 未实现"错误，用于 `switch` 的 `default` 分支 |
| `Constants.LINEAR_TOLERANCE` | — | 线性容差常量，用于浮点比较与零尺寸防护 |

**使用建议**：
- 所有浮点比较一律使用 `UtilityFunctions.CompareDouble***` 方法，禁止直接使用 `==`、`>`、`<` 运算符，以避免浮点精度问题。
- 每个 `RunPDB***` 方法开头应调用 `CheckParameterForZeroValues` 校验关键尺寸参数。
- `Constants.LINEAR_TOLERANCE` 是判断"是否为零尺寸"的标准阈值，配合 `CompareDoubleGreaterThan` 使用。
