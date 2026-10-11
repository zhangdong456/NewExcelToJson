# ExcelToJsonCs · Excel 配置表导出工具

> 扫描一个目录下所有的 Excel 配置表（`csv` / `xls` / `xlsx` / `xlsm`），
> 为每张表导出一份 **JSON 数据** 和一份 **C# 解析类**，可一键打包成 Windows exe 双击运行。
> 内置完整的**校验与分级日志**，拒绝静默失败——任何异常数据都会在日志里被点名。

---

## 目录

- [一、功能特性](#一功能特性)
- [二、目录结构](#二目录结构)
- [三、环境要求](#三环境要求)
- [四、快速开始](#四快速开始)
- [五、Excel 表格约定](#五excel-表格约定)
- [六、支持的类型](#六支持的类型)
- [七、配置文件 config.json](#七配置文件-configjson)
- [八、校验规则与日志](#八校验规则与日志)
- [九、输出产物](#九输出产物)
- [十、C# 侧接入示例](#十c-侧接入示例)
- [十一、增量导出缓存](#十一增量导出缓存)
- [十二、常见问题 FAQ](#十二常见问题-faq)
- [十三、已知边界](#十三已知边界)
- [十四、License](#十四license)

---

## 一、功能特性

| 能力 | 说明 |
|---|---|
| 多格式支持 | `csv` / `xls` / `xlsx` / `xlsm`，每个文件只读取**第一个 Sheet** |
| 零配置起步 | 首次运行自动生成 `config.json`，改一个源目录即可用 |
| 数组字段 | `int[]` / `string[]` / `List<...>` 在表里横向占多列，导出为**真正的 JSON 数组** |
| 增量导出 | 用 `size + mtime + md5` 判断文件是否变化，没变就跳过，秒级完成 |
| 校验 + 日志 | 表名 / Sheet / 行号 / 列号 / 字段名 / 具体原因，中文、可读、分级、可追溯 |
| 跨表引用校验 | 可配置规则，检查 A 表字段引用的 Id 是否真存在于 B 表 |
| 可裁剪产物 | Unity / Newtonsoft / 自定义 `[Config]` 特性均可开关 |
| 一键打包 | 附带 PyInstaller 配置，打包成独立 exe，无需用户装 Python |

---

## 二、目录结构

```
ExcelToJson/
├─ ExcelToJsonCs.py            # 工具主程序（单文件，无框架依赖）
├─ ExcelToJsonCs.spec          # PyInstaller 打包配置
├─ config.json                 # 运行配置（首次运行自动生成）
├─ README.md                   # 本文件
├─ .gitignore                  # Git 忽略规则
│
├─ Role.xlsx                   # 示例源表（演示各种类型/数组/注释）
├─ ParseExample.cs             # C# 侧「读取并反序列化 JSON」的示例代码
├─ Role.cs / RoleCategory.json # 示例导出产物（与 Cs/、Json/ 内容一致）
├─ Json/                       # JSON 输出目录（含示例产物）
├─ Cs/                         # C#   输出目录（含示例产物）
│
├─ build/                      # PyInstaller 中间产物      [已忽略]
└─ dist/                       # PyInstaller 打包结果 exe  [已忽略]
```

---

## 三、环境要求

运行源码需要：

- **Python 3.8+**
- 第三方库：
  - `openpyxl` —— 读取 `.xlsx` / `.xlsm`
  - `xlrd` —— 读取旧版 `.xls`
  - `csv` —— 由标准库提供，无需安装
- 打包 exe 需要：`pyinstaller`

```bash
pip install openpyxl xlrd pyinstaller
```

> 直接使用打包好的 `ExcelToJsonCs.exe` 时，用户机器**无需**安装任何东西。

---

## 四、快速开始

### 方式一：使用打包好的 exe（推荐给策划 / 美术）

1. 把 `ExcelToJsonCs.exe` 和 `config.json` 放到同一个目录；
2. 打开 `config.json`，把 `EXCEL_SOURCE_DIR` 改成你的配置表目录：

   ```json
   {
     "EXCEL_SOURCE_DIR": "D:/GameProjects/Configs/Excel"
   }
   ```

3. **双击** `ExcelToJsonCs.exe`，等待控制台输出 `处理完成`。
4. 结果会出现在 exe 同目录的 `Json/` 和 `Cs/` 下，日志在 `export_log.txt`。

> Windows 下双击 exe 结束后窗口不会立刻关闭，可回滚查看输出；按回车退出。

### 方式二：Python 直接运行源码

```bash
# 在工程根目录
python ExcelToJsonCs.py
```

### 方式三：自行打包成 exe

```bash
# 在工程根目录（已附带 spec 文件）
pyinstaller --noconfirm ExcelToJsonCs.spec
# 产物：dist/ExcelToJsonCs.exe
```

`ExcelToJsonCs.spec` 已配置为 **onefile + console**（单文件、带控制台窗口）。

---

## 五、Excel 表格约定

工具对表格结构有固定约定，**前 3 行是结构行，从第 4 行开始才是数据**：

| 行号 | 含义 | 示例 |
|:---:|---|---|
| 第 1 行 | 字段英文名（即生成的 C# 字段名） | `Id`、`Name`、`Cast` |
| 第 2 行 | 字段类型 | `int`、`string`、`string[]` |
| 第 3 行 | 中文注释（生成到 C# 字段后的 `//`） | `编号`、`名字`、`特性` |
| 第 4 行起 | 数据 | `5001`、`天蝎`、`6090` |

以示例表 `Role.xlsx` 为例（`_` 表示空单元格）：

```
      A     B       C            D         E        F    G    H
1    Id    Name    Name_short   NickName  Cast     _    _    Ability_1   ...
2    int   string  string       string    string[] _    _    int         ...
3    编号  名字    短名字        外号      特性      _    _    体质        ...
4    5001  _       _            天蝎      6090     _    _    7           ...
```

**规则要点：**

1. **数组字段横向占多列**：`Cast` 是 `string[]`，类型行占 E、F、G 三列，**续列的「表头」必须留空**——工具靠"表头为空"来判断数组的跨列边界。导出后它就是一个 JSON 数组。
2. **注释行**：某行**第 1 列（最左列）**以 `#` 开头 → 整行忽略。
3. **注释列**：某列**第 1 行（表头）**以 `#` 开头 → 整列忽略（可用来临时屏蔽不想导出的字段）。
4. **只读取第一个 Sheet**（`.csv` 只有一张"表"）。
5. **第 1 个字段必须是非数组类型**（`int` 或 `string`），它会被当作**主键**，用于生成 C# 的索引字典。

---

## 六、支持的类型

类型写在**第 2 行**，大小写不敏感、空格会被忽略（`List<int>`、`list< int >` 等价）。

| Excel 第 2 行写法 | 生成的 C# 类型 | 说明 |
|---|---|---|
| `int` | `int` | 整数，主键常用 |
| `string` | `string` | 字符串 |
| `int[]` / `List<int>` | `int[]` / `List<int>` | 整数集合，**横向多列** |
| `string[]` / `List<string>` | `string[]` / `List<string>` | 字符串集合，**横向多列** |

> 出现表里没定义的其它类型（如 `float`、`bool`）会被判为**致命错误并中止该文件**，防止生成错误的 C#。需要新类型时请改工具的类型表。

---

## 七、配置文件 config.json

配置文件与工具（exe）放在**同一目录**；不存在时会自动生成带默认值的版本。

| 键 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `EXCEL_SOURCE_DIR` | string | `"."` | 配置表源目录。相对路径基于工具目录，绝对路径直接使用 |
| `JSON_OUTPUT_DIR` | string | `"Json"` | JSON 输出目录（相对工具目录） |
| `CS_OUTPUT_DIR` | string | `"Cs"` | C# 输出目录（相对工具目录） |
| `USE_UNITY` | bool | `true` | 是否生成 Unity 相关内容（`using UnityEngine;`、`[SerializeField]`）。纯 .NET 项目改 `false` |
| `USE_NEWTONSOFT` | bool | `true` | 是否使用 `Newtonsoft.Json`（`using Newtonsoft.Json;`、`[JsonIgnore]`）。用 `System.Text.Json` 改 `false` |
| `USE_CONFIG_ATTR` | bool | `true` | 是否为 `Category` 包装类添加 `[Config]` 特性（依赖你的框架，不需要改 `false`） |
| `LOG_FILE` | string | `"export_log.txt"` | 日志文件名（相对工具目录）；设为 `""` 则只输出到控制台 |
| `INT_RANGE_CHECK` | bool | `true` | 是否校验 `int` 是否落在 C# `Int32` 范围 `[-2147483648, 2147483647]` |
| `MAX_ISSUES_PER_FILE` | int | `300` | 单文件最多逐条输出的提示数，超出部分折叠成一行汇总，防刷屏 |
| `REQUIRED_FIELDS` | object | `{}` | 额外必填字段：`{"表名": ["字段1", "字段2"]}`，命中为空时告警（主键列始终必填） |
| `REF_RULES` | object | `{}` | 跨表引用校验规则，见下 |

**跨表引用规则 `REF_RULES`** 的两种写法：

```json
{
  "REF_RULES": {
    "Role.Cast":     "Trait",         // 校验 Role.Cast 里的 Id 都存在于 Trait 表的主键列
    "Role.SepName":  "Nation.Id"      // 校验 Role.SepName 里的 Id 都存在于 Nation 表的 Id 列
  }
}
```

- 键：`源表名.源字段名`；值：`目标表名` 或 `目标表名.目标字段名`。
- 不写目标字段时，默认校验目标表的**主键列**（第 1 个字段）。
- 该表若因增量缓存被跳过，本次不会对它做引用校验（改动源表或清缓存即可）。

---

## 八、校验规则与日志

每次运行会同时把日志写到**控制台**和 `LOG_FILE`（默认 `export_log.txt`，每次覆盖）。

### 日志格式

```
[级别] [表名] Sheet:<Sheet名> 第<行号>行 <列号>列 字段'<字段名>': <中文原因>
```

真实示例：

```
[ERROR] [Role.xlsx] Sheet:Sheet1 第5行 A列 字段'Id': 主键列为空，已跳过该行（该行其余非空单元格: B列='b'、C列='200'）。
[WARN]  [Role.xlsx] Sheet:Sheet1 第7行 E列 字段'Cast': 值 'abc' 不是合法整数，已按 0 处理。
[ERROR] [Bad.xlsx]  Sheet:Sheet1 第1行 D列 字段'Bad-Type': 字段名不是合法的 C# 标识符（只能由字母/数字/下划线组成，且不能以数字开头）。
```

### 三级处理策略

| 级别 | 含义 | 处理动作 |
|:---:|---|---|
| **ERROR（文件级致命）** | 结构性问题，继续导出会产出**错误的代码/数据** | 提示 + **中止该文件**，不写任何产物、不更新缓存 |
| **ERROR（行级）** | 该行数据有问题 | 提示 + **跳过该行**，其余行照常导出 |
| **WARN** | 可疑但不致命 | 提示 + **容错继续**（用安全默认值处理） |

### 规则总表

| 级别 | 触发场景 | 处理 |
|:---:|---|---|
| ❌ 致命 | 表头行为空 | 中止文件 |
| ❌ 致命 | 缺少第 2 行（类型行） | 中止文件 |
| ❌ 致命 | 某字段的类型行为空 | 中止文件 |
| ❌ 致命 | 类型不在支持列表（如 `float`） | 中止文件 |
| ❌ 致命 | 字段名不是合法 C# 标识符（含中文/空格/数字开头等） | 中止文件 |
| ❌ 致命 | 字段名是 C# 保留字（如 `class`） | 中止文件 |
| ❌ 致命 | 字段名重复 | 中止文件 |
| ❌ 致命 | 第 1 个字段是数组/集合类型（不能当字典键） | 中止文件 |
| ⛔ 跳行 | **主键列（第 1 个字段）为空** | 跳过该行（并打印该行其余非空单元格便于定位） |
| ⛔ 跳行 | 主键为 `int` 但填了非整数 | 跳过该行 |
| ⛔ 跳行 | 主键重复 | 跳过该行（避免 C# 端 `dict.Add` 运行时抛异常） |
| ⛔ 跳行 | 引用校验失败（`REF_RULES` 命中的 Id 不存在） | 报错（不阻断源表导出） |
| ⚠️ 警告 | 缺少第 3 行（注释行） | 注释回退为字段名 |
| ⚠️ 警告 | `int` 列填了非数字 / 小数 | 取整（小数截断）后继续 |
| ⚠️ 警告 | `int` 超出 `Int32` 范围 | 继续（建议改用 `long`） |
| ⚠️ 警告 | 字符串含控制字符 | 继续（提示可能误粘贴了不可见字符） |
| ⚠️ 警告 | 数据区中间出现空行 | 跳过该空行 |
| ⚠️ 警告 | 表头为空但类型行有值（游离列） | 忽略该列 |
| ⚠️ 警告 | 数组字段只占 1 列（续列表头没留空） | 继续，但数组无法展开 |
| ⚠️ 警告 | `REQUIRED_FIELDS` 配置的必填字段为空 | 继续 |
| ⚠️ 警告 | 数据区出现表头定义范围之外的单元格 | 忽略（只报一次） |

### 运行结果汇总

```
========================================================
处理完成: 导出 1 个 | 跳过 0 个 | 失败 0 个 | 警告 0 条 | 错误 0 条
详细日志: ...\export_log.txt
========================================================
```

---

## 九、输出产物

每个源表 `Xxx.xlsx` 会产出两个文件：

### 1. JSON —— `Json/XxxCategory.json`

```json
{
  "dataList": [
    {
      "Id": 5001,
      "NickName": "天蝎",
      "Cast": ["6090"],
      "Ability_1": 7
    }
  ]
}
```

- 数组字段是**真正的 JSON 数组**（`["6090"]`）；
- 空单元格：`int` → `0`，`string` → `""`，数组 → 空数组或 `[]`；
- 中文直出（`ensure_ascii=False`），文件编码 UTF-8。

### 2. C# —— `Cs/Xxx.cs`

包含两个 `partial class`：

- `XxxCategory`：包装类，持有 `Dictionary<主键, Xxx> dict`，提供 `GetOne(id)` / `isContain(id)` / `GetDict()`，在 `EndInit()` 里建索引；
- `Xxx`：行类，每个字段一行，行尾带第 3 行的中文注释。

生成的代码头部带有醒目提示：**属于自动生成，切勿手改，重新导表会被覆盖**。

---

## 十、C# 侧接入示例

参考工程内的 `ParseExample.cs`（使用 Newtonsoft.Json）：

```csharp
string json = File.ReadAllText("Json/RoleCategory.json");
RoleCategory cfg = JsonConvert.DeserializeObject<RoleCategory>(json);

cfg.EndInit();                       // 建索引字典

Role r = cfg.GetOne(5002);           // 按主键取一条
Console.WriteLine(r.NickName);       // 千里眼

foreach (var c in r.Cast)            // 遍历数组字段
    Console.WriteLine(c);

bool has = cfg.isContain(5001);      // 判存在
var all  = cfg.GetDict();            // 取全部
```

> 想在**纯 .NET 控制台**里编译自测：把 `config.json` 的 `USE_UNITY` 改为 `false` 重新导出，
> 生成的 `Role.cs` 就不再依赖 `UnityEngine`，可直接独立运行。

---

## 十一、增量导出缓存

工具在同目录维护 `.excel_export_cache.json`，记录每个文件的 `size / mtime / md5`：

- `size` 与 `mtime` 都没变 → **直接跳过**，不读取文件内容；
- 只有 `mtime` 变了但 `md5` 一致 → 更新时间戳后跳过。

**需要“强制重新导出”时**，删掉部署目录里的 `.excel_export_cache.json` 即可
（例如希望新版本工具的全新校验逻辑对旧表重新跑一遍）。

> 因**致命错误而中止**的文件**不会**写入缓存，下次运行会再次校验，直到问题修好。

---

## 十二、常见问题 FAQ

**Q：导出的 JSON 是空的 / 少了几行？**
先看日志。最常见的是**主键列为空**——这类行会被明确报 `ERROR ... 主键列为空，已跳过该行`，并打印该行其它非空单元格，方便定位。

**Q：改了表格，但重新运行没有更新？**
增量缓存把它跳过了。确认文件确实保存过；必要时删除 `.excel_export_cache.json`。

**Q：为什么我的数组列只导出了第一个值？**
数组字段的**续列表头必须留空**。如果紧邻的下一列写了表头，工具会认为数组只占 1 列，日志会告警 `数组无法横向展开`。

**Q：exe 双击一闪而过 / 没反应？**
程序结束时按回车才关闭窗口；若真的异常退出，看同目录的 `export_log.txt`。源目录不存在时也会给出明确提示。

**Q：中文乱码？**
JSON 以 UTF-8 写入；用支持 UTF-8 的编辑器打开。控制台已强制 UTF-8 输出。

**Q：想临时屏蔽某几个字段？**
把对应列的**表头**改成以 `#` 开头（该列会被整体忽略），或把某行最左列设为 `#` 开头忽略整行。

**Q：能只处理某一张表吗？**
最简单的方式是把源目录指向只含该表的文件夹，或临时把其它表移走（工具会扫描目录下所有 Excel 文件）。

---

## 十三、已知边界

- **只读第一个 Sheet**，不支持跨 Sheet 校验。
- **不做类型归一化比较**：引用校验按字符串精确匹配（`'01'` 与 `'1'` 视为不同）。
- **仅校验 `Int32` 范围**，暂不支持按字段自定义 `min/max`。
- **控制字符检测**只覆盖 ASCII 控制字符，不包含全角/零宽字符/emoji。
- **空行判定**：整行所有列皆空才算空行。
- **公式单元格**取缓存值（`data_only`，无缓存即视为空）；**合并单元格**仅左上角有值。
- **主键唯一性为全局唯一**，不支持复合主键。
- 单文件同类提示超过 `MAX_ISSUES_PER_FILE`（默认 300）会折叠为汇总行，逐行细节可能省略。
- 引用校验只对**本次实际导出**的源表执行；被增量缓存跳过的表不参与。

---

## 十四、License

本项目仅供学习与内部使用；如需开源请补充 License 文件。

---

<sub>部分产物（JSON / C# / exe）由工具自动生成，请勿手工修改。</sub>
