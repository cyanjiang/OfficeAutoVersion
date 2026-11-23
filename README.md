# Office 自动版本插件（OfficeAutoVersion）

在 **Word / Excel / PowerPoint** 中打开插件后，每次保存文档时，都会自动把磁盘上保存前的内容
复制成一份带时间戳的**版本快照**，相当于给每个文档一个轻量的“历史版本”能力。
全部在本机完成，不依赖任何云服务。

```
D:\项目\报告.docx                      ← 当前文档
D:\项目\报告 v1_Versions\              ← 自动生成的版本文件夹
        报告 v1_20261025-143012.docx  ← 第 1 次保存前的内容
        报告 v1_20261025-151833.docx  ← 第 2 次保存前的内容
        报告 v1_20261026-090211.docx  ← 第 3 次保存前的内容
```

---

## 1. 功能特性

| 功能 | 说明 |
| --- | --- |
| 保存即建版本 | 挂接 Word/Excel/PowerPoint 的“保存前”事件，自动生成快照 |
| 时间戳命名 | `名称_yyyyMMdd-HHmmss.扩展名`，同一秒内多次保存自动加 `(1)(2)` 防冲突 |
| 内容去重 | 内容与最新版本完全一致（MD5）时自动跳过，不产生重复快照 |
| 数量上限 | 默认保留最近 20 个版本，超出的旧版本**移入回收站**（可恢复）；设为 0 表示不限 |
| 两种存放位置 | ① 原文件同目录的 `xxx_Versions` 文件夹（默认，随文件走）；② 集中目录（文档\OfficeAutoVersions，镜像原路径） |
| 文件名格式可配 | 支持 `{name}` `{ext}` `{ts}` 及任意 .NET 日期格式（如 `{yyyy年MM月dd日}`） |
| 功能区开关 | 新增“自动版本”选项卡：总开关、版本历史、打开版本文件夹、设置 |
| 版本历史窗口 | 列出某文档的全部版本（生成时间 / 大小 / 文件名），双击即比较，可恢复、可打开 |
| **比较差异** | 调用 **Office 自带的比较能力**，不自己造 diff：Word「比较文档」生成带修订的对比稿；PowerPoint「审阅 → 比较」；Excel「并排查看」同步滚动 |
| **差异摘要** | 版本历史窗口一键生成“当前文档 vs 所选版本”的**内容级差异清单**：新增 / 删除 / 修改各多少处，逐条给出位置与新旧内容，可复制、可存为 txt。Word 还会**标出改动所在页码**（页码由 Word 按当前排版算出，文件里不保存）。纯文件解析（zip + XML）得到内容，不打开 Office 窗口 |
| **恢复版本** | 用所选版本覆盖当前文件；**恢复前自动把当前状态另存为一个新版本**，恢复错了可再恢复回来 |
| 绝不添乱 | 任何失败（只读目录、无权限、文件被占用……）只写日志，**绝不影响用户的保存操作**；原文件旁边放不下时自动退回到集中目录 |
| 每应用独立配置 | Word/Excel/PowerPoint 各自一套配置，互不影响 |

### 差异摘要的口径（避免误读）

| 应用 | 位置口径 | 比较方式 | 页码 |
| --- | --- | --- | --- |
| Excel | 工作表!单元格地址（如 `Sheet1!B2`） | 按地址键值对比，与行顺序无关 | 无（表格无“页”） |
| Word | 第N段；删除的条目写作「版本第N段」 | **只比段落内容，不比段落编号**，LCS 对齐；相邻的删除+新增配对成“修改” | **有**，由 Word 按当前排版逐段算出 |
| PowerPoint | 幻灯片N · 形状名 | 按“幻灯片+形状”键值对比 | 无（幻灯片序号即位置） |

Word 这一条尤其重要：段落编号会随前面段落的增删整体平移，如果把编号算进比较内容，
在开头插入一段就会把后面所有段落都误报成差异。因此编号只用于**展示定位**，不参与比较。
页码同理只是排版结果，docx 文件里没有页码，只能向打开中的 Word 逐段询问；取不到就不标，不影响摘要本身。

命令行核对（不启动 Office，直接对两个文件出摘要）：

```
OfficeAutoVersion.Selftest.exe diff <当前文件> <版本文件>
```

## 2. 技术选型：为什么是 VSTO

| 方案 | 能否满足“保存时自动建版本” |
| --- | --- |
| **VSTO（.NET Framework）** ✅ | 可以挂接 `DocumentBeforeSave` 等宿主事件，直接读写本地文件系统，随 Office 桌面版启动 |
| Office Web 加载项（Office.js） | ❌ 受沙箱限制，无法把文件静默写到本地磁盘，也没有可用的“保存前”事件 |
| VBA 宏 | 能实现同样效果，但需手动分发宏、按应用各写一份、维护困难（本文末尾附了一份 VBA 参考实现） |

本插件只支持 **桌面版 Office（Windows）**，不支持 Office for Mac / Office 网页版。

## 3. 工作原理

各应用挂接的事件（均为“保存**前**”事件，因此快照保存的是本次保存之前的磁盘内容，
与当前文档合起来构成完整的历史）：

| 应用 | 事件 | C# 处理器签名 |
| --- | --- | --- |
| Word | `Application.DocumentBeforeSave` | `(Word.Document, ref bool SaveAsUI, ref bool Cancel)` |
| Excel | `Application.WorkbookBeforeSave` | `(Excel.Workbook, bool SaveAsUI, ref bool Cancel)` |
| PowerPoint | `Application.PresentationBeforeSave` | `(PowerPoint.Presentation, ref bool Cancel)` |

处理流程（每一步失败都只记日志，不抛出、不拦截保存）：

```
保存事件触发
   │
   ▼
插件已启用？ ──否──▶ 直接返回
   │是
   ▼
文件已保存过磁盘？ ──否（从未保存的新文档）──▶ 跳过
   │是
   ▼
计算版本目录 ──▶ 原文件旁 <名称>_Versions/ 或 集中目录（镜像原路径）
   │
   ▼
内容与最新版本相同？ ──是──▶ 跳过（避免重复快照）
   │否
   ▼
复制文件为 名称_时间戳.扩展名（重名自动加序号）
   │
   ▼
超过保留数量？ ──是──▶ 最旧的版本移入回收站
   │
   ▼
写日志（%APPDATA%\OfficeAutoVersion\logs\autoversion.log）
```

> 注：文档从未保存过时没有“保存前”的旧内容，属于正常跳过；
> “另存为”会在旧路径上留下一个快照，同样是有意为之。

## 3.1 比较与恢复：全部交给 Office 自己做

版本管理最容易做砸的是“差异比较”——自己实现文本 diff 既不可靠也不符合用户习惯。
本插件**不自己算差异**，而是把两个文件交给 Office 自带的比较能力：

| 应用 | 比较方式（Office 自带） | 用户看到什么 | 为什么这样选 |
| --- | --- | --- | --- |
| **Word** | `Application.CompareDocuments`（等同“审阅 → 比较”） | 新开一份对比文档，所有差异都是**修订标记**，可逐条接受/拒绝 | Word 的比较引擎是业内最完整的（格式、表格、页眉页脚、脚注、批注都能比） |
| **PowerPoint** | `Presentation.Merge`（等同“审阅 → 比较”） | 右侧“审阅”窗格列出增删改，可逐条接受/拒绝 | PPT 没有独立的对比文档，官方就是合并式审阅 |
| **Excel** | `Windows.CompareSideBySideWith` + 强制开启同步滚动（等同“视图 → 并排查看”） | 两个工作簿窗口左右并排、一起滚动 | **Office 2024 已移除 Spreadsheet Compare 独立工具**，Excel 内部也没有内容级 diff API；并排查看是唯一可靠的自带方案 |

恢复（回滚）采用**文件级替换**，这是三个应用统一、可预期的方式：

```
选中某个版本 → 点“恢复到此版本”
   │
   ▼
先保存当前文档（把未保存的修改落盘）
   │
   ▼
把“当前状态”额外存成一个新版本  ← 关键：恢复本身也可逆
   │
   ▼
关闭当前文档（不保存）→ 用版本文件覆盖 → 重新打开
```

两个实现细节值得注意：

- **关闭文档前一定会先建安全快照**，所以“恢复”永远可撤销——包括误点；
- 恢复期间的保存会临时关掉自动快照（`SetSuppressAutoVersion`），
  否则一次恢复会同时产生“保存前”和“保存后”两份版本，白白多一份。

> Word 里有个坑专门处理过：`Documents.Open` 对**已经打开**的文档会返回同一个 `Document` 对象，
> 如果比较时重新打开“当前文档”、之后又关闭它，就会把用户正在编辑的文档一起关掉。
> 因此 Word 的比较直接使用用户当前打开的文档当“修订版”，只把版本文件临时打开。

## 4. 工程结构

```
OfficeAutoVersion/
├── OfficeAutoVersion.sln                  # 解决方案（5 个工程）
├── src/
│   ├── OfficeAutoVersion.Core/            # 核心版本引擎（net48 类库，无 Office 依赖）
│   │   ├── VersionManager.cs              #   快照/去重/清理主流程
│   │   ├── VersionPaths.cs                #   版本目录与文件名规则（纯函数）
│   │   ├── VersionOptions.cs              #   配置项
│   │   ├── OptionsStore.cs                #   配置持久化（%APPDATA%\OfficeAutoVersion\settings.json）
│   │   ├── MiniJson.cs                    #   零依赖 JSON 读写
│   │   ├── Logger.cs                      #   滚动日志
│   │   ├── RecycleBin.cs                  #   SHFileOperation 回收入口
│   │   ├── ContentDiff.cs                 #   差异引擎（键值比较 + 段落 LCS）
│   │   ├── OpcContent.cs                  #   从 docx/xlsx/pptx（zip+XML）提取带位置的内容
│   │   └── VersionDiffer.cs               #   差异摘要入口（按扩展名分发）
│   ├── OfficeAutoVersion.Common/          # 三个加载项共享层（net48）
│   │   ├── VersioningRibbon.cs            #   功能区（Ribbon XML）
│   │   ├── Ribbon.xml                     #   功能区定义（嵌入资源）
│   │   ├── VersionHistoryForm.cs           #   版本历史窗口（列出 / 比较 / 差异摘要 / 恢复）
│   │   ├── DiffSummaryForm.cs              #   差异摘要窗口（统计 + 明细，可复制/保存）
│   │   ├── IVersionHostActions.cs          #   宿主能力接口（比较、恢复由各加载项实现）
│   │   └── SettingsForm.cs                #   设置对话框（WinForms）
│   ├── OfficeAutoVersion.Word/            # Word 加载项（VSTO）
│   │   └── WordVersionActions.cs          #   比较 = Word「比较文档」；恢复 = 文件级替换
│   ├── OfficeAutoVersion.Excel/           # Excel 加载项（VSTO）
│   │   └── ExcelVersionActions.cs         #   比较 = Excel「并排查看」；恢复 = 文件级替换
│   ├── OfficeAutoVersion.PowerPoint/      # PowerPoint 加载项（VSTO）
│   │   └── PowerPointVersionActions.cs    #   比较 = PPT「审阅→比较」；恢复 = 文件级替换
├── tests/OfficeAutoVersion.Selftest/      # 自测工程（net8.0，直接链接交付源码）
└── tools/
    ├── smoketest-excel.ps1                # 真机冒烟测试（保存→生成版本）
    └── ActionHarness/                     # 真机动作测试（反射调用真实动作类，验证比较/恢复）
```

## 5. 环境要求

- Windows + **桌面版 Office**（Word/Excel/PowerPoint 任一或全部；2013 及以后均可）
- **Visual Studio 2022 / 2026**（Community 免费版即可），安装工作负载
  **“Office/SharePoint 外接程序开发”**（提供 VSTO 生成目标和 Office PIA）
- .NET Framework 4.8 开发人员包（VS 会自动带）
- 仅构建/自测核心逻辑：.NET 8 SDK 即可

> **本机已验证**：VS 2026 Community（`D:\Tools\System\Microsoft Visual Studio\18\Community`）
> + “Office/SharePoint 外接程序开发”工作负载 + Office 2024（16.0）。
> 解决方案已用 MSBuild 18.11.1 完整编译通过，三个加载项均产出
> `.dll` + `.dll.manifest` + `.vsto`，并在真实 Excel 2024 中跑通了“保存即生成版本”的冒烟测试（见第 9 节）。
>
> ⚠️ 微软已宣布**从 Visual Studio 2026 起弃用 VSTO 开发、并在后续版本中移除**。
> 当前版本（18.x）仍可正常创建、构建、F5 调试；若要做长期维护的新项目，
> 建议评估 Office Web Add-in（Office.js）路线。

## 6. 构建与运行

### 在 Visual Studio 中（推荐）

1. 用 VS 打开 `OfficeAutoVersion.sln`（首次打开若提示迁移/升级项目，直接确认即可）。
   > Word / Excel 工程的 `ProjectProperties` 带完整 `HostPackage`，VS 能直接识别为 VSTO 加载项；
   > PowerPoint 工程的 `HostPackage` 留空（命令行构建不受影响，F5 也能启动）。
   > 若 VS 提示该工程“项目类型不受支持”，最省事的修法：在 VS 里新建一个空白的
   > “PowerPoint VSTO 外接程序”，把它 csproj 里 `<FlavorProperties GUID="{BAA0C2D2-...}">`
   > 那一整段 `ProjectProperties` 复制过来覆盖即可。
2. 右键想启动的加载项工程（Word/Excel/PowerPoint）→ **设为启动项目**。
3. 按 **F5**：VS 会自动生成清单、注册加载项并启动对应的 Office 应用程序。
4. 打开任意文档 → 编辑 → 保存（Ctrl+S）→ 文档旁边出现 `<名称>_Versions` 文件夹。
5. 功能区多出一个 **“自动版本”** 选项卡：开关 / 版本历史 / 打开版本文件夹 / 设置。
6. 点 **“版本历史...”** 打开列表：选中某个版本后点 **“比较”**（调用 Office 自带比较；
   窗口会先关闭再执行，比较结果随后出现），或点 **“恢复到此版本”** 回滚
   （恢复前会自动给当前状态存一份版本，恢复本身也可撤销）。

### 命令行构建

用 VS 自带的 MSBuild（路径按自己的安装位置调整）：

```bat
"D:\Tools\System\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" ^
  OfficeAutoVersion.sln -t:Build -p:Configuration=Release
```

> **关于清单签名**：VSTO 工程必须给 ClickOnce 清单签名才能生成 `.vsto`，
> 否则报 `Cannot build because the ClickOnce manifest signing option is not selected`。
> 三个 csproj 已指向当前用户证书库中的自签名证书（`CN=OfficeAutoVersion Dev`，
> 指纹 `2795C5C2B6BDD504FF696E2DD3643C231D9FEF71`），首次构建前若证书不存在，重新生成一次：
>
> ```powershell
> $c = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=OfficeAutoVersion Dev" `
>       -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5)
> $c.Thumbprint   # 把输出写进三个 csproj 的 ManifestCertificateThumbprint
> ```
>
> 也可以在 VS 里打开工程属性 → **签名** → 勾选“为 ClickOnce 清单签名”，让 VS 自动生成临时密钥。

### 运行核心逻辑自测（无需 Office）

```bat
dotnet run --project tests/OfficeAutoVersion.Selftest
```

### 本机当前的注册状态（可以直接试）

为了做真机验证，**Excel 加载项已注册到当前用户**（HKCU，无需管理员权限、不需要装 VS 也能用）：

```
HKCU\Software\Microsoft\Office\Excel\Addins\OfficeAutoVersion.Excel
  Manifest     = file:///…/src/OfficeAutoVersion.Excel/bin/Debug/OfficeAutoVersion.Excel.vsto|vstolocal
  LoadBehavior = 3
```

现在直接打开 Excel：功能区会多出 **“自动版本”** 选项卡，编辑并保存后文档旁就会出现 `xxx_Versions`。
想恢复干净状态，删掉这一个注册表键即可：

```powershell
Remove-Item "HKCU:\Software\Microsoft\Office\Excel\Addins\OfficeAutoVersion.Excel" -Recurse -Force
```

Word / PowerPoint 未注册，用 VS 打开解决方案按 **F5** 即可（VS 会自行注册，停止调试后自动撤销）。

## 7. 部署给他人使用

方式一（推荐，最简单）：在 VS 中右键加载项工程 → **发布** → 文件夹/共享路径生成 ClickOnce 包，
把生成的 `.vsto` 安装链接发给用户双击安装。

方式二：生成后把 `bin\Release` 输出目录（含 `.dll`、`.manifest`）按 VSTO 注册表项手动部署，
适合域环境批量装机。

开发机清理：VS 菜单 **生成 → 清理解决方案**，会同时移除注册表和信任设置。

## 8. 配置说明

配置按应用分别保存在
`%APPDATA%\OfficeAutoVersion\settings.json`，例如：

```json
{
  "Word": {
    "Enabled": true,
    "StorageMode": 0,
    "VersionFolderName": "{name}_Versions",
    "CentralRoot": "",
    "MaxVersionsToKeep": 20,
    "SkipIfUnchanged": true,
    "FileNamePattern": "{name}_{ts}",
    "LogEnabled": true
  }
}
```

| 设置 | 默认 | 说明 |
| --- | --- | --- |
| Enabled | true | 总开关（功能区上也能切） |
| StorageMode | 0 | 0=原文件旁边；1=集中目录 |
| VersionFolderName | `{name}_Versions` | 版本文件夹名，支持 `{name}` |
| CentralRoot | 空 | 集中模式根目录，空 = `文档\OfficeAutoVersions` |
| MaxVersionsToKeep | 20 | 保留数量，0=不限；超出的移入回收站 |
| SkipIfUnchanged | true | 内容没变就跳过 |
| FileNamePattern | `{name}_{ts}` | 支持 `{name}` `{ext}` `{ts}` 及 .NET 日期格式 |
| LogEnabled | true | 写运行日志 |

在功能区“自动版本 → 设置...”里可以可视化修改，并能**实时预览**当前文档的版本存放路径。

## 9. 自测与验证状态

**（1）单元自测：84 项断言全部通过**（`dotnet run --project tests/OfficeAutoVersion.Selftest`，.NET 8 真实编译运行）
覆盖快照生成、去重、数量清理、集中目录镜像、配置持久化与容错、文件名格式、边界情况、
功能区回调、日志、文件被其它进程占用时的去重、**版本历史窗体的列表与“比较”转交（含模态分支）**、
**设置窗口尺寸自适应**，以及三个加载项（含宿主动作类）的编译验证。

**（2）真机动作测试：18 项全部通过**（`tools/ActionHarness`，Word 2024 / Excel 2024 / PowerPoint 2024）
这个测试宿主用反射直接调用加载项里真实的 `Word/Excel/PowerPointVersionActions`，
因此验证的就是交付代码，而不是复刻一份逻辑：

```
=== Excel ===
  [PASS] 比较时版本以第二个窗口打开 / 已开启自带的“并排查看”同步滚动
  [PASS] 恢复后单元格内容等于所选版本 / 恢复前建立了安全快照 / 恢复只新增 1 个快照
  [PASS] 活动工作簿不是目标时：其它工作簿没有被关闭 / 内容未被改写
=== Word ===
  [PASS] 生成了带修订标记的对比文档（比较结果2，2 处修订）
  [PASS] 原文档在比较后仍然打开（未被误关闭）
  [PASS] 恢复后正文等于所选版本 / 恢复前建立了安全快照 / 恢复只新增 1 个快照
=== PowerPoint ===
  [PASS] 恢复后首屏标题等于所选版本（first）/ 恢复前建立了安全快照 / 恢复只新增 1 个快照
RESULT: ALL PASSED (18 checks)
```

另有 `ActionHarness.exe ui`：把设置窗口真实显示出来，逐个控件断言“完整可见”，并截图留证
（长中文路径场景，专门防“信息显示不全”复发）。

> 本次为空引用缺陷新增了 4 项回归断言（宿主为 null 时必须抛可读的 `InvalidOperationException`
> 而不是 `NullReferenceException`；同一动作实例在宿主“稍后就绪”后必须照常工作）。
> 这批断言需要在 VS 里重新编译解决方案 + `ActionHarness` 之后才能跑。

**（3）真机测试抓出并修掉的真实缺陷**（桩环境无法复现，只有真机才暴露）：

- **比较/恢复必现 `NullReferenceException`（用户实际遇到的“Excel 提示比较失败：未将对象引用设置到对象的实例”）**：
  根因是**功能区对象创建得比加载项初始化更早**。`CreateRibbonExtensibilityObject()` 里
  `new ExcelVersionActions(Application, ...)` 把当时的 `Application` 存进了字段，而那一刻
  `ThisAddIn.Application` 还没被赋值（null）。于是之后每次比较/恢复都在
  `_app.Workbooks.Open(...)` / `_app.Documents.Open(...)` 上抛空引用——
  真机测试为什么没抓到？因为它是在工具里 `new` 动作类、传的是现成的 Application，绕过了这个时机差。
  修复：三个动作类都改成保存 **`Func<Application>` 委托**，调用时才取宿主对象；
  仍未就绪时抛出可读的“Excel 还没准备好……请重启 Excel”，而不是空引用。
  同时功能区创建处记一条日志（`功能区已创建，Application=ok/null`），下次一眼就能确认时机。
- **恢复时“文件正由另一进程使用”**：Office 关闭文档后仍会短暂持有文件句柄（Word 还有 `~$` 所有者文件），
  原来只重试 4×80ms 就放弃。修复：新增 `VersionManager.WaitForFileUnlocked`（最多等 3 秒，句柄一放开立即返回），
  覆盖复制重试放宽到 12 次；并且**覆盖失败时把原文件重新打开**，避免“文档被关掉、版本又没恢复成”的双重损失。
- **模态窗口期间 Office 对象模型不可靠（用户实际遇到的“比较失败：当前没有打开的工作簿”）**：
  版本历史是模态窗口，显示期间 Excel 的 `ActiveWorkbook` 返回 null，
  恢复也就一样失败。修复分两层：
  ① 窗口只负责“选择”，真正的比较 / 恢复在窗口关闭后由功能区执行（不再处于模态状态）；
  ② 各应用的动作类改为**按传入的文件路径**在 `Workbooks`/`Documents`/`Presentations`
  集合里定位宿主对象（找不到还会自动打开），完全不再依赖 `ActiveWorkbook` 等属性——
  这同时修掉了“活动文档不是目标文档时会误操作别的文件”的隐患（已有真机回归用例）。
- **Word 的比较会误关用户正在编辑的文档**：`Documents.Open` 对已打开的文档返回同一个
  `Document` 对象，原来“重新打开当前文档做对比”再关闭临时文档，等于把用户的文档关了。
  已改为直接用当前打开的文档当“修订版”（回归检查已加进真机测试）。
- **Excel 的“同步滚动”是用户级持久设置**：用户曾经关掉过，比较时两表就不会一起滚。
  已在比较时显式 `SyncScrollingSideBySide = true`（这个坑也是真机测试跑出来的：第二次运行才失败）。
- 事件委托签名、`FileShare` 去重、`ICachedDataProvider` 参数（见上一轮记录）；
  这次新增代码还纠正了 `CompareSideBySideWith` 位于 `Windows` 集合（而非 `Window`）。

> 重跑方式：先编 `OfficeAutoVersion.sln`，再编 `tools/ActionHarness/ActionHarness.csproj`，
> 运行 `ActionHarness.exe`（可传 `excel` / `word` / `ppt` / `ui` 只跑其一）。
> 保存→生成版本那条链路另用 `tools/smoketest-excel.ps1` 验证。

## 10. 常见问题（FAQ）

**Q：比较 / 恢复弹出“未将对象引用设置到对象的实例”（NullReferenceException）？**
A：这是**旧版本**的缺陷：功能区创建得比加载项初始化早，动作类抓到的是 null 的 `Application`。
已改为“调用时才取 Application”。请用 VS 重新生成解决方案后**重启 Office** 再试；
若仍报错，看 `%APPDATA%\OfficeAutoVersion\logs\autoversion.log`：
里面的 `[ERROR] 比较失败` 会带上完整堆栈，新的启动日志还会写 `功能区已创建，Application=ok/null`。

**Q：恢复时提示“文件正由另一进程使用”？**
A：Office 关闭文档后仍会短暂持有句柄，插件现在会等句柄释放（最多 3 秒）再覆盖；
若仍失败，会自动把原文档重新打开，文件内容保持恢复前的状态，不会丢。

**Q：构建时提示找不到 `Office` / `Microsoft.Office.Interop.Word` 等引用？**
A：说明没装 VSTO 工作负载。VS 安装器 → 修改 → 工作负载 → 勾选
“使用 C# 的桌面开发” + “Office/SharePoint 外接程序开发”，会自动带上 Office PIA。

**Q：保存了，但文档旁边没有出现版本文件夹？**
A：按顺序检查三件事：
① 功能区“自动版本”选项卡里的**开关是否打开**（配置保存在
`%APPDATA%\OfficeAutoVersion\settings.json`，按应用独立，Word 关了不影响 Excel）；
② 文档是否保存到磁盘（未保存过的新文档没有“上一次保存”可快照）；
③ 日志 `%APPDATA%\OfficeAutoVersion\logs\autoversion.log`（里面会写明跳过原因；
加载项每次启动也会记一条“已启动”，用来区分“没加载”和“事件没触发”）。

**Q：开启了 OneDrive/SharePoint 的“自动保存”（AutoSave），版本会很多？**
A：会。自动保存同样会触发保存事件，而且 **Office 每次保存都会重写整个文件**
（内部时间戳、修订号等都会变），所以按字节比对的去重几乎不会命中，等于每次自动保存都留一份快照。
建议把“最多保留版本数”调小（默认 20），或临时用功能区开关关掉。

**Q：文档在“受保护的视图”里打开，为什么没有版本？**
A：受保护视图会禁用所有加载项，点击“启用编辑”后即恢复正常。

**Q：只读打开的文档、从邮件直接打开的附件？**
A：只读文档不会触发保存事件；附件在临时目录中打开，版本会生成在临时目录里，随系统清理消失。

**Q：版本文件夹会占多少空间？会同步到 OneDrive 吗？**
A：每个版本都是一份完整拷贝，空间 ≈ 文档大小 × 版本数（可设置上限）。
`xxx_Versions` 就在文档旁边，会随文档一起同步/拷贝——这是默认设计的取舍
（版本跟着文件走）；如果介意，可切换到“集中目录”模式。

**Q：如何彻底卸载？**
A：控制面板卸载（ClickOnce 安装的），或 VS 里清理解决方案；再删除
`%APPDATA%\OfficeAutoVersion`（配置与日志）即可，版本文件夹随文档自行处理。

## 11. 已知限制 / 可扩展点

- 仅支持 Windows 桌面版 Office 2013+。
- 去重是“整份文件 MD5”级别：Office 每次保存都会重写文件元数据，所以对 Word/Excel/PowerPoint
  实际几乎不会命中。它只作为兜底（避免明显重复、并让“文件被占用”时不再误判），
  不能指望用它压缩版本数量——版本数量请用“最多保留版本数”控制。
- 版本是整份拷贝，没有做增量/差异存储（文档级场景下整份拷贝最简单可靠）。
- **恢复是文件级替换**：会先保存当前文档、再关闭并重新打开。因此
  ① 文档若正被他人/另一个进程占用（如共享盘协同编辑），恢复可能失败并提示；
  ② 恢复后撤销栈会清空（因为重新打开了文件）；
  ③ 恢复的是“磁盘文件”，文档里的自定义状态（如 Excel 的数据连接会话）会重新初始化。
- **Excel 的比较是“并排查看”，不是内容级 diff**：没有逐单元格差异高亮。
  原因：Office 2024 已移除 Spreadsheet Compare 独立工具，Excel 本身也没有比较 API。
  想要逐格差异，需要自己解析 xlsx（或装第三方工具），属于另一个量级的工作。
- **PowerPoint 的“比较”需要在真实界面中用**：它走 `Presentation.Merge`，会打开交互式
  “审阅”窗格，无法在无界面环境自动化验证（真机测试里只验证了 PPT 的恢复链路）。
- Word 比较以**当前打开的文档**为“修订版”：若文档有未保存修改，比较的是修改后的状态；
  想比较磁盘上的版本，先保存一次即可。
- Outlook/Visio/Project 可照抄同模式扩展：新建加载项工程 → 挂接保存事件 → 复用 Core/Common，
  并实现一个自己的 `IVersionHostActions`。

## 12. 附：VBA 快速替代方案（不想装 VS 时）

不想部署插件的话，用 VBA 也能实现同样效果。以 Word 为例（Excel 类似）：

1. `Alt+F11` 打开 VBA 编辑器，**插入 → 类模块**，命名 `CAutoVersion`：

```vba
Public WithEvents App As Word.Application

Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    Dim fso As Object, folder As String, name As String, ext As String
    If Len(Dir(Doc.FullName)) = 0 Then Exit Sub
    name = Left(Doc.Name, InStrRev(Doc.Name, ".") - 1)
    ext = Mid(Doc.Name, InStrRev(Doc.Name, "."))
    folder = Doc.Path & "\" & name & "_Versions"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    FileCopy Doc.FullName, folder & "\" & name & "_" & Format(Now, "yyyymmdd-hhmmss") & ext
End Sub
```

2. **插入 → 标准模块**，粘贴：

```vba
Public AppEvents As New CAutoVersion

Sub AutoExec()
    Set AppEvents.App = Word.Application
End Sub
```

3. 保存为 `Normal.dotm`（或放进 startup 目录），重启 Word 即生效。

> VBA 版是精简演示：没有去重、数量上限和日志，也不需要功能区开关。
> 正式使用建议还是用本插件的 VSTO 版本。
