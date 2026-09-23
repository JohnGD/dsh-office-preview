# dsh-office-preview

Word（`.docx`）与 Excel（`.xlsx`）的**右侧栏文档预览**插件，为 DeepSeek Harness 补齐
`@deepseek-ai/dsh-client-ui-sidebar-documentpreview` 目前只覆盖文本 / 代码 / 图片 / PDF / HTML
的那一块空白：官方预览器遇到 Office 文件时会落到纯文本渲染器，于是显示
「非文本文件，暂时无法预览」。

本插件不调用 Office、不启动外部程序、不联网，也不写任何文件：它在浏览器里把
OOXML 容器（ZIP）解开，读取 `word/document.xml`、`xl/workbook.xml` 等部件，
直接渲染成 HTML。

## 支持的格式

| 扩展名 | 渲染 |
|---|---|
| `docx` `docm` `dotx` `dotm` | 标题层级、段落（保留缩进与项目符号）、加粗/斜体/下划线、表格（含嵌套表格） |
| `xlsx` `xlsm` `xltx` `xltm` | 工作表分页签、行列表头、合并单元格、列宽、加粗/对齐、日期与时间格式、公式的缓存值 |

不支持的旧二进制格式（`.doc`/`.xls`）不会被接管，仍走官方纯文本回退；`.pptx` 目前不在范围内。

## 实现

- **宿主半**（`lib/index.js`）：空的 `apply()`。宿主侧不需要任何服务，它的唯一作用是让 Cordis
  loader 里存在一个声明了 `dsh.client` 的条目，模块系统据此把浏览器半打进 boot graph。
- **浏览器半**（`lib/client.js`）：单文件、零依赖（只 `require("react")`），向官方两个扩展点注册：
  - `ctx.documentPreviews.register({ id, extensions, priority: "external", loading: "bytes-complete" })`
  - `ctx.slots.register` 到 `sidebar.right.tab.document`，key 与上面的 `id` 相同
  于是官方文档面板读完全部字节后把 `content.data` 交给本渲染器；面板顶部的「打开方式」里
  会多出「Word 文档 / Excel 表格」两个选项，随时可以切回纯文本。
- **自带解析器**：`lib/client.js` 内含一个小型 XML 读取器与 ZIP 中央目录读取器
  （`DecompressionStream("deflate-raw")` 解压），不引入 JSZip / SheetJS / mammoth，
  也不修改 Harness 任何源码或官方包。

## 安装

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

（若系统禁用了脚本执行，也可直接用 `-ExecutionPolicy Bypass`，或手工执行脚本里的三步。）

脚本只做三件事（`-HarnessHome` 可覆盖默认的 `%APPDATA%\dsh-desktop\harness`）：

1. 复制包到 `<home>\plugins\dsh-office-preview`；
2. 建立 `<home>\profiles\node_modules\dsh-office-preview` 联结（junction）指向它 ——
   这是 Cordis loader 解析插件条目名所用的回退根目录；
3. 在 profile 自己的补丁层 `<home>\profiles\web\cordis.patch.yml` 追加一条 `insert` 行
   （改动前自动备份为 `.bak-<时间戳>`）。

profile 的 `patchReload` 为 `live`，补丁层文件被监听，插件通常**不需要重启**即可挂载；
刷新 Harness 页面（Ctrl+R）即可加载新的客户端 bundle。若界面未出现变化，重启 DSH Desktop 一次即可。

### 它出现在哪里

- **Harness 设置 → 插件清单**：Cordis Loader 的插件清单会列出 `dsh-office-preview`（活动插件），
  与其他内置插件并列，可在这里确认它已加载。
- **市场（插件市场）的「已安装」列表**：只有把它装成 profile 依赖（`dsh plugin add`，或以
  `link:` 依赖形式登记）时才会出现，标记为「本地开发」，可在市场里启用/禁用/卸载。
  用上面的 `install.ps1` 安装（补丁层 + 回退目录链接）则不在市场列表里——好处是完全不触碰
  profile 的依赖、锁文件与代际投影。两种形态不能同时启用：同一条目 id 挂两次会冲突。
- **插件市场「发现」页**：本插件已发布到社区精选目录 `awesome-dsh-plugin`（分类 UI 增强），
  安装包是 GitHub Release 上的预构建 tarball。**本包不发 npm**——npm 会把发布者注册邮箱写进
  每个包的公开元数据（`maintainers[].email`），这是作者刻意规避的。

## 卸载

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1                # 保留 <home>\plugins 下的文件
powershell -ExecutionPolicy Bypass -File uninstall.ps1 -RemoveFiles   # 连文件一起删除
```

若插件是以 profile 依赖装的，请在市场 → 已安装 → 卸载里移除它。

## 自检

```powershell
$node = "$env:APPDATA\dsh-desktop\harness\.desktop-bin\node.cmd"
& $node test\parse.mjs samples\sample.xlsx samples\sample.docx   # 解析器：真实样本
& $node test\render.mjs samples\sample.xlsx samples\sample.docx  # 渲染路径：空表/越界合并等边界
& $node test\plugin.mjs                                          # 客户端注册契约
```

## 边界

- 只读：不写回文档，也不允许通过预览修改任何内容。
- 图片、批注、修订痕迹、图表、宏不渲染，其文字内容会按纯文本保留。
- 单个文件读取上限取决于 Harness 的 `workspaceFiles.readAll`（默认 32 MiB）。
- 大表格最多渲染 4000 行 × 160 列，超出部分给出提示。

## License

MIT
