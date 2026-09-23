# dsh-office-preview

**English** · [中文](#中文说明)

Word (`.docx`) and Excel (`.xlsx`) previews for the DeepSeek Harness sidebar document pane — read-only, no Office installation, no network.

Without this plugin, opening a `.docx` or `.xlsx` in the sidebar falls through to the plain-text renderer and shows *"Not a text file, preview is unavailable for now."*

## What it renders

| Extension | Rendered |
|---|---|
| `.docx` `.docm` `.dotx` `.dotm` | Headings, paragraphs with indentation and bullets kept, bold / italic / underline, tables including nested ones |
| `.xlsx` `.xlsm` `.xltx` `.xltm` | Sheet tabs, row and column headers, merged cells, column widths, number and date formats, cached formula results |

Legacy binary formats (`.doc`, `.xls`) are **not** claimed by this plugin and still fall back to plain text. `.pptx` is out of scope today.

## How it works

- **No dependencies, no services.** The plugin registers two renderers on the official extension points `ctx.documentPreviews` and the `sidebar.right.tab.document` slot. It patches no Harness source, calls no external service, and its only import is `react`.
- **OOXML is parsed in the browser.** `lib/client.js` contains a ZIP central-directory reader (raw-DEFLATE through `DecompressionStream`) and a small XML reader; no JSZip, SheetJS or mammoth.
- **The host half is empty.** `lib/index.js` exists only so the Cordis loader carries a client entry; all visible behaviour lives in the browser bundle.

## Install

**From the plugin market (recommended).** In Harness: Settings → Plugin Market → Discover → search `dsh-office-preview` → Install → restart Harness when asked.

**From the prebuilt tarball (no market, no npm).**

```sh
dsh plugin --profile web add https://github.com/JohnGD/dsh-office-preview/releases/latest/download/dsh-office-preview.tgz
```

**Manual, offline** — `install.ps1` on Windows, `install.sh` on macOS and Linux:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

The script does three things:

1. copies the package to `<home>\plugins\dsh-office-preview`;
2. links `<home>\profiles\node_modules\dsh-office-preview` at it — a root the Cordis loader resolves plugin entry names through;
3. appends one `insert` row to the profile's own patch layer `<home>\profiles\web\cordis.patch.yml`, backing that file up first.

The patch layer is watched (`patchReload: live`), so a page refresh (Ctrl+R) usually activates the plugin; if nothing changes, restart DSH Desktop.

This package is **not published to npm**: npm writes the publisher's account e-mail into every published package's public metadata (`maintainers[].email`). Shipping a prebuilt tarball on a GitHub Release avoids that entirely.

## Where it shows up

- **Harness → Settings → Plugin inventory** lists `dsh-office-preview` as an active plugin, next to the built-in ones.
- **Plugin market → Installed** lists it only when it is installed as a profile dependency (`dsh plugin add`, or a `link:` dependency). There it is labelled "local development" and can be enabled, disabled or uninstalled. The `install.ps1` form — patch layer plus fallback link — does not appear in that list; its advantage is that it touches no profile dependency, lockfile or generation projection. The two forms must not be enabled at once: mounting one entry id twice collides.

## Uninstall

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1                # keeps the copied files
powershell -ExecutionPolicy Bypass -File uninstall.ps1 -RemoveFiles   # removes them as well
```

If the plugin was installed as a profile dependency, remove it from the market instead: Market → Installed → Uninstall.

## Self-check

```powershell
$node = "$env:APPDATA\dsh-desktop\harness\.desktop-bin\node.cmd"   # or any Node >= 20
& $node test\parse.mjs  samples\sample.xlsx samples\sample.docx   # parsers
& $node test\render.mjs samples\sample.xlsx samples\sample.docx   # render path, edge cases included
& $node test\plugin.mjs                                           # client registration contract
```

Each script prints `ALL CHECKS PASSED` on success. `samples/` carries a small workbook and a small document, so the checks run without any external file.

## Limits

- Read-only: a preview never writes to the document.
- Images, comments, tracked changes, charts and macros are not rendered; their text is still shown.
- The per-file read cap comes from Harness's `workspaceFiles.readAll` — 32 MiB by default.
- Large sheets are truncated at 4000 rows × 160 columns, with a notice.

## License

MIT

---

# dsh-office-preview（中文说明）

**中文** · [English](#dsh-office-preview)

DeepSeek Harness 右侧栏文档面板的 Word（`.docx`）与 Excel（`.xlsx`）预览插件 —— 只读、不依赖 Office、不联网。

没有这个插件时，在右侧栏点开 `.docx` / `.xlsx` 会落到纯文本渲染器，显示「非文本文件，暂时无法预览」。

## 渲染能力

| 扩展名 | 渲染内容 |
|---|---|
| `.docx` `.docm` `.dotx` `.dotm` | 标题层级；段落（保留缩进与项目符号）；加粗 / 斜体 / 下划线；表格（含嵌套表格） |
| `.xlsx` `.xlsm` `.xltx` `.xltm` | 工作表分页签；行列号表头；合并单元格；列宽；数字与日期格式；公式的缓存值 |

旧二进制格式（`.doc`、`.xls`）本插件**不接管**，仍走官方纯文本回退；`.pptx` 目前不在范围内。

## 实现方式

- **零依赖、无外部服务。** 插件只在官方扩展点 `ctx.documentPreviews` 与插槽 `sidebar.right.tab.document` 上注册两个渲染器；不修改 Harness 任何源码，不调用外部服务，唯一引入的模块是 `react`。
- **OOXML 在浏览器内解析。** `lib/client.js` 内置 ZIP 中央目录读取器（用 `DecompressionStream` 解 raw-DEFLATE）与一个小型 XML 读取器，不引入 JSZip、SheetJS 或 mammoth。
- **宿主半是空的。** `lib/index.js` 的唯一作用是让 Cordis loader 里存在这个条目，全部可见行为都在浏览器半。

## 安装

**从插件市场安装（推荐）。** 在 Harness 里：设置 → 插件市场 → 发现 → 搜索 `dsh-office-preview` → 安装 → 按提示重启 Harness。

**用预构建包安装（不用市场、不用 npm）。**

```sh
dsh plugin --profile web add https://github.com/JohnGD/dsh-office-preview/releases/latest/download/dsh-office-preview.tgz
```

**离线手动安装** —— Windows 用 `install.ps1`，macOS 与 Linux 用 `install.sh`：

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

脚本只做三件事：

1. 复制包到 `<home>\plugins\dsh-office-preview`；
2. 建立 `<home>\profiles\node_modules\dsh-office-preview` 联结指向它 —— 这是 Cordis loader 解析插件条目名所用的回退根目录；
3. 在 profile 自己的补丁层 `<home>\profiles\web\cordis.patch.yml` 追加一条 `insert` 行，改前自动备份该文件。

补丁层被 live 监听（`patchReload: live`），通常刷新页面（Ctrl+R）即可生效；若界面无变化，重启一次 DSH Desktop。

本包**不发 npm**：npm 会把发布者的注册邮箱写进每个已发布包的公开元数据（`maintainers[].email`）。改为把预构建包挂在 GitHub Release 上，可以完全避开这一点。

## 它出现在哪里

- **Harness → 设置 → 插件清单**：会列出 `dsh-office-preview` 处于活动状态，与内置插件并列。
- **插件市场 → 已安装**：只有把它装成 profile 依赖（`dsh plugin add`，或 `link:` 依赖）时才会出现，标记为「本地开发」，可在市场里启用 / 禁用 / 卸载。用 `install.ps1` 装的形式（补丁层 + 回退目录链接）不在该列表中；它的好处是完全不触碰 profile 依赖、锁文件与代际投影。两种形式不能同时启用：同一条目 id 挂两次会冲突。

## 卸载

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1                # 保留已复制的文件
powershell -ExecutionPolicy Bypass -File uninstall.ps1 -RemoveFiles   # 连文件一起删除
```

若插件是以 profile 依赖装的，请在市场 → 已安装 → 卸载里移除它。

## 自检

```powershell
$node = "$env:APPDATA\dsh-desktop\harness\.desktop-bin\node.cmd"   # 或用任意 Node >= 20
& $node test\parse.mjs  samples\sample.xlsx samples\sample.docx   # 解析器
& $node test\render.mjs samples\sample.xlsx samples\sample.docx   # 渲染路径（含边界用例）
& $node test\plugin.mjs                                           # 客户端注册契约
```

三个脚本成功时都会打印 `ALL CHECKS PASSED`；`samples/` 里带了一个小工作簿与一个小文档，不需要外部文件就能跑。

## 边界

- 只读：预览不会写回文档。
- 图片、批注、修订痕迹、图表、宏不渲染，其文字内容仍会显示。
- 单文件读取上限取决于 Harness 的 `workspaceFiles.readAll`，默认 32 MiB。
- 大表格最多渲染 4000 行 × 160 列，超出部分给出提示。

## 许可

MIT
