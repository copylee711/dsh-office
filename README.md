# dsh-office

[![npm](https://img.shields.io/npm/v/@copylee/dsh-office)](https://www.npmjs.com/package/@copylee/dsh-office)

让 DeepSeek Harness 直接在你电脑上打开的 Word、Excel、PowerPoint 里改文档。AI 不改写文件，而是通过 Office 的自动化接口（COM）让应用自己执行修改，所以改动会实时出现在窗口里，你可以边看边一起编辑，格式不会因为换工具保存而走样。

仅支持 Windows，需要本机装有微软 Office。

## 功能

| 功能 | 说明 |
|---|---|
| 在打开的文档里改 | 文件正开着也能改（直接改文件的做法这时会因为文件被占用而失败）；文件已经打开就直接接着用，不会再开一份 |
| 实时可见，可以一起改 | 每一步修改即时出现在 Word / Excel / PowerPoint 窗口里。你可以同时在里面打字、调格式，AI 每次动手前读到的都是文档当下的样子 |
| 不需要窗口在前台 | 修改不靠鼠标键盘，Office 窗口被挡住或在后台也照常进行，你可以同时用电脑做别的事 |
| 批量编辑 | `office_edit` 一次调用按顺序执行一批修改（写一整篇文档、填一张表加公式加图表、做几页幻灯片），中途某一项失败就停下并报告已完成哪些 |
| 改错位置的防护 | Word 按“段落编号 + 这段开头的文字”定位。你刚改过导致编号对不上时，这一步会被拒绝并要求重新读取；改写和删除段落必须带开头文字 |
| Word 一次撤销 | AI 的一批修改在 Word 里是一个撤销步骤，按一次 Ctrl+Z 整批撤回 |
| 读结构 | Word：带编号的段落、标题级别、表格；Excel：工作表、单元格的值和公式；PowerPoint：每页的形状（名称、类型、位置、文字） |
| 看排版 | `office_render` 把 Word 的一页、PowerPoint 的一页或 Excel 的一个区域渲染成图片给模型核对，不依赖截图，窗口在后台也能出图 |
| 写错参数会被指出 | 操作里出现它不认识的字段时，结果里会点名哪个字段被忽略，而不是悄悄不生效 |
| 零安装依赖 | 首次使用时用 Windows 自带的 .NET Framework `csc.exe` 编译一个小 helper（约 1 秒，之后缓存），不需要 Python、LibreOffice 或 node-gyp |

## 安装

DSH 桌面版：**插件 → 添加插件**，输入 `@copylee/dsh-office`，安装后启用。

命令行（web 等其他 profile）：

```bash
dsh plugin --profile web add @copylee/dsh-office@latest
```

要求：Windows 10 / 11（自带 .NET Framework 4.8），微软 Office（Word、Excel、PowerPoint，装了哪个就能用哪个），DSH 0.2.0-rc.2 或更高。

## 使用

直接说要做什么，例如：

- 用 Word 写一份季度工作总结，保存到桌面
- 打开 D:\数据\销售.xlsx，加一列金额并在最后一行求和，再画一个柱状图
- 把我现在开着的 PPT 第 3 页的标题改短一点，正文改成三条要点

AI 干活的同时你可以直接在窗口里改。你正在单元格里输入、或者开着对话框时，Office 不接受外部修改，AI 会等几秒再试。

AI 只有在调用保存时才写入文件，不会替你关闭文档。

## Agent 工具

| 工具 | 作用 |
|---|---|
| `office_open` | 在 Word / Excel / PowerPoint 里打开文件；路径不存在则新建；已打开则直接使用 |
| `office_status` | 列出三个应用里现在打开了什么、哪些有未保存的修改 |
| `office_read` | 把打开的文档读成文本（段落 / 单元格 / 形状） |
| `office_edit` | 按顺序执行一批修改，不保存 |
| `office_render` | 把一页 / 一张幻灯片 / 一个单元格区域渲染成图片 |
| `office_save` | 保存；给出路径时另存为，扩展名为 `.pdf` 时导出 PDF |

`office_edit` 支持的操作：

- Word：`insert_paragraphs`、`set_text`、`replace_text`、`format_text`、`delete_range`、`insert_table`、`set_cell`、`insert_image`
- Excel：`write_range`（值或公式）、`format_range`、`autofit`、`insert_rows`、`delete_rows`、`add_sheet`、`rename_sheet`、`delete_sheet`、`add_chart`
- PowerPoint：`add_slide`、`set_text`、`add_textbox`、`add_image`、`set_shape`、`delete_shape`、`delete_slide`、`move_slide`、`set_notes`

## 设置

设置 → Office。每项改动自动保存。

| 设置 | 默认 | 说明 |
|---|---|---|
| 打开文档时显示到前台 | 开 | 关闭后文档仍会打开，只是不抢当前窗口 |
| 核对排版的图片宽度 | 1100 像素 | `office_render` 给模型的图片宽度；越大越清晰，也越费 token |

设置页还会显示三个应用是否安装、是否在运行、打开了哪些文档。

## 限制

- Excel 和 PowerPoint 里 AI 的修改无法用 Ctrl+Z 撤回（Excel 的外部修改不进撤销栈，并且会清空已有的撤销记录）。只有 Word 支持整批撤销。
- 你正在单元格里输入、或 Office 开着模态对话框时，修改会被 Office 拒绝；插件等待约 5 秒后报告“应用忙”。
- Excel 区域的图片是通过剪贴板取得的，渲染时剪贴板里的文字会被保留，图片、文件等其他内容会丢失。
- 目前不支持 WPS。
- 没有打开 Office 时处理文件、批量转换格式不在本插件范围内。

## 数据与隐私

插件只在本机运行，不联网。它读写的是你让 AI 处理的文档；`office_render` 生成的图片作为附件保存在会话历史里，和 DSH 的其他图片附件一样。

## 工作原理

宿主侧是一个 Cordis 插件，注册上面六个工具。它启动一个常驻的小程序（`helper/OfficeHelper.cs`，C# 5，运行时由系统自带的 `csc.exe` 编译，按源码哈希缓存在 `%LOCALAPPDATA%\dsh-office\`），通过标准输入输出逐行交换 JSON。helper 用 COM 附着到正在运行的 Office 应用（没有运行则启动并显示出来），按后期绑定调用它们的对象模型。每条命令结束后 helper 释放全部 COM 引用，这样你关掉 Office 时进程能正常退出。

## 开发

```bash
pnpm install
pnpm run typecheck
pnpm test
pnpm run build
node scripts/office-smoke.mjs        # 真实 Office：只操作脚本自己在临时目录里新建的文档
```

## 许可证

MIT
