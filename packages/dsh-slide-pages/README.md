# dsh-slide-pages

把消息附件里的演示文稿（`.pptx` / `.ppt` / `.odp`）**逐张幻灯片转换成图片**，并把这些图片**作为消息发回 DSH 会话**：用户能直接在对话里看到每一张幻灯片，模型也会在同一个请求里拿到这些图像，同时还附带整份 deck 的文字大纲。

## 为什么是两段式

`@deepseek-ai/libreoffice-kit` 有两条出图路线，两者都实测过：

| 路线 | 结果 |
|---|---|
| 直接栅格化演示文稿 | **不可用**：内容被按约 1.5 倍放大后裁切，右侧标题与表格列丢失（两份 deck 都能复现） |
| 先导出 PDF，再栅格化 PDF | **可用**：标题、页码、表格、结论框完整，与 PowerPoint 导出的 PDF 一致 |

所以本插件的管线固定为：`deck →（LibreOffice convert）→ PDF →（PDFium）→ 每页 PNG`。导出的 PDF 会按内容缓存，因此"看第 12–15 张"这种追加请求不必再付一次导出成本。

## 工作方式

1. **自动转换**：`agent/pre-step` 阶段检查即将进入这一步的消息，出现演示文稿文件块就用上面的管线把前 N 张（默认 8 张）转成图片，经 `ctx.attachments` 存成持久附件，然后在同一批消息后**追加一条 user 角色消息**：说明 + 全 deck 文字大纲 + 每张幻灯片的 image 块。这条消息作为普通的 `user/message` 落盘，模型同一轮可见，Web UI 里就是一条带图片的消息。
2. **文字大纲**：用 `fflate` 直接读 pptx 的 ZIP/XML（`ppt/slides/slideN.xml` 与 `notesSlides`），取出每页标题、正文、演讲备注。所以即使只转了 8 张图，模型也知道整份 deck 讲了什么；旧版二进制 `.ppt` 没有这份大纲。
3. **按需转换（`slide_pages` 工具）**：指定页码、DPI、页数上限，继续转后面的页；也接受磁盘上的绝对路径。

失败一律降级为一条说明原因的文字消息（损坏/加密、超过 64 MiB 输入上限、超时、后端没有宿主路径），绝不会打断这一轮。

## 安装

插件依赖 `@deepseek-ai/libreoffice-kit`、`@deepseek-ai/dsh-tools`、`@deepseek-ai/dsh-llm` 等**随 DSH 发行版提供**的包，profile 的包解析器会让它们从安装目录解析；`fflate` 也是随发行版提供的。所以插件本身不需要安装任何第三方依赖。

```sh
npm pack                     # 得到 dsh-slide-pages-0.1.1.tgz
dsh plugin --profile desktop add ./dsh-slide-pages-0.1.1.tgz
# 再把 "dsh-slide-pages" 加入 ~/.dsh/profiles/desktop/package.json 的 dsh.profile.bundles
```

或者直接从仓库装（会进 profile 的 lockfile，之后的 `pnpm install` 不会把它清掉）：

```sh
dsh plugin --profile desktop add github:LinYonghang/dsh-page-images#path:/packages/dsh-slide-pages
```

安装后重启 DSH（profile 组合在启动时计算）。工具 `slide_pages` 会出现在模型的工具表里，插件的行会出现在插件列表里。

## 配置

```yaml
- id: slide-pages
  config:
    autoConvert: true      # false = 只在被调用时转换
    autoMaxSlides: 8       # 自动转换的张数上限
    autoMaxAttachments: 2  # 一步里最多自动处理几个 deck
    dpi: 110               # 48–300
    maxSlidesPerCall: 24   # 单次工具调用上限
    outline: true          # 是否附带逐页文字大纲
    outlineBodyChars: 300  # 每页正文摘要上限
    outlineNotesChars: 160 # 每页演讲备注上限
    outlineTotalChars: 8000# 大纲总长度上限
    timeoutMs: 300000      # 导出 + 栅格化总超时
```

## 使用

- 把 `.pptx` 拖进输入框发送：下一轮消息里就会出现"演示文稿《…》第 1–8 张已转换为图片"、整份大纲和这 8 张幻灯片。
- 看后面几页：`用 slide_pages 把第 9-16 张按 150 dpi 转出来`。
- 磁盘上的 deck 也可以：工具 `source` 接受绝对路径。

## 性能与限制

- **导出一份 deck 约 6–9 秒**（37 张 / 2 MB 实测 9.0 秒，含引擎启动），之后**每张约 0.25 秒**；同一 deck 的追加请求命中 PDF 缓存，只付栅格化成本。
- 一条消息里的图片数量受部署的图片准入限制（默认每消息 20 张），插件自动分批写入附件库。
- 引擎输入上限 64 MiB、输出上限 128 MiB；超限会得到说明性失败消息。默认超时已提到 300 秒。
- 版式由 LibreOffice 渲染：**字体替换、动画、SmartArt、视频、3D 只取静态帧**，与 PowerPoint 显示可能存在差异；消息里会附一行同样的提示，缺字体时还会点名。
- 宏格式 `.pptm` 会先在同构的 OOXML 容器下复制成 `.pptx` 再交给引擎（引擎按扩展名分派）。
- 旧版二进制 `.ppt` 与 LibreOffice 导出的 `.odp` 走同一条管线；本机没有 `.ppt` 样张，这条标为"已实现、待样张验证"。

## 验证

`test/` 里四组脚本，对真实引擎与真实附件库跑（`outline.mjs` 用合成 deck，不需要真文件）：

```sh
node test/smoke.mjs   /abs/path/deck.pptx   # 自动注入（含大纲）+ 工具 + 越界页码
node test/edges.mjs   /abs/path/deck.pptx   # 关闭自动、坏 deck、无宿主路径、路径来源、.pptm
node test/outline.mjs [/abs/path/deck.pptx] # 大纲分段/实体/页码标记/备注/预算截断
node test/timing.mjs  /abs/path/deck.pptx   # 导出 / 栅格化耗时
```

脚本需要能解析到随 DSH 提供的包。开发时在插件目录里放一个指向 DSH 安装目录 `node_modules` 的联接即可（不要随插件分发）：

```powershell
New-Item -ItemType Junction -Path .\node_modules `
  -Target 'C:\Users\<you>\AppData\Local\Programs\DSH Desktop\resources\app\node_modules'
```
