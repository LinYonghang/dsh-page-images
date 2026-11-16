# dsh-pdf-pages

把消息附件里的 PDF **逐页转换成图片**，并把这些图片**作为消息发回 DSH 会话**：用户能直接在对话里看到每一页，模型也会在同一个请求里拿到这些图像。

插件只有 Host 一半，没有客户端代码——图片走的是 DSH 现成的「用户消息附件」通道（`ctx.attachments` + `ImageBlock`），因此它们会像用户自己上传的图片一样被持久化、在历史里回放、在被压缩时正常保留。

## 工作方式

1. **自动转换（默认开启）**：在 `agent/pre-step` 阶段检查即将进入这一步的消息。只要里面出现 PDF 文件块（`{ type: 'file', attachment: { name: '*.pdf' } }`），插件就用 DSH 自带的 PDFium 渲染器（`@deepseek-ai/libreoffice-kit`）把页面栅格化成 PNG，逐张存入附件库，然后**在同一批消息后面追加一条 user 角色消息**，正文是一行说明加上每页的图片块。
   - 消息会作为普通的 `user/message` 事件落盘，模型在同一轮就能读到，Web UI 里也直接显示为一条带图片的消息。
   - 任何一个 PDF 转换失败（文件损坏、后端没有宿主路径等）只会追加一条说明失败原因的文字消息，绝不会打断这一轮。
2. **按需转换（`pdf_pages` 工具）**：模型（或你要求模型）可以指定页码、DPI、页数上限，重新转换或继续转换长文档。工具结果的正文同样带图片块。

图片在请求里消耗的 token 与像素数成正比，所以自动转换默认只转前 10 页、110 dpi；其余页面通过工具按需补转。

## 安装

插件依赖 `@deepseek-ai/libreoffice-kit`、`@deepseek-ai/dsh-tools`、`@deepseek-ai/dsh-llm` 等**随 DSH 发行版提供**的包，profile 的包解析器会让它们从安装目录解析，所以插件本身不需要安装任何第三方依赖。

```sh
# 1. 打包（在插件目录内）
npm pack                     # 得到 dsh-pdf-pages-0.1.0.tgz

# 2. 装进某个 profile（示例：desktop）
dsh plugin --profile desktop add ./dsh-pdf-pages-0.1.0.tgz

# 3. 把 bundle 登记进 profile 组合
#    编辑 ~/.dsh/profiles/desktop/package.json，在 dsh.profile.bundles 里加入 "dsh-pdf-pages"
```

也可以在插件市场/插件管理页里以本地 bundle 的方式安装（等价于上面三步）。安装后**重启 DSH**（profile 组合在启动时计算），插件的行会出现在插件列表里，工具 `pdf_pages` 会出现在模型的工具表里。

## 配置

配置写在 profile 的 `cordis.patch.yml` 里，按行 id 覆盖：

```yaml
- id: pdf-pages
  config:
    autoConvert: true      # false = 只在被调用时转换
    autoMaxPages: 10       # 自动转换的页数上限
    autoMaxAttachments: 2  # 一步里最多自动处理几个 PDF 附件
    dpi: 110               # 栅格密度（48–300）
    maxPagesPerCall: 24    # 单次工具调用的页数上限
    timeoutMs: 180000      # 单次渲染的超时
```

## 使用

- 直接把 PDF 拖进输入框发送。下一轮消息里就会出现「PDF 附件已逐页转换为图片：xxx 的第 1–10 页」以及这些页面。
- 让模型看某几页：`用 pdf_pages 把刚上传的 PDF 第 12-15 页按 150 dpi 转出来`。
- 也可以直接给路径：`pdf_pages` 的 `source` 支持绝对 `.pdf` 路径，因此磁盘上的 PDF 同样可以转图进对话。

## 性能与限制

- 渲染走 PDFium：每次渲染约 2 秒启动 + 每页 0.15–0.9 秒（A4、110 dpi 时约 0.3 秒/页，含 PNG 编码与落盘），首次运行还要建一次字体索引。
- 一条消息里的图片数量受部署的图片准入限制（默认每消息 20 张），插件会自动分批写入附件库；单次工具调用可用页数另有 `maxPagesPerCall` 上限。
- 单个附件必须是 PDF 且由本机文件支撑（`ctx.attachments.fileHostPath()` 可用）。不支持的部署会得到一条说明性的失败消息，而不是崩溃。
- 图片受字体、编码影响可能与阅读器略有差异；纯图像型 PDF 同样可以转（走的是页面栅格化，不需要文本层）。

## 验证

`test/` 里有三组脚本，直接对着真实的附件存储与真实的 PDFium 渲染器跑：

```sh
node test/smoke.mjs  /abs/path/sample.pdf   # 自动注入 + 工具 + 页码越界
node test/edges.mjs  /abs/path/sample.pdf   # 关闭自动转换、坏 PDF、无宿主路径、路径来源、错误提示
node test/timing.mjs /abs/path/sample.pdf   # 每页渲染耗时
```

脚本需要能解析到随 DSH 提供的 `@deepseek-ai/*` 包。最省事的做法是在插件目录里建一个指向 DSH 安装目录的 `node_modules` 联接（只在开发时用，不要随插件分发）：

```powershell
New-Item -ItemType Junction -Path .\node_modules `
  -Target 'C:\Users\<you>\AppData\Local\Programs\DSH Desktop\resources\app\node_modules'
```
