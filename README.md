# dsh-page-images

把消息附件里的文档**逐页变成图片**，作为一条新消息发回 DSH 会话——用户直接看到每一页，模型在同一个请求里拿到这些图像。

这个仓库放两个插件，它们共用同一套机制（`agent/pre-step` 拦截附件 → 渲染 → 经 `ctx.attachments` 存入附件库 → 追加一条带图片块的 user 消息），共享同一批测试骨架与安装方式，但各自是一个可独立安装、独立启停的包。

| 包 | 处理什么 | 说明 |
|---|---|---|
| [dsh-pdf-pages](packages/dsh-pdf-pages) | `.pdf` | PDFium 逐页栅格化；工具 `pdf_pages` |
| [dsh-slide-pages](packages/dsh-slide-pages) | `.pptx` / `.ppt` / `.odp` | 先导出 PDF 再栅格化（LibreOffice 直渲会裁切）；附整份 deck 文字大纲；工具 `slide_pages` |

## 共同点

两个插件都**不引入第三方依赖**：渲染引擎（`@deepseek-ai/libreoffice-kit`）、工具注册（`dsh-tools`）、消息构造（`dsh-llm`）、附件存储（`dsh-attachment`）全部随 DSH 发行版提供，通过 profile 的包解析器从当前安装解析。也因此换 DSH 版本不需要跟着升级依赖。

两者都只在有对应附件进入消息时自动工作；失败一律降级成一条说明原因的消息（文件损坏、超过引擎输入上限、超时、附件后端没有宿主路径），不会打断这一轮。

## 安装

任选一种，然后重启 DSH（profile 组合在启动时计算）。

**从 GitHub 直接装**（需要 DSH 支持 `github:` 依赖，profile 里已有同类用法）：

```sh
dsh plugin --profile desktop add github:LinYonghang/dsh-page-images#path:/packages/dsh-pdf-pages
dsh plugin --profile desktop add github:LinYonghang/dsh-page-images#path:/packages/dsh-slide-pages
```

**从本地 tarball 装**：

```sh
git clone https://github.com/LinYonghang/dsh-page-images
cd dsh-page-images && pnpm -r pack          # 每个包出一个 tgz
dsh plugin --profile desktop add ./packages/dsh-pdf-pages/dsh-pdf-pages-0.1.0.tgz
dsh plugin --profile desktop add ./packages/dsh-slide-pages/dsh-slide-pages-0.1.0.tgz
```

装完记得把包名加进 profile `package.json` 的 `dsh.profile.bundles`（插件管理页会替你写）。

> ⚠️ 不要把包**直接拷进** `node_modules` 就完事：profile 里任何一次 `pnpm install`（装别的插件时就会触发）都会把没写进 `dependencies` 的目录清掉，插件随之消失。用上面两种方式，`file:` 或 `github:` 依赖会进 lockfile，才不会被 prune。

## 配置

两个插件的行 id 分别是 `pdf-pages`、`slide-pages`，在 profile 的 `cordis.patch.yml` 里按行覆盖：

```yaml
- id: pdf-pages
  config:
    autoConvert: true      # 关掉就只在被调用时转换
    autoMaxPages: 10       # 自动转换的页数上限
    dpi: 110

- id: slide-pages
  config:
    autoConvert: true
    autoMaxSlides: 8
    outline: true          # 附带逐页标题/正文/演讲备注
    dpi: 110
```

完整的字段表见各包 README。

## 实测代价与保真度

- PDF：每页约 0.25 秒（PDFium），每次渲染约 2 秒引擎启动。
- 幻灯片：导出 PDF 是固定大头（37 张 / 2 MB 实测 9 秒），之后每张约 0.25 秒；同一 deck 的追加请求命中缓存。
- 幻灯片与 PowerPoint 自己导出的版本逐页对比：均方误差 3.2–12.2 / 255、可视差异像素 2.5%–9.6%，差异集中在字符边缘；字体替换、动画、SmartArt、视频、3D 取静态帧。
- 引擎限制：单文件输入 64 MiB，输出 128 MiB，插件把超时提到 300 秒。

## 测试

测试对着**真实的引擎与真实的附件存储**跑，不是 mock：

```sh
# 在包目录里，先让 @deepseek-ai/* 可解析（开发用联接，别提交）
New-Item -ItemType Junction -Path .\node_modules `
  -Target 'C:\Users\<you>\AppData\Local\Programs\DSH NEXT\resources\app.asar.unpacked'

node packages/dsh-pdf-pages/test/smoke.mjs   /abs/path/sample.pdf
node packages/dsh-pdf-pages/test/edges.mjs   /abs/path/sample.pdf
node packages/dsh-pdf-pages/test/timing.mjs  /abs/path/sample.pdf

node packages/dsh-slide-pages/test/smoke.mjs   /abs/path/deck.pptx
node packages/dsh-slide-pages/test/edges.mjs   /abs/path/deck.pptx
node packages/dsh-slide-pages/test/outline.mjs /abs/path/deck.pptx
node packages/dsh-slide-pages/test/timing.mjs  /abs/path/deck.pptx
```

## 目录

```
packages/
  dsh-pdf-pages/     package.json  cordis.patch.yml  lib/{index,render}.js  test/  README.md
  dsh-slide-pages/   package.json  cordis.patch.yml  lib/{index,render,outline}.js  test/  README.md
```

两个 `lib/index.js` 目前有约两百行是近乎重复的机制代码（pre-step 注入、附件分批入库、引用跟踪、失败降级）。现在刻意保持重复，好让两个包能独立安装、独立演进；真出现第三个格式（`.docx`/`.xlsx`）或两侧机制开始分叉时，再在同一仓库里抽一个 `packages/core`。

## 已知待办

- 幻灯片大纲的标题合并：有些 deck 把编号与标题拆成两段（`1` + `、矩阵的转置`），大纲标题会变成《1》。
- 可选的 PowerPoint 原生渲染：本机 PowerPoint COM 实测 13 张 10.1 秒、像素级一致，可作为 `renderer: auto | powerpoint | libreoffice` 的第二路径。

## 许可

MIT，见 [LICENSE](LICENSE)。
