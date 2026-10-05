# 文件管理与知识检索平台

基于 Docker Compose 的文件管理、分类和知识检索平台。

## 环境要求

- Windows 10/11（64 位）
- Docker Desktop 4.x 或更高版本
- Docker Compose V2（命令：`docker compose`）
- WSL2，Docker Desktop 使用 Linux 容器
- 可用内存至少 4 GB
- 可用磁盘空间至少 5 GB
- 主机端口 `5173`（网页）和 `8000`（API）未被占用
- Windows 自带 `curl.exe`（启动脚本用于健康检查）
- Docker Desktop 已允许访问项目所在目录

平台依赖已封装在 Docker 镜像中，宿主机不需要单独安装 Python、Node.js 或 FastAPI。

Compose 项目名固定为 `documents_workspace`，因此项目目录可以使用中文名称或包含中文字符的路径。

## 启动

推荐双击项目根目录的 `启动平台.bat`。项目可直接复制到另一台 Windows 电脑；脚本使用项目根目录相对路径创建默认数据目录，并在窗口中逐项检查项目文件、`.env`、Docker 客户端/引擎、Compose 服务定义、四个服务、API 和网页。缺少 Docker、Compose、项目文件、写权限或服务时，会显示具体缺项和修复方法。
如若启动失败，检查是否将项目套入标题过长或者层级过多的目录中，也可以先运行一遍 `关闭平台.bat`再运行 `启动平台.bat`
迁移到新电脑时，只需复制项目目录和需要保留的 `runtime` 数据目录，在新电脑安装并启动 Docker Desktop 后双击启动脚本。若没有 `.env`，脚本会提示并使用 Compose 默认值；可先复制 `.env.example` 为 `.env` 再填写本机配置。默认数据库和文档目录为项目下的 `runtime/postgres`、`runtime/storage`，不会依赖原电脑的盘符。

也可以在项目根目录运行：

```bash
docker compose up -d --build
```

启动后访问：

- 网页：<http://localhost:5173>
- API 文档：<http://localhost:8000/docs>
- API 健康检查：<http://localhost:8000/api/health>

关闭服务请双击 `关闭平台.bat`。网页“设置”中的“结束服务”按钮会提示此操作；浏览器不能直接执行本机 BAT 文件。

## 主要功能

- 批量上传、下载、移除和归档文件
- 单个文件和批量文件改名；支持查找替换、前缀、后缀和扩展名修改
- 支持 PDF、DOCX、PPT/PPTX、XLS/XLSX、TXT、Markdown、CSV、JSON、YAML、XML、HTML 及常见代码文件；PPTX、XLSX、XLS 可提取正文建立索引，旧版 PPT 可保存和下载，若无法提取会标记为“暂不支持”
- 文件名搜索、正文关键词搜索、语义检索
- 真实文本向量生成、持久化和余弦相似度检索
- 文件索引状态显示，失败后可重试
- PDFium WebAssembly 渲染 PDF 原文；预览接口使用 JSON 传输 PDF 字节，避免浏览器 PDF 导航被下载工具拦截；同时支持 DOCX、HTML、PPTX 和 XLS/XLSX 原文预览，并可切换纯文本预览
- 列表视图和分类树视图
- 多级分类、分类改名、文件直接归类
- 分类、归档、格式、时间范围和排序筛选
- 点击资料列表表头按名称、分类、大小、索引状态或上传时间正序/逆序排列
- 数据库概览、文件路径和索引状态查看
- 设置中的操作台透明度、背景图片透明度、衬色颜色/透明度和悬浮窗动效选项
- 设置中的语义检索首屏显示数量；其余语义检索结果默认折叠，可展开查看
- 设置中的主标题、副标题文字及颜色

## 自然语言检索示例

在搜索模式中选择“语义检索”，可以使用与原文不同的说法：

语义检索实际执行以下链路：

`文档解析 → Chunk 分段 → Embedding 向量入库 → Query Embedding → 向量 Top-K 召回 + BM25 关键词召回 → RRF 结果融合 → 本地可解释重排 → 来源文件和片段返回`

其中向量和文本片段保存在 PostgreSQL 的 `document_chunks` 表中；BM25 在检索时按候选片段计算，RRF 同时利用向量排名和关键词排名，最后的重排会提高标题命中、完整短语和正文证据的权重。联网 Embedding 不可用时自动回退到本地向量，索引和检索接口保持不变。

| 查询                                  | 预期相关资料                                |
| --- | --- |
| 项目代码提交前需要做哪些检查            | 代码提交规范、发布检查清单                    |
| 同一个网络请求重复发送会不会生成两条记录 | 文件服务接口约定、故障复盘、用户访谈纪要       |
| 旧项目资料如何隐藏但以后还能恢复        | 文档分类与归档规范、用户访谈纪要              |
| 上传成功但正文搜不到应该从哪里排查      | 本地部署故障排查、故障复盘、文件服务接口约定   |
| 怎样区分通用发布规范和某次发布执行结果  | 发布检查清单、项目发布检查清单、用户访谈纪要   |

结果应显示来源文件、相关度和正文片段。

## 数据目录与自定义路径

- `runtime/postgres`：默认 PostgreSQL 数据库目录
- `runtime/storage`：默认上传文件目录
- `backend`：API 和索引 worker
- `frontend`：网页前端

项目可以放在任意目录。Compose 使用相对项目根目录的路径，不依赖固定盘符或目录。若要分别选择数据库和文档的存放位置，在 `.env` 中设置：

```dotenv
POSTGRES_HOST_PATH=E:/data/file-platform/postgres
STORAGE_HOST_PATH=E:/data/file-platform/storage
```

路径可以是相对路径或绝对路径；Windows 绝对路径建议使用正斜杠。修改路径后重新运行 `启动平台.bat` 或 `docker compose up -d`。

网页“设置”中也可以直接修改数据库目录和文档目录。点击“保存设置”并确认后，平台会通过已启动的 `启动平台.bat` 监视器停止服务、移动两个目录、更新 `.env` 并自动重启；目标目录必须为空。迁移期间网页会短暂不可用，完成后会自动恢复。迁移失败时原目录会保留，详情写入 `runtime/.storage-migration.log`。


网页设置中的存储方式：

- `flat`：原文件放在统一目录，分类保存在数据库
- `folders`：按“资料/归档/分类层级”保存，分类变化时同步移动平台内文件

浏览器上传只能保存平台副本，不能删除或移动电脑上的原文件；资料工作台会提示源文件保持不变。

不要执行 `docker compose down -v`，否则会删除 Docker 数据卷。索引失败不会删除原文件，文件仍可下载和重试索引。

## 可选联网模型

默认使用本地向量，不需要联网。可在网页“设置”中选择联网模型并填写 OpenAI-compatible embeddings 接口、模型名称和 API Key。联网模型不可用时会回退到本地向量。

## 依赖与模块

宿主机只需要 Docker Desktop；Python、Node.js、Nginx 和应用依赖均在镜像内安装。

- `backend/requirements.txt`：FastAPI、Uvicorn、python-multipart、psycopg、NumPy、pypdf、python-docx、python-pptx、openpyxl、xlrd
- `backend/app/main.py`：API、上传下载、预览、搜索、设置和数据库概览
- `backend/app/indexer.py`：PDF、DOCX 和文本提取、分块、索引任务
- `backend/app/embeddings.py`：本地文本向量、可选 OpenAI-compatible 向量接口和余弦检索
- `backend/app/worker.py`：后台索引 worker
- `frontend/package.json`：React、Vite、`@embedpdf/pdfium`（PDFium WebAssembly 原文渲染）、SheetJS `xlsx`（Excel 原文解析）、PptxJS 浏览器查看器（通过 `pptxjs` npm 别名使用 `pptx-viewer`）
- `runtime/postgres`：PostgreSQL 持久化数据（可在设置中迁移）
- `runtime/storage`：上传文档持久化目录（可在设置中迁移）
- `runtime/ui-background`：操作台背景图片持久化目录

背景图片支持 PNG、JPG、JPEG、WEBP、GIF，单张不超过 10 MB。设置保存到数据库，保存后立即应用；“保存设置”成功后设置窗口会自动关闭。

## 常见检查

- Docker 未启动：启动 Docker Desktop 后重新运行 `启动平台.bat`
- 端口冲突：释放 `5173` 或 `8000` 端口
- 服务异常：运行 `docker compose ps` 和 `docker compose logs --tail=80 api worker web db`
- 配置检查：运行 `docker compose config`
- 网页空白：重新运行 `docker compose up -d --build web`
