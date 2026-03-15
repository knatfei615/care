# 临床药学监护系统

基于 Flask 的临床药学监护工作平台。支持 Web 界面与 CLI 两种使用方式，可从病例清单 Excel 导入患者数据，通过 LLM 或预设模板生成标准化药学监护记录，并写入 `.xlsm` 工作簿（保留 VBA 宏）。

## 功能特性

**Web 端**
- 上传 / 下载 `.xlsm` 监护工作簿
- 查看患者列表，按床号、住院号、姓名检索
- 借助 LLM 将查房口述自动整理为"问题→分析→处理→结果/计划"标准格式
- 一键保存监护记录到 Excel 指定槽位
- 用户注册 / 登录、管理员后台（用户管理、重置密码）

**CLI 端**
- 从病例清单 Excel 批量导入患者基本信息
- 8 套预设模板快速生成监护记录
- 直接将记录写入工作簿，支持自动槽位检测

## 技术栈

| 类别 | 技术 |
|------|------|
| Web 框架 | Flask |
| 认证 | Flask-Login |
| 数据库 | SQLite + Flask-SQLAlchemy |
| Excel 读写 | openpyxl（`keep_vba=True`） |
| LLM | OpenAI 兼容 API（默认 OpenRouter） |
| 生产部署 | Gunicorn |

## 项目结构

```
├── app.py                      # Flask 主应用，路由定义
├── config.py                   # 环境变量与配置
├── models.py                   # 数据模型（User）
├── auth.py                     # 认证蓝图（登录/注册/登出）
├── admin.py                    # 管理员蓝图（用户管理）
├── excel_io.py                 # Excel I/O 线程安全封装
├── llm.py                      # LLM 集成（结构化查房记录）
├── import_case_list.py         # CLI：导入病例清单
├── generate_picu_note.py       # CLI：模板式记录生成
├── write_picu_note_to_excel.py # CLI：写入记录到 Excel
├── picu_note_templates.json    # 8 套监护记录模板
├── templates/                  # HTML 模板
│   ├── index.html              # 主界面（SPA）
│   ├── login.html              # 登录页
│   ├── register.html           # 注册页
│   └── admin.html              # 管理后台
├── requirements.txt            # Python 依赖
└── render.yaml                 # Render 部署配置
```

### 架构概览

```
病例清单 Excel (.xlsx)
       │
       ▼
 import_case_list.py ──► 监护工作簿 (.xlsm, 列 A-K)
                                   │
                   ┌───────────────┴───────────────┐
                   ▼                               ▼
             Flask Web App                    CLI Scripts
           (app.py + index.html)      (generate_picu_note.py,
                   │                   write_picu_note_to_excel.py)
                   ▼                               │
              excel_io.py ◄────────────────────────┘
           (线程安全封装)
                   │
                   ▼
             openpyxl (keep_vba=True)
```

## 快速开始

```bash
# 克隆仓库
git clone <repo-url>
cd care

# 创建虚拟环境
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 配置环境变量（复制并编辑）
cp .env.example .env

# 启动开发服务器
python app.py
# 访问 http://localhost:5000
```

首次注册的用户将自动成为管理员。

## 环境变量

在项目根目录创建 `.env` 文件：

```env
OPENAI_API_KEY=sk-your-api-key
OPENAI_BASE_URL=https://openrouter.ai/api/v1
OPENAI_MODEL=anthropic/claude-sonnet-4-6
DATA_DIR=./data
PORT=5000
SECRET_KEY=your-secret-key
ADMIN_USERNAME=admin
ADMIN_PASSWORD=your-admin-password
```

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `OPENAI_API_KEY` | OpenAI 兼容 API 密钥 | （必填） |
| `OPENAI_BASE_URL` | API 基础地址 | `https://openrouter.ai/api/v1` |
| `OPENAI_MODEL` | 使用的模型名称 | `anthropic/claude-sonnet-4-6` |
| `DATA_DIR` | 数据存储目录 | `./data` |
| `PORT` | 服务端口 | `5000` |
| `SECRET_KEY` | Flask 会话密钥 | `dev-change-me-in-production` |
| `ADMIN_USERNAME` | 预设管理员用户名 | （可选） |
| `ADMIN_PASSWORD` | 预设管理员密码 | （可选） |

## CLI 使用

### 导入病例清单

```bash
python import_case_list.py --pharmacist "药师姓名" --employee-id "工号"
```

自动查找当前目录下最新的 `*病例清单*.xlsx` 文件，去重后写入监护工作簿。

### 生成监护记录

```bash
python generate_picu_note.py                          # 交互式生成
python generate_picu_note.py --list                    # 列出可用模板
python generate_picu_note.py --template abx_review --copy  # 生成并复制到剪贴板
```

### 写入记录到 Excel

```bash
python write_picu_note_to_excel.py --list-patients
python write_picu_note_to_excel.py --inpatient-no "12345" --template tdm --set "药物=万古霉素"
```

## Excel 工作簿格式

监护工作簿（`.xlsm`）中，第 1-2 行为表头，第 3-962 行为数据行。

| 列 | 内容 |
|----|------|
| A-K | 患者基本信息（科室、药师、工号、住院号、床号、姓名、年龄、性别、体重、入院日期、入院诊断） |
| L-O | 记录槽位 1（日期、分级、类型、记录） |
| P-S | 记录槽位 2 |
| T-W | 记录槽位 3 |
| X-AA | 记录槽位 4 |
| AB-AE | 记录槽位 5 |
| AF-AI | 记录槽位 6 |

每位患者最多 6 条监护记录。监护分级包括：一级监护、二级监护、三级监护。记录类型包括：药学查房、药物重整、药学监护、用药咨询、用药教育。

## 监护记录模板

系统内置 8 套标准化模板，均遵循"问题→分析→处理→结果/计划"四段式格式：

| 模板 ID | 名称 |
|---------|------|
| `abx_review` | 抗菌药物评估 |
| `renal_crrt` | 肾功能/CRRT 剂量评估 |
| `tdm` | TDM/血药浓度监护 |
| `adr` | 不良反应监护 |
| `interaction` | 相互作用评估 |
| `sedation_analgesia` | 镇静镇痛监护 |
| `nutrition` | 营养支持评估 |
| `infusion_compatibility` | 输注/配伍评估 |

## 部署

项目提供 `render.yaml` 配置，支持一键部署至 [Render](https://render.com)：

```bash
# 生产环境启动
gunicorn app:app --bind 0.0.0.0:$PORT
```

需在 Render 控制台配置环境变量 `OPENAI_API_KEY`、`ADMIN_USERNAME`、`ADMIN_PASSWORD`。工作簿数据存储在持久化磁盘（`/data`，1 GB）。

## 许可证

待定。
