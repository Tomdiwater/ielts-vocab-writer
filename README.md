# ielts-vocab-writer

一个 AI 技能（Skill）：把你背单词时随手发的**单词、短语、同义/反义词组**，自动整理成一份排版统一的 Word 词汇本。首次使用时在桌面生成带分类章节的模板，之后持续追加。

## 效果

在对话框里直接发：

```
methane
take off
situated in / be located in
burst / reduction
```

AI 自动生成音标、释义、搭配、词族、词根、辨析，并写入桌面《雅思阅读词汇.docx》：

- **单词** 按语义归入对应章节（动物与生态 / 学术与写作 / 环境与科学 / 社会与职场 / 情感与日常 / 通用词汇）
- **短语** 进入短语区
- **两个词用 `/` 分隔** → 同义替换，AI 判断是同义还是反义，进入同义替换区对应子区域

## 录入规则

| 输入 | 结果 |
|---|---|
| `excess` | 单个单词 → 自动分类到对应章节 |
| `take off` | 短语 → 短语区 |
| `situated in / be located in` | 同义/近义 → 同义替换区·同义词组 |
| `burst / reduction` | 反义 → 同义替换区·反义词组 |
| 一次发 3 个以上 | 批量录入，各归各的分类 |
| 完整句子 | 不录入（只做解析） |

## 排版格式（模板内置）

- 词头加粗黑色，后跟 IPA 音标、词性、中文释义
- 补充行（▸ 开头）：9pt 深灰色、左缩进，内容含搭配、词族、词根、辨析
- 章节标题：13pt 加粗，8 个分类章节

## 安装

本技能遵循通用的 **SKILL.md 技能格式**，任何支持该格式的 AI Agent 均可使用。

### 前置条件：Python + python-docx

1. 到 https://www.python.org/downloads/ 下载安装 Python（Windows 安装时**务必勾选 `Add Python to PATH`**）
2. 打开终端（Windows 按 `Win + R` 输入 `powershell`），执行：

```bash
pip install python-docx
```

### 方式一：豆包桌面端「上传技能」（最简单，推荐）

1. 在本仓库页面点绿色 **`Code` → `Download ZIP`**，下载并解压
2. 打开豆包桌面端 → 技能管理 → 点右上角 **`+ 添加` → `上传技能`**
3. 选择下载的 zip 包（或解压后的 `ielts-vocab-writer` 文件夹），等待解析
4. 确认 Ielts Vocab Writer 开关已打开，直接发一个单词（如 `methane`）测试

### 方式二：手动放入 skills 目录（其他 AI / 进阶）

下载或克隆本仓库：

```bash
git clone https://github.com/Tomdiwater/ielts-vocab-writer.git
```

把 `ielts-vocab-writer` 整个文件夹放进你所用 AI 的 skills 目录：

| 平台 | 目录 |
|---|---|
| 豆包（桌面端） | `%LOCALAPPDATA%\Doubao\User Data\Default\.doubao\agent_mode\workspace\.user_skills\`（Win+R 粘贴直达） |
| Claude Code | `~/.claude/skills/`（或项目 `.claude/skills/`） |
| 其他支持 SKILL.md 的 Agent | 其文档指定的 skills 目录 |

放入后**完全重启 AI 会话**，直接发单词即可。首次使用会在桌面自动生成《雅思阅读词汇.docx》。

### 安装后没反应？

按顺序排查：

1. AI 是否完全重启过（豆包需在右下角托盘图标右键退出）
2. 技能列表里 Ielts Vocab Writer 开关是否打开
3. 文件夹是否套了两层：路径必须以 `ielts-vocab-writer/SKILL.md` 结尾，不能再多一层
4. `python-docx` 是否安装成功（重新执行一次 `pip install python-docx`）

## 生成的文档在哪里？

**默认在你的「桌面」上**，文件名叫 `雅思阅读词汇.docx`：

| 系统 | 完整路径 |
|---|---|
| Windows | `C:\Users\<你的用户名>\Desktop\雅思阅读词汇.docx` |
| macOS | `/Users/<你的用户名>/Desktop/雅思阅读词汇.docx` |

说明：

- Windows 下脚本读取注册表定位真实桌面——即使你把桌面改到了 D 盘，或开启了 OneDrive 桌面备份，文件也会出现在你实际看到的那个桌面上，不用手动找盘。
- 每次写入后，脚本都会在终端输出一行 `[done] 已写入 N 条，文档: <完整路径>`，AI 会把这个路径告诉你。
- 想放到其他位置（如 `D:\notes\`）：命令行加 `--document "完整路径"`，或直接让 AI 写入你指定的文档。
- 文档正在 Word 中打开时写入，脚本会自动另存为 `雅思阅读词汇_v2.docx`、`_v3.docx`，路径仍在桌面，不会丢内容。

## 命令行手动使用（不依赖 AI）

也可以直接运行脚本，把词条写成 JSON：

```bash
# 仅初始化文档（在桌面生成模板）
python scripts/add.py --init

# 从 JSON 文件写入
python scripts/add.py --input entries.json

# 指定文档路径
python scripts/add.py --document "D:\notes\我的词汇.docx" --input entries.json
```

JSON 格式见 [`examples/entries.example.json`](examples/entries.example.json)，支持三种类型：

```json
{ "type": "word",   "section": "环境", "word": "methane",
  "phonetic": "/ˈmiːθeɪn/", "pos_meaning": "n. 甲烷，沼气",
  "supplements": ["methane emissions 甲烷排放"] }

{ "type": "phrase", "phrase": "take off", "meaning": "起飞；脱下",
  "supplements": ["反义：land 降落"] }

{ "type": "synonym", "relation": "synonym",
  "terms": ["be situated in", "be located in"], "meaning": "位于……",
  "supplements": ["situated 更正式"] }
```

`relation` 取 `synonym`（同义）或 `antonym`（反义）。

## 自定义章节

章节不写死在脚本里，而是**动态读取文档中的标题**。想增加分类（如"心理与认知"），直接在 Word 里添加一个同级标题（如 `九、心理与认知`），AI 会自动把相关词汇写入该章节。

## 工作原理

- `SKILL.md`：给 AI 的工作流指令（识别输入、生成内容、调用脚本）
- `assets/template.docx`：带格式的空模板（首次运行复制到桌面）
- `scripts/common.py`：章节定位、格式排版、文件占用自动另存
- `scripts/add.py`：统一写入入口（word / phrase / synonym）

文件被 Word 打开占用时，脚本会自动另存为 `雅思阅读词汇_v2.docx`、`_v3.docx`，不会丢失内容。

## 依赖

- Python 3.8+
- [python-docx](https://pypi.org/project/python-docx/)
