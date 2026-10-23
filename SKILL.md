---
name: ielts-vocab-writer
description: 把用户提供的英语单词、短语、同义/反义词组录入本地 Word 词汇本（.docx），自动生成 IPA 音标、词性、中文释义、搭配、词族、词根和辨析，并按语义分类写入固定章节。当用户要求整理/录入/添加雅思（或其他英语考试）词汇到 Word 文档、词汇本、生词本，或连续发送单词、短语、两个斜杠分隔的词对时，使用本技能。首次使用自动在桌面生成带 8 个章节和固定排版的模板文档，之后持续追加。
---

# IELTS Vocab Writer

把用户发的英语词汇写进本地 Word 词汇本。内容（音标、释义、搭配等）由你生成；排版、分类落位、文件初始化由 `scripts/add.py` 确定性完成。

## 依赖

仅需 `python-docx`：`pip install python-docx`（未安装时先安装）。

## 输入识别规则

| 用户输入 | 模式 | 处理 |
|---|---|---|
| 单个单词（`excess`、`methane`） | word | 生成完整词条，按语义归入章节 |
| 单个短语（`take off`、`be located in`、`make out`） | phrase | 生成词条，写入「短语区」 |
| 两个词/短语用 `/` 分隔（`situated in / be located in`） | synonym | 判断同义或反义，写入同义替换区对应子区域 |
| 一次发 3 个及以上、无 `/` | 批量 word/phrase | 逐个识别、各自归类 |
| 完整句子 | 不录入 | 句子不写入文档；按用户要求做解析，其中值得记的词经用户确认后再录入 |

## 每次必须生成的内容

### word
- `phonetic`：IPA 音标（用 `/ /`）
- `pos_meaning`：词性 + 简明中文释义，多词性都列出（如 `n. 过量，过剩；adj. 过量的`）
- `supplements`：3–5 条，从下列中选最有价值的：
  - 高频搭配（英中对照）
  - 词族（同根词，标注词性）
  - 词根/词缀拆解
  - 近义词/反义词辨析（说明语义差别）
- `section`：语义所属章节的关键词（见下节）

### phrase
- `meaning`：中文释义（多义项用分号分隔）
- `supplements`：高频搭配、近义/反义短语、名词形式等

### synonym（`/` 分隔的两个词/短语）
- `relation`：`synonym`（同义/近义）或 `antonym`（反义），根据词义判断
- `terms`：两个词/短语的数组（用动词原形/常见形态）
- `meaning`：共同含义或反义对比
- `supplements`：两者用法/语体/搭配差异、拓展词

## 章节分类（section 关键词）

模板默认章节，`section` 只需填关键词，脚本做模糊匹配：

| 关键词 | 章节 | 典型词汇 |
|---|---|---|
| 动物 | 动物与生态 | prey, predator, habitat, species |
| 学术 | 学术与写作 | estimate, comparison, predict, reduction, drastically |
| 环境 | 环境与科学 | methane, destruction, decompose, emission |
| 社会 | 社会与职场 | unemployment, corporate, executive, reliance |
| 情感 | 情感与日常 | 情绪、日常生活动词 |
| 通用 | 通用词汇 | 跨话题通用词（vertical, via, amid, burst） |

判断不了时归入「通用」。若用户改过模板章节名，选语义最接近的章节关键词；没有匹配项时脚本会自动落到通用章节。

## 执行流程

### 第 0 步：首次使用必须先确认保存位置（强制）

**禁止在用户不知情的情况下静默生成文档。** 每次开始录入前先运行：

```bash
python "<skill目录>/scripts/add.py" --where
```

- 若输出 `first_run: true`（用户第一次用）：**先停下询问用户**，给出明确选项，例如：
  > "词汇本要保存在哪里？① 桌面（默认，路径为 C:\Users\你的用户名\Desktop\雅思阅读词汇.docx）；② 指定其他位置（告诉我文件夹路径，如 D:\notes）。"
  - 用户选桌面 → 直接进入第 1 步（脚本会在桌面自动初始化）
  - 用户指定目录/文件 → 先运行 `python add.py --set-path "<路径>" --init`，再进入第 1 步
- 若 `first_run: false`：用户已配置过，直接使用，**不要重复询问**，也不要随意改路径。
- 用户中途说"以后写到 D 盘/某文件夹" → 运行 `--set-path "<路径>"`；说"改回桌面" → 运行 `--reset-path`。
- 改路径时若旧文档已有词条，提醒用户：旧文档不会自动移动，建议先手动把旧 .docx 移到新位置，或把新路径直接指向旧文档。

### 第 1 步：写入

准备词条 JSON（单个对象或数组），字段见 `scripts/add.py` 顶部注释，然后调用：

```bash
python "<skill目录>/scripts/add.py" --input entries.json
```

其他参数：`--json '<json>'` 直接传 JSON；`--init` 仅初始化；`--document <路径>` 仅本次临时指定文档（不修改用户的永久默认路径）。

### 第 2 步：报告结果

1. 检查脚本输出，向用户简报每条落位的章节，并把脚本输出的**文档完整路径**（`[done]` 那一行）告诉用户。
2. 若输出 `[warn] 原文件被占用...已另存为`，告知用户新文件名和完整路径（文档在 Word 中打开时会触发）。

## JSON 格式

```json
[
  {
    "type": "word",
    "section": "环境",
    "word": "methane",
    "phonetic": "/ˈmiːθeɪn/",
    "pos_meaning": "n. 甲烷，沼气",
    "supplements": ["methane emissions 甲烷排放", "greenhouse gas 温室气体"]
  },
  {
    "type": "phrase",
    "phrase": "take off",
    "meaning": "（飞机）起飞；脱下；腾飞",
    "supplements": ["反义：land 降落；put on 穿上"]
  },
  {
    "type": "synonym",
    "relation": "synonym",
    "terms": ["be situated in", "be located in"],
    "meaning": "位于……",
    "supplements": ["situated 更正式，located 更常用"]
  },
  {
    "type": "synonym",
    "relation": "antonym",
    "terms": ["burst", "reduction"],
    "meaning": "突发增加 / 减少",
    "supplements": ["描述数据变化时趋势相反"]
  }
]
```

## 约束

- 全部用中文写释义，音标用 IPA。
- 不要给例句（用户明确要求）；搭配≠例句，保留。
- 不修改模板和脚本中的格式常量（缩进、字号、颜色已与模板排版绑定）。
- 不新建文档章节；用户需要新分类时，让其自行在 Word 中添加标题，脚本会动态识别。
