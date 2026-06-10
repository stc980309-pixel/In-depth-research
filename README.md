我的skill是整合academic-paper、ppt-generator、planning-with-files等有利于文献调研的高星skill，所以整体的可靠性和实用性还是不错的。然后我的skill主要针对的是需要 深度调研某个领域，因为现在的academic-paper这个skill（已整合到我的skill中）是可以实现对几乎所有文献的摘要提取阅读，所以对于一些基础性的调研已经很成熟了，但是因为没有仔细阅读文献全文导致对一些细节以及工艺调研不深入。我的skill是可以通过校园网绕过期刊的人机交互，从而实现自动化下载高质量的文献进行阅读。最后生成高质量的报告文档和PPT。
目前由于是通过校园网绕过人机交互，所以codex这种需要翻墙的agent无法使用。其次我在skill中是写入了地大和华科的校园网，具体华科网能否使用，目前还没有测试。
由于我的目的是深度调研，所以我在skill中选择的文献是顶级期刊，但是是我人为定义的顶级期刊（Nuture、子刊、Angew、JACS、AM等），如果是其他同学、老师使用，得和Claude code交流添加其他领域的顶级期刊作为筛选的白名单。

# 深度调研 Claude Code Skill 套件

> 文献深度调研自动化流水线：一句话触发，从理解问题到输出 Word 报告 + PPT，全流程自动完成。

## 这是什么？

一套 Claude Code 的 **5 个联动 Skill**，实现文献深度调研的标准化 8 步流程。输入 "深度调研XX"，自动完成：

```
理解问题 → 诊断根因 → AI搜索文献 → 筛选+下载 → 全文精读 → 对比分析 → 生成Word报告 → 生成PPT
```

## 适用场景

- 进入新研究领域，需要快速了解文献全景
- 实验中遇到问题，需要文献佐证和解决方案
- 写论文/开题报告，需要系统性地整理文献
- 导师要求"做个文献调研PPT汇报"

## 环境要求

| 项目 | 要求 | 说明 |
|------|------|------|
| 操作系统 | **Windows 10/11** | 当前版本仅支持 Windows |
| Python | 3.8-3.13 | 首次运行自动检测，找不到则 winget 自动安装 |
| 浏览器 | Microsoft Edge | 自动检测，找不到则 winget 自动安装 |
| MS Office | PowerPoint（可选） | 有则用 COM 生成高质量 PPT，无则降级为 python-pptx |
| 网络 | 需要校园网 IP | 用于通过机构订阅下载 paywall 论文 |

**首次运行会自动安装所有缺失依赖，全程无需手动操作。**

## 安装方法

### 1. 复制 Skill 到 Claude Code 配置目录

```bash
# 将所有文件夹复制到 Claude Code 的 skills 目录
cp -r skills/* ~/.claude/skills/
```

详细步骤：

```
1. 打开文件管理器，在地址栏输入：
   %USERPROFILE%\.claude\skills\

2. 将本仓库中以下 5 个文件夹复制进去：
   - literature-survey/
   - planning-with-files/
   - deep-research/
   - academic-paper/
   - ppt-generator/

3. 将 pdf_to_txt.py 复制到你的常用工作目录（或任一路径）

4. 重启 Claude Code（或新开一个终端）
```

### 2. 验证安装

在 Claude Code 中输入：
```
/skills
```
应该能看到 `literature-survey` 在列表中。

### 3. 首次使用

直接输入触发词即可，首次运行会自动安装 Python 依赖：

```
深度调研 有机闪烁体材料
```

首次运行会自动：
1. 检测 Python → 找不到则 winget 安装
2. `pip install PyMuPDF python-docx python-pptx websocket-client`
3. 检测 Edge 浏览器 → 找不到则 winget 安装
4. 启动 Edge CDP 调试模式（用于下载 paywall 论文）

## 使用方法

### 触发词

| 你说 | 触发的 Skill |
|------|-------------|
| **"深度调研XX"** | literature-survey（主流程，含全部 8 步） |
| "深度文献调研XX" | 同上 |
| "调研XX"（普通调研） | 不激活 8 步流程，可用 deep-research 单独搜索 |

### 8 步流程详解

| 步骤 | 做什么 | 输出 |
|------|--------|------|
| **Step 0** | 理解你的研究体系：材料、工艺、参数、卡点 | findings.md |
| **Step 1** | 诊断根因：你的参数 vs 文献成功案例的参数 | 诊断结论 |
| **Step 2** | AI 多角度搜索文献（仅限白名单期刊） | 候选论文列表 |
| **Step 3** | 二次筛选：去重、确认相关性 | 精选论文列表 |
| **Step 4** | 全自动下载：OA论文 curl 直下，paywall论文 CDP 浏览器下载 | PDF 文件 |
| **Step 5** | PDF→TXT + 逐篇精读，提取关键工艺参数 | 精读笔记 |
| **Step 6** | 对比分析：你的参数 vs 文献参数，提出实验方案 | 对比表 + 方案 |
| **Step 7** | 生成 Word 报告（.docx） | 文献调研报告 |
| **Step 8** | 生成 PPT（.pptx），用于汇报 | 调研 PPT |

### 白名单期刊（只下载这些）

Nature, Science, Nature Communications, Nature Photonics, Nature Materials, Nature Chemistry, Nature Physics, Nature Nanotechnology, Science Advances, Angewandte Chemie, JACS, Advanced Materials (AM), Advanced Energy Materials (AEM), Advanced Functional Materials (AFM)

### 支持的出版社下载策略

| 出版社 | 方法 | 说明 |
|--------|------|------|
| Nature/Science OA | curl 直下 | OA 论文无需浏览器 |
| Wiley (Angew/AM/AFM) | CDP 浏览器 | 通过校园网 IP 认证 |
| ACS (JACS) | CDP 浏览器 | 通过校园网 IP 认证 |
| RSC | PMC OA 优先 | 先查 PubMed Central |
| arXiv | curl | 预印本直接下载 |

## 文件结构

```
深度调研 Claude code版skill/
├── README.md                          # 本文件
├── pdf_to_txt.py                      # PDF→TXT 转换工具（PyMuPDF）
│
├── literature-survey/                 # ★ 主 Skill：8步调研流程
│   └── skill.md                       # 核心定义文件（v4.0.0）
│
├── planning-with-files/               # 流程管控 Skill
│   ├── SKILL.md                       # 任务计划/进度追踪
│   ├── examples.md
│   ├── reference.md
│   ├── scripts/                       # Shell 脚本
│   └── templates/                     # 计划模板
│
├── deep-research/                     # AI 文献搜索 Skill
│   ├── SKILL.md                       # 核心定义（v2.9.4）
│   ├── agents/                        # 13 个 AI Agent
│   │   ├── research_architect_agent.md
│   │   ├── synthesis_agent.md
│   │   ├── source_verification_agent.md
│   │   └── ...（共 13 个）
│   ├── examples/                      # 7 个使用示例
│   ├── references/                    # 22 个参考资料
│   └── templates/                     # 6 个模板
│
├── academic-paper/                    # Word 论文写作 Skill
│   ├── SKILL.md                       # 核心定义（v3.1.2）
│   ├── agents/                        # 12 个写作 Agent
│   ├── examples/                      # 8 个示例
│   ├── references/                    # 25 个写作参考
│   └── templates/                     # 12 个论文模板
│
└── ppt-generator/                     # PPT 生成 Skill
    ├── SKILL.md                       # 核心定义（v1.0.0）
    └── references/                    # MS Office COM API 参考
```

## 常见问题

### Q: 下载论文时浏览器打开了但没反应？

检查 Edge 是否以调试模式启动。正常情况下 skill 会自动处理。如果手动排查：

```bash
# 检查 CDP 浏览器是否在线
curl -s http://localhost:9223/json/version

# 如果没有响应，手动启动
taskkill /F /IM msedge.exe
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" ^
  --remote-debugging-port=9223 ^
  --remote-allow-origins="*" ^
  --user-data-dir="%USERPROFILE%\.lit_survey_browser" ^
  --no-first-run
```

### Q: 学校不是中国地质大学/华中科技大学？

目前 skill 的机构订阅检测只认这两所学校。如果你的学校有其他出版社的订阅，可以修改 `literature-survey/skill.md` 中 Step 4f 的学校名称列表。或者你的学校使用 IP 认证（无需登录），下载也会正常工作。

### Q: 非 Windows 系统能用吗？

当前 v4.0.0 只支持 Windows。macOS/Linux 的浏览器自动化部分需要适配（欢迎 PR）。

### Q: Word 报告和 PPT 效果如何？

- Word 报告（.docx）：包含问题诊断、文献参数对比表、实验方案（首选+备选）、失败模式与对策、参考文献列表。中文排版。
- PPT（.pptx）：10 页标准结构，学术蓝底配色，橙色高亮差距，16:9 格式。面向汇报场景。

### Q: 下载论文总是失败怎么办？

Skill 内置 3 轮修复机制：
1. Round 1：直接下载
2. Round 2：诊断问题 → 针对性修复（重启浏览器/预热cookies/换URL）
3. Round 3：降级方案（OA版本 → arXiv → ResearchGate）

3 轮都失败才会跳过该论文，并记录详细原因到报告中。

## 版本

| Skill | 版本 | 更新日期 |
|-------|------|----------|
| literature-survey | v4.0.0 | 2026-06-10 |
| planning-with-files | - | - |
| deep-research | v2.9.4 | 2026-05-18 |
| academic-paper | v3.1.2 | 2026-05-18 |
| ppt-generator | v1.0.0 | 2026-05-30 |

## 作者与许可

- 作者：[史庭昌]
- 许可：MIT License
- 反馈：[769115586@qq.com]

## 致谢

本套件基于 Claude Code Skills 机制构建，调用了以下能力：
- Claude Code CLI (Anthropic)
- Edge DevTools Protocol (CDP)
- PyMuPDF (PDF 解析)
- python-docx / python-pptx (文档生成)
- MS Office COM (PowerPoint 高级渲染，可选)
