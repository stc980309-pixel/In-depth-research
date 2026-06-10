---
name: literature-survey
description: 文献深度调研标准化8步流程：理解→诊断→deep-research搜索→下载→精读→对比→academic-paper写Word→ppt-generator做PPT。用户说"深度调研XX"时激活。白名单期刊下载全部不询问。跨平台零配置。
metadata:
  version: "4.0.0"
  last_updated: "2026-06-10"
  status: active
  platform: any
compatibility: "planning-with-files + deep-research + CDP浏览器(Python websocket-client) + academic-paper + ppt-generator + Python 3.8+ | 全部缺项自动安装"
triggers:
  - "深度调研*"
  - "深度文献调研*"
allowed-tools:
  - Bash
  - Write
  - Read
  - Edit
  - WebSearch
  - WebFetch
  - Skill
  - Agent
  - TaskCreate
  - TaskUpdate
---

# 文献深度调研标准化流程（8步法）

## 前置：一键环境安装（首次执行时自动运行）

**核心原则：缺什么装什么，全部自动，不询问用户。不预设任何软件已安装。**

执行 8 步之前，按顺序完成以下检测和安装。每步成功后报告进度。

### 1. Python 检测与自动安装 (Windows)

**Step 1a: 按优先级搜索 Python**

```bash
PYTHON=""

# 1. 先试命令名（兼容已加入 PATH 的 Python）
for cmd in python3 python py python3.10 python3.11 python3.12 python3.13; do
    if command -v "$cmd" >/dev/null 2>&1; then
        ver=$("$cmd" --version 2>/dev/null || true)
        if echo "$ver" | grep -q "Python 3"; then
            PYTHON=$(command -v "$cmd" 2>/dev/null || echo "$cmd")
            echo "FOUND via PATH: $PYTHON ($ver)"
            break
        fi
    fi
done

# 2. 没找到 → 扫 Windows 常见安装路径
if [ -z "$PYTHON" ]; then
    # Python 3.10-3.13, 两种安装来源: python.org 和 Microsoft Store
    for ver in 313 312 311 310 39 38; do
        for base in \
            "/c/Users/$USER/AppData/Local/Programs/Python/Python${ver}/python.exe" \
            "/c/Program Files/Python${ver}/python.exe" \
            "/c/Program Files (x86)/Python${ver}/python.exe" \
            "/c/Python${ver}/python.exe" \
            "D:/py/Python${ver}/python.exe"; do
            if [ -f "$base" ]; then
                PYTHON="$base"
                echo "FOUND: $PYTHON"
                break 2
            fi
        done
    done
fi

# 3. 再试 Microsoft Store Python（路径含 WindowsApps）
if [ -z "$PYTHON" ]; then
    for p in "$LOCALAPPDATA/Microsoft/WindowsApps/python3.exe" \
             "$LOCALAPPDATA/Microsoft/WindowsApps/python.exe" \
             "/c/Users/$USER/AppData/Local/Microsoft/WindowsApps/python3.exe" \
             "/c/Users/$USER/AppData/Local/Microsoft/WindowsApps/python.exe"; do
        p=$(echo "$p" | sed 's|\\|/|g')  # normalize backslashes
        if [ -f "$p" ]; then
            PYTHON="$p"
            echo "FOUND Microsoft Store: $PYTHON"
            break
        fi
    done
fi
```

**Step 1b: 没找到 → winget 自动安装**

```bash
if [ -z "$PYTHON" ]; then
    echo "Python not found. Auto-installing via winget..."
    winget install Python.Python.3.12 --accept-package-agreements --accept-source-agreements 2>/dev/null
    # winget 安装后 Python 在这个位置
    NEW_PY="$USERPROFILE/AppData/Local/Programs/Python/Python312/python.exe"
    # 也尝试旧版 winget 的安装路径
    [ ! -f "$NEW_PY" ] && NEW_PY="$LOCALAPPDATA/Programs/Python/Python312/python.exe"
    if [ -f "$NEW_PY" ]; then
        PYTHON="$NEW_PY"
        echo "INSTALLED: $PYTHON"
    else
        echo "FATAL: winget install failed. Please install Python 3.10+ manually."
        exit 1
    fi
fi

echo "PYTHON=$PYTHON"
$PYTHON --version
```

### 2. Python 依赖自动安装

```bash
$PYTHON -m pip install --upgrade pip --quiet 2>/dev/null
$PYTHON -m pip install PyMuPDF python-docx python-pptx websocket-client --quiet 2>/dev/null
echo "Python dependencies ready: PyMuPDF, python-docx, python-pptx, websocket-client"
```

### 3. CDP 浏览器检测与启动

**核心方案：Python `websocket-client` 直连 CDP 浏览器，不依赖 playwright MCP。**

MCP 在新会话可能需要重启才能加载，因此直接通过 CDP WebSocket 控制浏览器更可靠。Python 依赖 `websocket-client`（已在步骤 2 安装）。

**Step 3a: 检测 CDP 浏览器是否在线**

```bash
curl -s http://localhost:9223/json/version >/dev/null 2>&1 && echo "CDP ONLINE" || echo "CDP OFFLINE"
```

如果在线 → 跳到步骤 4。

**Step 3b: 不在线 → 启动 Edge CDP 浏览器**

自动找到 Edge（Windows 全覆盖）：
```bash
BROWSER=""
# 搜索所有可能的 Edge 安装位置（按优先级）
for p in \
    "/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe" \
    "/c/Program Files/Microsoft/Edge/Application/msedge.exe" \
    "$PROGRAMFILES/Microsoft/Edge/Application/msedge.exe" \
    "$PROGRAMFILES(x86)/Microsoft/Edge/Application/msedge.exe" \
    "$LOCALAPPDATA/Microsoft/Edge/Application/msedge.exe"; do
    # 标准化路径分隔符
    p=$(echo "$p" | sed 's|\\|/|g')
    if [ -f "$p" ]; then
        BROWSER="$p"
        echo "FOUND Edge: $BROWSER"
        break
    fi
done

# 如果没找到 Edge，找 Chrome（部分 Windows 用户只用 Chrome）
if [ -z "$BROWSER" ]; then
    for p in \
        "/c/Program Files/Google/Chrome/Application/chrome.exe" \
        "/c/Program Files (x86)/Google/Chrome/Application/chrome.exe"; do
        if [ -f "$p" ]; then
            BROWSER="$p"
            echo "FOUND Chrome (fallback): $BROWSER"
            break
        fi
    done
fi

# 两个都没有 → 装 Edge
if [ -z "$BROWSER" ]; then
    echo "No browser found. Installing Edge via winget..."
    winget install Microsoft.Edge --accept-package-agreements --accept-source-agreements 2>/dev/null
    for p in "/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe" \
             "/c/Program Files/Microsoft/Edge/Application/msedge.exe"; do
        if [ -f "$p" ]; then BROWSER="$p"; echo "INSTALLED Edge: $BROWSER"; break; fi
    done
fi

[ -z "$BROWSER" ] && { echo "FATAL: Cannot find or install a browser"; exit 1; }
```

启动浏览器（**关键参数，Windows**）：
```bash
BROWSER_USER_DATA="$HOME/.lit_survey_browser"
mkdir -p "$BROWSER_USER_DATA"

# 两个参数都是必须的：
#   --remote-allow-origins=*    不加 → Python websocket 403 Forbidden
#   --user-data-dir             固定目录保留校园网 cookies/profile
#   --no-first-run              跳过首次运行向导
"$BROWSER" --remote-debugging-port=9223 \
           --remote-allow-origins="*" \
           --user-data-dir="$BROWSER_USER_DATA" \
           --no-first-run --no-default-browser-check &
sleep 4
curl -s http://localhost:9223/json/version >/dev/null 2>&1 && echo "CDP ONLINE" || echo "CDP OFFLINE"
```

**⚠️ 不要随意重启浏览器**：重启 = 丢失 session/cookies。仅在 CDP OFFLINE 时才启动。

**如果已有浏览器运行但 CDP 不在线**（用户自己开的 Edge 没加 `--remote-debugging-port`）：
- 先告知用户"需要重启浏览器开启调试端口"
- 用 `taskkill //F //IM msedge.exe` 杀掉已有进程，再启动新的



### 4. Office / PPT 生成能力检测

- **Windows**：检查 MS Office PowerPoint COM 是否可用 → 优先用 `ppt-generator` skill
- **所有平台备选**：`python-pptx`（已在步骤 2 安装），纯 Python 生成无外部依赖

```bash
# Windows: 检测 PowerPoint COM
$PYTHON -c "import win32com.client; win32com.client.Dispatch('PowerPoint.Application')" 2>/dev/null && echo "PPT_COM_READY" || echo "PPT_COM_NOT_FOUND"
```

如果 COM 不可用 → PPT 自动走 python-pptx 方案。

### 5. pdf_to_txt.py 自动创建

检查当前目录是否有 `pdf_to_txt.py`，如果没有则自动创建：

```python
#!/usr/bin/env python3
"""Convert all PDFs in a directory to plain text files using PyMuPDF."""
import sys, os, fitz

def pdf_to_txt(pdf_dir, merge=False):
    pdf_dir = os.path.abspath(pdf_dir)
    pdfs = sorted([f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')])
    if not pdfs:
        print(f"No PDFs found in {pdf_dir}")
        return

    all_text = []
    for pdf_file in pdfs:
        pdf_path = os.path.join(pdf_dir, pdf_file)
        txt_path = os.path.splitext(pdf_path)[0] + '.txt'
        try:
            doc = fitz.open(pdf_path)
            lines = [f"=== {pdf_file} ===\n"]
            for i, page in enumerate(doc):
                text = page.get_text()
                if text.strip():
                    lines.append(f"--- Page {i+1} ---\n{text}\n")
            doc.close()
            full_text = ''.join(lines)
            if not merge:
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write(full_text)
                print(f"  -> {os.path.basename(txt_path)} ({len(full_text)} chars)")
            else:
                all_text.append(full_text)
        except Exception as e:
            print(f"  ERROR {pdf_file}: {e}")

    if merge and all_text:
        merged_path = os.path.join(pdf_dir, 'all_papers.txt')
        with open(merged_path, 'w', encoding='utf-8') as f:
            f.write('\n\n'.join(all_text))
        print(f"  -> {merged_path}")

if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('pdf_dir', help='Directory containing PDF files')
    p.add_argument('--merge', action='store_true', help='Merge all into single all_papers.txt')
    args = p.parse_args()
    pdf_to_txt(args.pdf_dir, args.merge)
```

### 6. 校园网/机构订阅检测

**只认两所学校，不猜其他的。**

MCP 浏览器运行在用户本机，网络出口 = 用户本机 IP。下载论文时，访问出版社主页或文章页后，检查页面是否出现以下**任一**学校名称：

- **China University of Geosciences**（中国地质大学（武汉））
- **Huazhong University of Science and Technology**（华中科技大学）

**出现以上任一 → 该出版社可下载。两个都没出现 → 该出版社无订阅。**

**核心原则**：不猜第三所学校，不模糊匹配。两所学校的英文名称必须精确出现。两个都没有就是没权限，直接走降级方案，不浪费时间试。

### 环境安装完成后报告

所有检测和安装完成后，输出一行汇总：
```
[环境就绪] Python=$PYTHON | CDP浏览器=OK | PDF工具=OK | PPT=COM/python-pptx | websocket-client=OK
```
如果某项失败，标为 FAIL 并说明原因。**FAIL 的项不阻塞流程，能走多远走多远。**

---

## "深度调研"的定义

**深度调研 ≠ 搜论文下载。深度调研 = 理解问题 → 诊断根因 → 文献佐证 → 提出方案 → 输出报告 + PPT。**

用户说"深度调研XX"，不是要一堆 PDF 列表，而是要你：
1. 先搞懂他们在做什么、卡在哪
2. 找到卡住的根本原因
3. 用文献数据佐证你的诊断
4. 给出可执行的解决方案
5. **输出一份 Word 报告 + 一份调研 PPT，PPT 用于向团队/导师汇报调研结果**

## 铁律（不可违反）

- **期刊范围**：仅限 Nature、Science、Nature Communications、Nature Photonics、Nature Materials、Nature Chemistry、Nature Physics、Nature Nanotechnology、Science Advances、Angew、JACS、AM、AEM、AFM（AFM 是最低级别）
- **下载策略**：所有相关论文**全部直接下载，不询问用户选哪些**
- **先理解再搜索**：先读用户论文/数据，诊断问题，再搜文献
- **PPT 必须出**：无论用户是否主动要求 PPT，文献调研的最终交付物必须包含 PPT
- **搜索→筛选→下载→汇总**：四步铁律，缺一步都不叫调研

## 流程管控：planning-with-files

在开始 8 步之前，调用 `planning-with-files` skill 创建三份追踪文件：

- **task_plan.md**：8 步任务清单，每步标记状态（pending → in_progress → done）
- **findings.md**：各步骤的关键发现（Step 0-1 诊断、Step 5 精读笔记、Step 6 对比结论）
- **progress.md**：整体进度 + 当前卡点

每完成一个 Step，更新对应文件的状态。

## 标准流程（8步）

### Step 0: 理解用户体系（最重要！）

**目标**: 搞清楚用户到底在做什么、卡在哪。

- 读用户论文、毕业论文、实验数据、项目文档
- 搞清楚：什么材料、什么工艺、什么参数、卡在哪
- 提取关键参数：Tg、Td、热压温度/压力/时间、膜厚、透过率等
- **用户说"用不上" → 说明你没理解他的体系，回到 Step 0**
- 将提取的参数写入 `findings.md`

### Step 1: 诊断根因

- 把用户参数和文献中成功案例的参数对比
- 找出差距（如"60°C vs 文献185°C"、"Tg+20°C vs 标准Tg+50-100°C"）
- **这个诊断比下载100篇论文都重要**
- 将诊断结论写入 `findings.md`

### Step 2: 搜索文献 → 调用 deep-research

**使用 `deep-research` skill 执行文献搜索**，但带以下硬约束：

```
调用 deep-research skill，模式：quick brief 或 full research
研究主题：从 Step 0-1 提取的关键词（材料+工艺+性能指标）
```

**必须向 deep-research 传入以下约束**：

1. **期刊白名单**（硬约束，不可放宽）：
   ```
   只搜索以下期刊，其他期刊的结果全部丢弃：
   Nature, Science, Nature Communications, Nature Photonics, Nature Materials,
   Nature Chemistry, Nature Physics, Nature Nanotechnology, Science Advances,
   Angewandte Chemie, JACS, Advanced Materials (AM),
   Advanced Energy Materials (AEM), Advanced Functional Materials (AFM)
   ```

2. **搜索角度**（要求 deep-research 覆盖）：
   - 材料名 + 工艺名组合（中英文）
   - 性能指标 + 期刊名组合
   - 从已知高相关论文的 references/citations 扩展

3. **输出要求**：每篇论文提供 DOI、期刊、作者、年份、摘要、关键性能指标

deep-research 返回结果后，**本 skill 接管后续的筛选、下载、精读、对比、输出**。

> 如果 deep-research skill 不可用 → 降级为 `WebSearch` 多条并行搜索作为备选。

### Step 3: 筛选

在 deep-research 结果基础上做二次筛选：

- **再次确认**所有论文来自白名单期刊（deep-research 可能遗漏非白名单结果）
- 标题/摘要与用户目标相关
- 优先有明确工艺参数的
- 去重（同一篇论文多渠道搜到 → 合并为一条）
- 记录每篇的 DOI、期刊、作者、年份
- 筛选结果写入 `findings.md`

### Step 4: 下载全部（不询问用户）

**绝对禁止在下载前问用户"要下载哪些论文"。所有筛选出的论文全部直接下载。**

#### 4a. 确定下载目录

```bash
# 用当前工作目录，不硬编码路径
PAPERS_DIR="$(pwd)/literature/papers"
mkdir -p "$PAPERS_DIR"
echo "PAPERS_DIR=$PAPERS_DIR"
```

所有下载 Python 脚本中使用此绝对路径。切换设备后 `$(pwd)` 自动适配新环境。

#### 4b. 下载方法总则

**不依赖 playwright MCP**。直接用 Python `websocket-client` 与 CDP 浏览器通信。

**决策树（每篇论文独立执行，最多 3 轮修复重试）：**

```
对每篇论文:
  │
  ├─ 是 OA 论文（Nature OA / Science OA / arXiv / RSC OA）？
  │   └─ YES → curl -L -o "temp_{id}.pdf" "{URL}"  （无需 CDP，用唯一临时名防覆盖）
  │         → 成功？ → ✅ 下一篇
  │         → 失败？ → 尝试 URL 变体，仍失败 → 标记 ❌，继续下一篇
  │
  └─ 是 Paywall 论文（Wiley / ACS / Elsevier）？
      │
      └─ 进入 CDP 下载循环（最多 3 轮，每轮失败后修复重试）:
          │
          ├─ Round 1: 直接尝试下载
          │   ├─ CDP 在线？ → Python websocket 直连 CDP 下载
          │   └─ CDP 不在线？ → 启动 CDP 浏览器，然后下载
          │   → 成功？ → ✅ 下一篇
          │   → 失败？ → 诊断问题，进入 Round 2
          │
          ├─ Round 2: 修复后重试
          │   ├─ 问题 = 403 WebSocket → 重启浏览器加 --remote-allow-origins=*
          │   ├─ 问题 = Cloudflare → 先导航到出版社主页预热 cookies，再试
          │   ├─ 问题 = pdfdirect 返回 HTML 包装页 → 等 embed 加载后再 fetch，或直接 fetch embed src
          │   ├─ 问题 = JS fetch 超时 → 增加页面等待时间到 5s，重试
          │   └─ 问题 = 文件 < 50KB → 检查页面 URL 是否正确，重新导航
          │   → 成功？ → ✅ 下一篇
          │   → 失败？ → 换个方法，进入 Round 3
          │
          ├─ Round 3: 换方法最后尝试
          │   ├─ 尝试 OA 版本 → arXiv 预印本 → ResearchGate
          │   ├─ 尝试直接用 curl 加浏览器 cookies
          │   └─ 尝试 Network.getResponseBody 捕获（可能抓到 embed 子请求）
          │   → 成功？ → ✅ 下一篇
          │   → 失败？ → 标记 ❌，记录原因和尝试过的方法，继续下一篇
          │
          └─ 批量结束后：汇总 ✅/❌

⚠️ 重要：每篇 paywall 论文下载失败时，必须在失败消息中说明：
   ① 第几轮失败  ② 具体症状  ③ 尝试了什么修复  ④ 为什么跳过
```

**核心原则**：
- 一篇失败不阻塞其他论文
- 每篇最多 3 轮修复，不要无限重试
- CDP 下载失败要诊断根因并尝试修复，不要直接标记 ❌
- OA 论文（Nature/Science/arXiv）直接用 curl，不启动浏览器
- 所有 ❌ 论文在最终 Word 报告中列出并说明原因

**时效原则**：
- 单次尝试 timeout ≤ 5s，轮询间隔 0.3s
- 每轮总耗时控制在 30s 内（3 轮共 ≤ 90s/篇）

#### 4c. 按出版社下载策略

| 出版社 | 方法 | 说明 |
|--------|------|------|
| **ACS** (JACS等) | CDP + JS blob | ① 浏览器级 WS SET `Browser.setDownloadBehavior` ② 打开 `pubs.acs.org/doi/pdf/{DOI}` ③ JS `fetch()` blob → `<a>.click()`，`a.download` 用唯一临时名 `temp_{doi后缀}.pdf` ④ rename 到最终名 |
| **Wiley** (AM/AFM/Angew等) | CDP + JS blob | 同上，`a.download = f"temp_{doi_suffix}.pdf"`，URL: `onlinelibrary.wiley.com/doi/pdfdirect/{DOI}` |
| **Nature OA** | curl 直下 | `curl -L -o "temp_{article_id}.pdf" "https://www.nature.com/articles/{id}.pdf"` — OA 论文无 Cloudflare |
| **Science OA** | curl 直下 | 类似 Nature |
| **RSC** | PMC OA 优先 | 先查 PubMed Central，没有再走 CDP |
| **Elsevier** | CDP + JS blob | URL 格式不同，需在文章页找 PDF 按钮 |
| **arXiv** | curl | `curl -L -o "temp_{arxiv_id}.pdf" "https://arxiv.org/pdf/{id}"` |

#### 4d. CDP + JS blob 下载详细步骤（ACS / Wiley 通用）

**Step 4d-1: 初始化 — 浏览器级下载行为**

```python
# 连接浏览器 WebSocket（注意：是 browser 级别，不是 page 级别！）
import json, websocket
from urllib.request import urlopen

# 获取 browser WS URL
resp = urlopen("http://localhost:9223/json/version")
browser_ws = json.loads(resp.read())["webSocketDebuggerUrl"]
ws = websocket.create_connection(browser_ws, timeout=30)

# 设置下载行为（必须在打开任何页面之前！）
ws.send(json.dumps({"id": 1, "method": "Browser.setDownloadBehavior", "params": {
    "behavior": "allow",
    "downloadPath": "<PAPERS_DIR>",  # 绝对路径
    "eventsEnabled": True
}}))
```

**Step 4d-2: 打开 PDF 页面**

```python
from urllib.request import Request
from urllib.parse import quote

pdf_url = f"https://onlinelibrary.wiley.com/doi/pdfdirect/{doi}"  # Wiley
# 或
pdf_url = f"https://pubs.acs.org/doi/pdf/{doi}"  # ACS

# 必须用 PUT 方法！GET 会返回 405
req = Request(f"http://localhost:9223/json/new?{quote(pdf_url, safe='')}", method='PUT')
page_data = json.loads(urlopen(req).read())
page_ws_url = page_data["webSocketDebuggerUrl"]
```

**Step 4d-3: 用 JS fetch blob 触发下载**

```python
ws2 = websocket.create_connection(page_ws_url, timeout=30)

# 等待页面加载 PDF（2s 足够）
time.sleep(2)

# 用 DOI 后缀生成唯一临时文件名，防止并发下载互相覆盖
doi_suffix = target_filename.split('_')[-1].replace('.pdf', '')  # e.g. "anie202422995"
temp_name = f"temp_{doi_suffix}.pdf"

# JS: fetch PDF → blob → createObjectURL → <a>.click()
js = f"""(async () => {{
    const r = await fetch(window.location.href, {{credentials: 'include'}});
    const b = await r.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(b);
    a.download = '{temp_name}';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    return 'OK:' + b.size;
}})()"""
ws2.send(json.dumps({{"id": 1, "method": "Runtime.evaluate",
                      "params": {{"expression": js, "awaitPromise": True, "returnByValue": True}}}}))
# 收一下结果即可，不需要长时间等
time.sleep(1)
ws2.close()
```

**Step 4d-4: 等待下载完成 + 重命名**

```python
# 用短轮询等待唯一的 temp 文件出现，不会与其他下载冲突
temp_path = os.path.join(PAPERS_DIR, temp_name)
filepath = os.path.join(PAPERS_DIR, target_filename)

for _ in range(30):  # 最多等 9s
    time.sleep(0.3)
    if os.path.exists(temp_path) and os.path.getsize(temp_path) > 50000:
        os.rename(temp_path, filepath)
        print(f"  [OK] {target_filename} ({os.path.getsize(filepath)} bytes)")
        break
```

#### 4e. 关键陷阱与教训（必须遵守）

1. **❌ 绝对不要用通用文件名 `paper.pdf`**：多篇论文并发下载时，`paper.pdf` 会互相覆盖，导致文件丢失或内容错乱。必须用唯一临时名（如 `temp_{doi_suffix}.pdf`），下载完成后 rename 到最终名。
   ```python
   # 正确：每篇论文用唯一临时名
   doi_suffix = target_filename.split('_')[-1].replace('.pdf', '')
   temp_name = f"temp_{doi_suffix}.pdf"
   a.download = temp_name  # JS 侧
   
   # 错误：所有论文都用 paper.pdf → 互相覆盖！
   # a.download = 'paper.pdf'  ← 灾难！
   ```

2. **❌ 绝对不要遍历 rename 目录中所有 PDF**：只 rename 与当前论文对应的唯一临时文件。之前的 bug 是用通配符匹配所有 PDF 文件 rename，导致已下载的 JACS 论文被错误 rename 成 Angew 文件名。
   ```python
   # 正确：只 rename 当前论文的临时文件
   if os.path.exists(temp_path) and os.path.getsize(temp_path) > 50000:
       os.rename(temp_path, filepath)  # temp_anie202422995.pdf → final_name.pdf
   
   # 错误：遍历所有 PDF！
   # for f in os.listdir(DIR):
   #     if f.endswith(".pdf") and 'paper' in f:  ← 已下载论文也会被 rename！
   #         os.rename(f, filepath)
   ```

3. **❌ curl 直下 Wiley/ACS paywall**：会被 Cloudflare 拦截，返回 5KB HTML challenge 页面。必须走 CDP 浏览器。

4. **❌ `Network.getResponseBody` 捕获主文档**：Wiley 的 pdfdirect 返回 HTML 包装页 + `<embed>` 标签，主文档 body 只有 348 bytes。真正的 PDF 是 embed 的子请求。用 JS fetch 比拦截 network 更可靠。

5. **❌ 浏览器级 vs 页面级混淆**：`Browser.setDownloadBehavior` 必须在浏览器级 WebSocket 上设置。页面级 WS 不认识这个命令。

6. **❌ `/json/new` 用 GET**：必须 PUT，CDP 文档写的。GET 返回 405。

7. **❌ 忘了 `--remote-allow-origins=*`**：Edge CDP 默认拒绝来自非本地 origin 的 WebSocket 连接。Python websocket 连接会报 403。

8. **❌ 超长 timeout**：每个 recv 等 15-30s 纯属浪费时间。轮询间隔 0.3s，总超时 5-10s 就够了。

9. **❌ 不必要重启浏览器**：重启 = 掉 cookies = 可能需要重新校园网认证。只在 CDP 不可用时重启。

#### 4f. 机构订阅验证（仅认两所学校）

访问出版社主页或文章页后，检查页面是否出现以下**任一**学校英文名：
- `China University of Geosciences`（中国地质大学（武汉））
- `Huazhong University of Science and Technology`（华中科技大学）

**出现任一 → 可下载。两个都没出现 → 该出版社无订阅，不浪费时间重试。**

#### 4g. 下载失败 → 诊断 → 修复 → 重试（最多 3 轮）

每篇论文独立处理，失败先诊断修复，不直接跳过。

**失败诊断表（常见症状 → 根因 → Round N 修复动作）：**

| 症状 | 根因 | Round 1 修复 | Round 2 修复 | Round 3 降级 |
|------|------|-------------|-------------|-------------|
| `curl localhost:9223` 无响应 | CDP 浏览器未启动 | 启动 Edge CDP（见 Step 3） | `taskkill` 杀残余进程后重启 | 所有 paywall 标记 ❌ |
| WebSocket 403 Forbidden | 缺少 `--remote-allow-origins=*` | `taskkill` + 重启浏览器加参数 | 同上，确认参数生效 | 标记 ❌ |
| 页面 title "Just a moment" | Cloudflare WAF | 先导到 `onlinelibrary.wiley.com` 主页，等 3s 预热 cookies，再导到 pdfdirect | 清除 `.lit_survey_browser` 目录重建 session | 换 OA/arXiv |
| JS fetch 返回 `HTTP_ERROR: 403` | 机构认证丢失 | 重新导航到出版社主页，检查是否显示学校名 | 重启浏览器重建 session | 检查 OA 版本 |
| 临时文件 < 50KB（HTML 包装页） | fetch 捕获了 embed 父页面而非 PDF 子请求 | 等 5s 让 embed 加载完，再 fetch；或用 JS 找 `<embed>` src fetch | 直接 fetch `pdfdirect` URL 两次（第二次拿缓存） | `Network.getResponseBody` 逐个试 |
| JS fetch 超时（> 15s 无返回） | PDF 太大或网络慢 | 增加 JS timeout 到 30s | 改用 `Browser.setDownloadBehavior` + 页面 reload 直接触发下载 | curl 加 cookie |
| 页面不显示学校名 | 无此出版社订阅 | 试 OA 版本 | 试 arXiv | 标记 ❌ |
| 临时文件未出现在磁盘上 | `a.download` 文件名冲突或浏览器拦截 blob 下载 | 确认使用了唯一临时文件名 | 改用 `window.location = blobUrl` | 降级为 base64 传回 Python 写盘 |

**三轮修复伪代码：**

```python
def download_with_retry(doi, filename, max_rounds=3):
    for round_num in range(1, max_rounds + 1):
        result = try_download(doi, filename)
        if result.success:
            return result  # ✅
        
        # 诊断 + 修复
        symptom = result.error
        if round_num == 1:
            fix_round1(symptom)  # 重启浏览器 / 预热 cookies / 等 embed 加载
        elif round_num == 2:
            fix_round2(symptom)  # 更激进：重建 session / 换 URL / 改 JS 策略
        else:
            fallback_oa_or_arxiv(doi)  # 降级到 OA / arXiv
            return result  # ❌ 记录原因
    
    return result  # ❌ 3 轮后仍失败
```

**批量下载结束后汇总格式：**

```
=== 下载完成 ===
✅ 6/8 成功
❌ 2/8 失败（均经过 3 轮修复）:
  - 10.1002/anie.2024xxxxx:
      Round 1: Cloudflare → 预热主页 → 仍 blocked
      Round 2: 重建 browser session → 仍 blocked  
      Round 3: OA/arXiv 无 → 跳过
      → 原因: Cloudflare persistent block (Wiley)
  - 10.1038/s41586-024-xxxxx:
      Round 1: CDP 不在线 → 启动浏览器
      Round 2: 启动成功但 fetch 超时 15s
      Round 3: curl + cookie 也失败
      → 原因: PDF too large, timeout
```

❌ 论文在 Step 5 精读时自动跳过，在 Step 7 Word 报告中列出完整修复历史。

#### 4h. 文件命名规范

PDF 保存到 `literature/papers/` 目录。
文件命名：`{year}_{JournalAbbrev}_{FirstAuthor}_{short_title}_{doi_suffix}.pdf`

示例：
- `2024_Angew_OrganogoldIII_Complexes_anie202401833.pdf`
- `2025_JACS_TemperatureAdaptive_Organic_Scintillator_jacs4c12872.pdf`
- `2024_AdvMater_OrganicPhosphorescent_TripletExciton_adma202409338.pdf`

### Step 5: PDF→TXT 全文转换 + 逐篇精读

**全文本转换，非部分提取**：

```bash
# 每篇 PDF → 同名 .txt（依赖 PyMuPDF，前置检测已安装）
$PYTHON pdf_to_txt.py literature/papers/
```

然后**逐篇 Read 精读全文**，不一次性读所有。每篇读完记录：

| 维度 | 要抓取的内容 |
|------|-------------|
| 摘要 | 什么材料、什么体系、什么目标 |
| 实验方法 | **关键工艺参数**（温度/压力/时间/配方/模具） |
| 结果 | 最终性能数据 |
| 讨论 | 成败经验、为什么选这个参数 |
| 与我方差异 | 和用户体系的对比 |

**为什么全文精读而非部分提取？**
- 工艺参数经常藏在 Methods 段，不在 Abstract/Conclusion
- 讨论中的成败经验对方案设计至关重要
- AI 推理判断比脚本预设规则更准确

### Step 6: 对比分析 + 提出可执行方案

- 把用户参数和文献参数放在**同一张表**里对比
- 直观展示差距
- 按优先级排列实验方案（首选→备选）
- 每条含具体参数、常见失败模式及对策

### Step 7: 输出 Word 报告 → 调用 academic-paper

**使用 `academic-paper` skill 生成 Word 报告**，传入 Step 0-6 累积的分析结果。

```yaml
调用参数:
  skill: academic-paper
  mode: full（生成完整 .docx）
  paper_type: 文献调研报告
  content:
    - 问题诊断（来自 Step 0-1 的 findings.md）
    - 文献参数对比表（来自 Step 6 的对比分析）
    - 实验方案（来自 Step 6 的首选+备选方案）
    - 失败模式与对策（来自 Step 6）
    - 参考文献列表（来自 Step 3-4 的 DOI 清单）
  format: DOCX via python-docx / Pandoc
  language: Chinese (中文)
```

**报告必须包含**：
1. 问题诊断（为什么现在不行）
2. 文献参数对比表
3. 实验方案（第1/2/3天做什么，含具体参数）
4. 失败模式与对策
5. 参考文献

> 如果 `academic-paper` skill 不可用 → 降级为 python-docx 直接生成。中文字体注意：新 run 用 `_element.get_or_add_rPr()`。如果 MS Office 可用，可选 `cli-anything-wps` 辅助。

### Step 8: 输出调研 PPT → 调用 ppt-generator

**PPT 是必出的交付物，和 Word 报告同等重要。**

**使用 `ppt-generator` skill 生成**：

```yaml
调用参数:
  skill: ppt-generator
  content:
    数据来源: Step 6 对比分析结果 + Step 7 Word 报告核心结论
```

**定位**：提炼核心信息、可视化关键数据、导向决策。面向"没时间读完整报告"的听众。

**标准 10 页结构**：

| # | 页面 | 内容 | 设计要点 |
|---|------|------|----------|
| 1 | **封面** | 标题：调研主题；副标题：文献调研报告；单位/日期 | 学术蓝底，大标题 60pt |
| 2 | **目录** | 5-6 项 | 左栏编号，右栏标题 |
| 3 | **研究背景** | 一句话说明"做什么、卡在哪" | 顶部蓝色横幅 + 底部统计卡片 |
| 4 | **文献概览** | 3-5 篇关键文献一览 | 三列卡片网格布局 |
| 5 | **核心对比表** | 用户参数 vs 文献参数 | **橙色高亮差距**，绿色标注匹配 |
| 6 | **关键发现** | 3-4 个关键结论 | 大号数字 + 一句话结论 |
| 7 | **实验方案** | 首选 + 备选，含具体参数 | 清晰标注"首选"和"备选" |
| 8 | **失败模式与对策** | 常见失败及应对 | 左列失败，右列对策 |
| 9 | **结论与建议** | 调研结论 + 下一步行动 | 顶部横幅 + 底部 timeline |
| 10 | **参考文献** | 5-10 条核心文献 | 期刊名加粗 |

**配色方案**：
```
PRIMARY_BLUE  = (26, 60, 139)     # 背景、横幅
ACCENT_ORANGE = (230, 119, 51)    # 高亮差距
ACCENT_GREEN  = (24, 128, 80)     # 成功方案
WHITE         = (255, 255, 255)
DARK          = (34, 34, 34)
LIGHT_BG      = (245, 248, 252)   # 卡片背景
```

> 如果 `ppt-generator` skill 不可用 → 降级为 python-pptx。

## 绝对禁止

- ❌ 下载前问用户"要下载哪些"
- ❌ 搜索非白名单期刊
- ❌ 用 curl/requests 直接下载 paywall 论文（必须走 CDP 浏览器）
- ❌ 只下载部分论文
- ❌ 不写 Word 报告
- ❌ 不写 PPT
- ❌ 一上来就搜论文（先理解再搜索）
- ❌ 遍历 rename 目录中所有 PDF 文件（只能 rename 当前论文对应的唯一临时文件 `temp_{doi后缀}.pdf`）
- ❌ 在两所学校之外猜测或匹配其他大学名称
- ❌ 长时间 timeout 等待（每次 recv/poll ≤ 5s）

## 工具对照表

| 用途 | 工具 |
|------|------|
| 流程管控 | `Skill` → `planning-with-files`（task_plan + findings + progress） |
| 文献搜索 | `Skill` → `deep-research`（约束白名单期刊） |
| 论文信息获取 | `WebFetch`（备用：获取摘要、作者、DOI等） |
| 浏览器自动化下载 | `Bash` + Python `websocket-client` 直连 CDP 9223（不依赖 playwright MCP） |
| OA 论文下载 | `Bash` + `curl -L -o` |
| PDF→TXT 转换 | `Bash` + `$PYTHON pdf_to_txt.py`（脚本自动创建） |
| 全文阅读 | `Read`（逐篇读，不批量） |
| Word 报告生成 | `Skill` → `academic-paper`（降级：python-docx） |
| PPT 生成 | `Skill` → `ppt-generator`（降级：python-pptx） |
| 并行搜索 | deep-research 内部并行 agents |

## Skill 调用链总览

```
用户说"深度调研XX"
  │
  ├─ planning-with-files  ─── 创建 task_plan / findings / progress（贯穿全程）
  │
  ├─ Step 0-1: 本 skill  ─── 理解体系 + 诊断根因
  │
  ├─ Step 2: deep-research  ─── 搜索文献（白名单期刊约束）
  │
  ├─ Step 3-6: 本 skill  ─── 筛选 + CDP下载 + 精读 + 对比分析
  │
  ├─ Step 7: academic-paper  ─── 生成 Word 报告
  │
  └─ Step 8: ppt-generator  ─── 生成 PPT
```

## 环境信息说明

- **Python**：自动检测，尝试 `python3` → `python` → `py` → 常见安装路径
- **浏览器**：自动检测 Edge（Windows）或 Chrome（macOS），CDP 端口 9223
- **CDP 关键参数**：`--remote-debugging-port=9223 --remote-allow-origins="*" --user-data-dir="$HOME/.lit_survey_browser"`
- **论文保存目录**：`literature/papers/`（绝对路径，避免 cd 混乱）
- **依赖自动安装**：PyMuPDF、python-docx、python-pptx、websocket-client（pip install --quiet）
- **校园网**：仅识别中国地质大学（武汉）和华中科技大学，不猜第三所
- **下载方法**：Python `websocket-client` 直连 CDP（不依赖 playwright MCP）
