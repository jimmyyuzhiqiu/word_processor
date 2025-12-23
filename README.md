# Word 格式炼化器（PyQt5 + Word COM）

一个在 **Windows** 上运行的图形化工具，支持拖拽 `.doc/.docx` 文件批量处理，自动清理文档中的 **Tab、连续空格、全角空格/nbsp**，并将 **“假列表”**（例如 `1. `、`（2）`、`-`、`•`）转换为 **Word 的真列表格式**。支持处理 **正文与页眉/页脚**，并可按不同策略命名输出文件。

> 💡 应用基于 **PyQt5** 提供 UI，使用 **win32com（Word COM）** 操作文档内容。线程内调用 `pythoncom.CoInitialize()`，避免多线程 COM 初始化错误。

---

### 目录结构

```
word_processor/
├─ app.py                 # 图形界面（PyQt5）
├─ word_processor.py      # 文档处理核心逻辑（win32com）
├─ ing-logo.png           # 应用 Logo（可选）
├─ app.ico                # 应用图标（可选）
├─ requirements.txt       # 依赖清单（建议）
└─ README.md              # 项目说明（本文）
```

---

## ✨ 功能特性

- **拖拽操作**：支持拖入 `.doc`/`.docx` 文件，或拖入文件夹自动读取其中的 Word 文件
- **批量处理**：一次可处理多个文件
- **空白清理**：
  - Tab → 空格（可选）
  - 连续空格压缩（多空格合并为 1 个，含全角空格/nbsp）
  - 连续空行压缩（最多保留 N 行）
- **假列表 → 真列表**：
  - 数字型前缀：`1.`、`2)`、`（3）`、`1、` 等
  - 项目符号：`-`、`•`、`*` 等  
  自动转换为 Word 原生编号/项目格式并 **连续衔接** 同一段列表
- **页眉/页脚处理**（可选）
- **输出策略**：
  - 覆盖模式：与原文件同名（按输出目录保存）
  - 后缀模式：原名 + 后缀（默认 `_cleaned`）
  - 自定义模式：仅单文件可用
- **输出格式**：`.docx（推荐）` 或 `.doc`
- **UI 风格**：深色霓虹科技风 QSS，自带圆角卡片、金色脚注

---

## 🚀 快速开始

> ✅ **系统要求**：**Windows**（必须安装 Microsoft Word）  
> ✅ **Python**：3.9+（建议 3.10/3.11）  
> ✅ **依赖**：`PyQt5`、`pywin32`、`pythoncom`（pywin32 附带）

### 1）克隆或下载项目

```bash
git clone https://github.com/<your-username>/word_processor.git
cd word_processor
```

### 2）创建虚拟环境并安装依赖

```bash
# 创建并激活虚拟环境（Windows PowerShell）
python -m venv .venv
.venv\Scripts\Activate.ps1

# 安装依赖
pip install -r requirements.txt
```

建议的 **requirements.txt** 内容：

```txt
PyQt5>=5.15
pywin32>=306
```

> 如果没有 `requirements.txt`，也可以：
> ```bash
> pip install PyQt5 pywin32
> ```

### 3）运行

```bash
python app.py
```

首次运行时，如果你没有放置 `ing-logo.png` 或 `app.ico` 在同目录，UI 会显示 “Logo 未找到”，不影响功能。

---

## 🧠 使用说明（GUI）

1. 打开程序后，将 `.doc/.docx` 文件 **拖入左侧列表**，或点击【➕ 添加文件】
2. 在右侧配置区设置：
   - **输出策略**：覆盖 / 后缀 / 自定义（单文件）
   - **输出位置**：原目录或选择输出目录
   - **清理规则**：Tab→空格、压缩空格、处理页眉/页脚、连续空行最多保留
   - **输出格式**：`.docx` 或 `.doc`
3. 点击【⚡ 一键炼化 / 开始处理】
4. 处理日志会显示在下方，进度条实时更新。完成后弹窗提示。

---

## 🛠️ 打包成可执行文件（PyInstaller）

本项目中的 `resource_path()` 已兼容 **PyInstaller onefile** 模式。打包命令示例：

```bash
pip install pyinstaller

pyinstaller ^
  --noconsole ^
  --onefile ^
  --name "WordCleanerUI" ^
  --add-data "ing-logo.png;." ^
  --icon "app.ico" ^
  app.py
```

- **`--add-data`**：将资源文件（logo、图标）打入包中  
- **`--icon`**：使用应用图标  
- 打包后生成的可执行文件在 `dist/WordCleanerUI.exe`

> 若你使用的是中文路径或网络盘路径，建议将项目放到英文目录，避免 PyInstaller 在路径编码上出现异常。

---

## ⚙️ 关键实现点（开发者向）

- **COM 初始化**：在 `QThread` 的 `run()` 中调用：
  ```python
  pythoncom.CoInitialize()
  ...  # 调用 win32com 操作 Word
  pythoncom.CoUninitialize()
  ```
  这是避免 “CoInitialize has not been called” 报错的关键。

- **COM 对象遍历**：Word 的集合使用 **1-based 索引**，用 `Count + Item(i)` 更稳：
  ```python
  paras = range_obj.Paragraphs
  for i in range(1, paras.Count + 1):
      p = paras.Item(i)
  ```

- **真列表连续性**：首项用 `ApplyNumberDefault` / `ApplyBulletDefault`，后续用 `ApplyListTemplate(..., ContinuePreviousList=True)` 保持同一个列表。

- **页眉/页脚**：通过 `doc.Sections(si).Headers(1)` 和 `Footers(1)` 处理 **Primary** 区域，异常用 `try/except` 忽略，保证鲁棒性。

- **保存格式**：  
  - `.docx` → `FileFormat=12 (wdFormatXMLDocument)`  
  - `.doc` → `FileFormat=0 (wdFormatDocument)`

---

## 📦 配置与持久化

- 使用 `QSettings("MY43DN", "WordCleanerUI_Neon")` 存储：
  - 上次打开目录、上次输出目录
  - 后缀、自定义名等 UI 参数
- 资源加载优先级：
  1. **同目录资源**（`resource_path("ing-logo.png")`）
  2. 绝对路径兜底（`C:\Users\MY43DN\Desktop\ing-logo.png` 等）

---

## 🔒 平台与限制

- **仅支持 Windows**（依赖 `win32com` 和 **本机安装 Microsoft Word**）
- 运行期间会后台启动 Word 进程，程序退出时会调用 `word.Quit()`，确保不残留
- 对页眉/页脚的处理仅覆盖 **Primary** 类型，若文档使用不同页眉/页脚或奇偶页不同，需扩展 `Headers(Footer)` 索引

---

## ❓常见问题（FAQ）

**Q1：运行时报错 `CoInitialize has not been called` 怎么办？**  
A：已在工作线程 `Worker.run()` 中调用了 `pythoncom.CoInitialize()`。请确保不要在非该线程中操作 COM 对象。

**Q2：提示找不到 `win32com.client`？**  
A：安装 `pywin32`：  
```bash
pip install pywin32
```
安装后如仍有问题，尝试运行：
```bash
python -m pip install --upgrade pip
python -m pip install --upgrade pywin32
```

**Q3：打包后 Logo/图标不显示？**  
A：确保在打包时使用了 `--add-data` 与 `--icon`，并且资源文件在与你的命令一致的路径下。也可放置到与 `app.exe` 同目录作为兜底。

**Q4：为什么有些列表没有被转成真列表？**  
A：假列表的检测基于正则：
- 数字：`^\s*(?:\d+\s*[.)、]|[\(\（]\s*\d+\s*[\)\）])\s+`
- 项目符号：`^\s*[-–—•●·*]\s+`  
如果你的文档前缀形式不在这些模式中，可扩展正则。

---

## 📝 .gitignore（建议）

在项目根目录添加 `.gitignore`，避免提交缓存与打包产物：

```
# Python
__pycache__/
*.pyc
.venv/
env/
venv/

# PyInstaller
build/
dist/
*.spec

# IDE
.vscode/
.idea/

# Logs
*.log
```

---

## 🤝 贡献与反馈

欢迎提交 Issue 或 Pull Request 来：
- 增强假列表模式识别
- 支持更多 Header/Footer 类型
- 增加 UI 国际化与快捷键
- 增加批处理报错可视化（红色日志行/导出错误报告）

---


---

## 🧪 测试文件

项目中已包含一个示例文件：`test.docx`，位于 `word_processor/test.docx`。

你可以用它快速验证程序功能：

```bash
python app.py
```

拖入 `test.docx`，选择默认配置，点击【⚡ 一键炼化】，查看输出效果。


## 📄 许可证

MIT License

Copyright (c) 2025 余智秋

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
