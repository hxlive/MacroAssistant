# MacroAssistant - 宏助手

![Python Version](https://img.shields.io/badge/python-3.9%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![UI Framework](https://img.shields.io/badge/UI-ttkbootstrap-brightgreen.svg)
![Core](https://img.shields.io/badge/Automation-PyAutoGUI_&_OpenCV-orange.svg)


## 📖 项目简介

**宏助手** 是一款开源免费的**桌面自动化工具**，专为简化重复性任务而设计。无论是游戏自动化、办公流程优化，还是 UI 测试，宏助手通过**图像识别**和 **OCR 文字识别**，帮助用户快速录制和执行自动化脚本。

无需编程经验，通过直观的 GUI 界面即可轻松上手。

经过多次迭代，优化了性能和稳定性，提供快速响应和高精度识别，适合个人用户和开发者使用。

<p align="center">
  <img src="screenshot.png" alt="Macro Assistant 截图" width="850">
</p>

## 核心功能

* **⚡ 图像识别 (OpenCV)**：通过 `cv2.matchTemplate` 实现毫秒级超快图像查找，支持多尺度和自定义置信度（`confidence`）。
* **🔤 多引擎 OCR**：智能集成 `WinOCR`（Windows 10+ 内置，最快）、`RapidOCR`（深度学习，精准）和 `Tesseract`（兜底），中英文识别准确率高。
* **🖱️ 完整键鼠模拟**：支持点击、移动（绝对/相对）、滚动滚轮、等待、输入文本（支持中文粘贴）和按下按键（支持组合键如 `ctrl+c`）。
* **🎛️ 高级流程控制**：支持 `IF/ELSE` 条件判断（基于图像/文本）和 `LOOP` 循环，可构建复杂的自动化逻辑。
* **🖥️ 窗口管理**：支持查找并**激活指定窗口**（按标题），确保自动化在正确的目标上执行。
* **✨ 现代 GUI**：基于 `ttkbootstrap` 构建，提供清晰的步骤编辑器、实时状态显示、主题切换和最近文件列表。
* **⌨️ 全局热键**：使用 `Ctrl+F10` 启动，`Ctrl+F11` 紧急停止，安全可控。
* **🎯 辅助工具**：
    * **实时坐标**：在添加“移动到”步骤时，实时显示当前鼠标坐标。
    * **即时测试**：在添加步骤前，可立即测试“查找图像”或“查找文本”是否有效。

## 适用场景

* **🎮 游戏自动化**：自动点击、刷任务、切换窗口。
* **💼 办公效率**：批量填写表单、数据录入、窗口操作。
* **🐛 测试工具**：模拟用户交互，验证 UI 功能。
* **🔄 日常重复任务**：自动处理固定流程，提升效率。

---

## 📚 动作列表详解

| 序号 | 动作 | 功能描述 |
| :--- | :--- | :--- |
| **01** | **查找图像** | (核心) 在全屏查找指定图像，找到后将鼠标移动到其中心。 |
| **02** | **查找文本 (OCR)** | (核心) 在全屏查找指定文本，找到后将鼠标移动到其中心。 |
| **03** | **相对移动** | 从*上一个*动作的位置，按偏移量移动鼠标。 |
| **04** | **移动到 (绝对坐标)** | 将鼠标移动到屏幕的精确坐标 (提供实时坐标参考)。 |
| **05** | **点击鼠标** | 在当前鼠标位置执行点击 (左/中/右键)。 |
| **06** | **滚动滚轮** | 在指定位置（或当前位置）滚动。 |
| **07** | **等待** | 暂停宏，等待固定时间 (毫秒)。 |
| **08** | **输入文本** | 模拟键盘输入（通过剪贴板粘贴实现，支持中文）。 |
| **09** | **按下按键** | 模拟按下单个按键或组合键 (如 `enter`, `ctrl+v`)。 |
| **10** | **激活窗口 (按标题)** | 查找标题包含指定文本的窗口，并将其激活到最前端。 |
| **11** | **IF 找到图像** | (流程控制) 如果找到该图像，则继续执行 `IF` 块。 |
| **12** | **IF 找到文本** | (流程控制) 如果找到该文本，则继续执行 `IF` 块。 |
| **13** | **ELSE** | (流程控制) 必须与 `IF` 配对使用。 |
| **14** | **END_IF** | (流程控制) 标记 `IF` 块的结束。 |
| **15** | **循环开始 (Loop)** | 指定循环体执行的次数。 |
| **16** | **结束循环 (EndLoop)**| 标记循环体的结束。 |

## --## 🛠️ 安装与依赖

### 1. 库依赖
    ** `requirements.txt` 内容:**
    ```
    ttkbootstrap
    pyautogui
    pynput
    pygetwindow
    opencv-python
    Pillow
    pytesseract
    rapidocr-onnxruntime
    windows-ocr
    ```

请确保已使用 pip 安装所有这些库。

### 2. OCR 引擎配置
* **WinOCR**: 仅支持 Windows 10 及以上版本，无需配置。
* **RapidOCR**:
    * `pip install rapidocr-onnxruntime` 即可。
    * **重要**：如果启动时或运行时提示 `DLL load failed` 或初始化失败，你**必须**安装最新的 [Microsoft Visual C++ Redistributable](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist) 运行库。
* **Tesseract**:
    * 这是备用引擎，你需要单独安装它。
    * **方式一 (推荐)**: 解压 `tesseract_local.7z`，确保 `tesseract_local` 文件夹（包含 `tesseract.exe` 和 `tessdata`）与 `.py` 脚本位于同一目录。
    * **方式二 (全局)**: 安装 Tesseract，并将其安装路径添加到系统的 `PATH` 环境变量中。

##
### ⌨️ 默认快捷键控制
- **Ctrl+F10**：启动宏执行
- **Ctrl+F11**：立即停止宏
- 避免与其他程序冲突，随时掌控

### 
## ⚠️ 注意事项

### 安全性
- ✅ **完全离线运行**，无任何网络请求
- ✅ **数据安全**，所有宏文件存储在本地
- ✅ **开源透明**，代码公开可审查

### 开源协议
- 📜 采用 **MIT 协议**，可自由修改和分享
- 📜 商业使用、私人使用均可
- 📜 需保留版权声明

### 调试支持
- 📊 控制台输出详细执行日志
- 📊 `macro_perf.log` 记录性能统计数据
- 📊 支持调试模式，实时查看识别结果

### 依赖问题
- 🔧 若 RapidOCR 初始化失败，请安装 [VC++ 运行库](https://aka.ms/vs/17/release/vc_redist.x64.exe)
- 🔧 若 WinOCR 不可用，请确保系统为 Windows 10 1903+
- 🔧 若 OpenCV 导入失败，尝试重新安装：`pip install opencv-python --upgrade`

## 📄 许可证

本项目基于 [MIT License](LICENSE) 开源。