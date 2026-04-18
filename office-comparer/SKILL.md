---
name: office-comparer
description: 提供了基于 Streamlit 的文件（Word/Excel/Code）与图片（PNG/JPG）的高性能并排比较与交互式编辑功能。支持万能文档对比、Minimap 导航、黑暗模式以及基于 CV2 的图像自动配准。
---

# office-comparer 技能使用指南

此技能允许 Gemini CLI 部署并运行一个专业的全能比对工具 **Office-Comparer Pro**。

## 核心功能

1.  **万能文档对比 (Universal Doc Diff)**：
    *   **多格式支持**：兼容 `.docx`, `.py`, `.c`, `.cpp`, `.txt`, `.md` 等多种办公与代码格式。
    *   **代码优化**：内置 `Fira Code` 等宽字体，完整保留缩进，适配代码审查。
    *   **查看模式**：支持“左右双栏”与“混合视图（Unified View）”一键切换。
2.  **专业 UI/UX 特性**：
    *   **Minimap 导航**：侧边栏实时渲染全文差异缩略图，快速定位修改点。
    *   **黑暗模式 (Dark Mode)**：一键切换深色主题，保护视力。
    *   **毛玻璃特效**：现代化的 Glassmorphism 界面设计。
3.  **Excel 表格对比**：
    *   **单元格高亮**：自动识别并标记表格中的数值差异。
4.  **高级图像配准 (CV2 Powered)**：
    *   **自动对齐**：使用 OpenCV ORB 算法自动纠正图片间的位移与旋转。
    *   **动态覆盖**：支持手动调节差异灵敏度与覆盖强度（Alpha）。

## 启动指令

在工作区运行以下指令启动 Pro 版工具：
```bash
python -m streamlit run app.py --server.port 8501 --server.headless true
```

## 资源清单

- `app.py`: 核心 Streamlit UI 与业务逻辑程序。
- `requirements.txt`: 包含 `opencv-python`, `python-docx`, `pandas`, `plotly` 等必需依赖。

## 环境要求

确保已安装以下依赖：
`streamlit`, `python-docx`, `pandas`, `openpyxl`, `pillow`, `numpy`, `plotly`, `opencv-python`, `jinja2`
