---
name: office-comparer
description: 提供了基于 Streamlit 的文件（Word/Excel）与图片（PNG/JPG）的高性能并排比较与交互式编辑功能。支持字符级差异高亮、表格对齐比对以及基于 CV2 的图像自动配准。
---

# office-comparer 技能使用指南

此技能允许 Gemini CLI 部署并运行一个专业的办公级比对工具 **Office-Comparer Pro**。

## 核心功能

1.  **Word 文档比对 (Pro)**：
    *   **字符级高亮**：精确识别段落内部的增删改，使用删除线与颜色标记。
    *   **实时同步**：支持单行或全量接受 A/B 版本，修改后可直接导出 `.docx`。
2.  **Excel 表格比对 (New)**：
    *   **Side-by-Side 视图**：基于 Pandas 对齐两个表格。
    *   **单元格差异标记**：自动以黄色背景高亮所有内容不一致的单元格。
3.  **高级图像配准 (Enhanced)**：
    *   **自动对齐**：使用 OpenCV ORB 算法自动纠正图片间的位移与旋转。
    *   **像素级热图**：支持透明叠加与灵敏度调节，适配高清图片。
4.  **性能优化**：
    *   **智能缓存**：采用 `@st.cache_data`，大文件处理零卡顿。

## 启动指令

在工作区运行以下指令启动 Pro 版工具：
```bash
python -m streamlit run app.py --server.port 8501 --server.headless true
```

## 资源清单

- `app.py`: 核心 Streamlit UI 与业务逻辑程序。
- `requirements.txt`: 包含 `opencv-python`, `python-docx`, `pandas` 等必需依赖。

## 环境要求

确保已安装以下依赖：
`streamlit`, `python-docx`, `pandas`, `openpyxl`, `pillow`, `numpy`, `plotly`, `opencv-python`, `jinja2`
