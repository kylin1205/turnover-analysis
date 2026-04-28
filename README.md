# 员工离职数据分析系统

一个基于 Streamlit 的员工离职数据分析可视化工具。

## 部署到 Streamlit Cloud

### 步骤 1: 创建 GitHub 仓库

1. 登录 GitHub (https://github.com)
2. 点击右上角 "+" → "New repository"
3. 仓库名称填: `turnover-analysis`
4. 选择 Private 或 Public
5. 点击 "Create repository"

### 步骤 2: 上传代码

在仓库页面:
1. 点击 "uploading an existing file"
2. 上传以下文件:
   - `streamlit_app.py`
   - `requirements.txt`
   - `.gitignore`
   - `README.md`
3. 点击 "Commit changes"

### 步骤 3: 部署到 Streamlit Cloud

1. 登录 https://share.streamlit.io
2. 点击 "New app"
3. 配置:
   - Repository: `你的用户名/turnover-analysis`
   - Branch: `main`
   - Main file path: `streamlit_app.py`
4. 点击 "Deploy!"

部署成功后，你会获得一个 URL，如: `https://你的用户名-turnover-analysis.streamlit.app`

## 本地运行

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## 功能

- 上传 Excel 文件（包含人员清单和离职数据）
- 按月份分析离职率
- 各部门离职率对比（柱状图）
- 离职类型分布
- 离职司龄分析
- 离职原因统计
- 职级分布
- 导出 Excel 报告

## 数据格式要求

Excel 文件需包含:
- **人员清单 Sheet**: 包含 `姓名`、`身份类别`、`一级组织` 等列
- **离职数据 Sheet**: 包含 `离职类型`、`离职原因`、`累计司龄（年）`、`职级`、`最后工作日` 等列
