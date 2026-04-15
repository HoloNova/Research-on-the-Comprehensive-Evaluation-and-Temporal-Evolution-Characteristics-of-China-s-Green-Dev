import os
import shutil

def organize_project():
    print("="*80)
    print("项目文件夹系统性整理")
    print("="*80)
    
    folders_to_create = ['docs', 'reports']
    for folder in folders_to_create:
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"✓ 创建文件夹: {folder}")
    
    report_files_to_move = [
        '数据字典.json',
        '熵权法权重分析报告.json',
        '熵权法权重分析报告.md',
        'TOPSIS综合评价报告.json',
        'TOPSIS综合评价报告.md',
        '全面审计报告.json',
        '全面审计报告.md',
        '全面文件检查报告.json',
        '全面文件检查报告.md'
    ]
    
    moved_reports = 0
    for report_file in report_files_to_move:
        if os.path.exists(report_file):
            shutil.move(report_file, os.path.join('reports', report_file))
            moved_reports += 1
            print(f"✓ 移动报告: {report_file} -> reports/")
    
    doc_files_to_move = [
        '使用说明.md',
        '项目整理总结.md',
        '文件检查与TOPSIS分析完成总结.md',
        '全面审计完成总结.md',
        '文件演化可视化模板.md'
    ]
    
    moved_docs = 0
    for doc_file in doc_files_to_move:
        if os.path.exists(doc_file):
            shutil.move(doc_file, os.path.join('docs', doc_file))
            moved_docs += 1
            print(f"✓ 移动文档: {doc_file} -> docs/")
    
    if os.path.exists(os.path.join('code', '__pycache__')):
        shutil.rmtree(os.path.join('code', '__pycache__'))
        print("✓ 清理: code/__pycache__")
    
    gitignore_content = '''# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg

# Virtual Environment
venv/
ENV/
env/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
Thumbs.db

# Temporary files
*.tmp
*.bak
*.log

# Data (optional - uncomment if you don't want to commit data)
# *.xlsx
# *.csv
# 国民消费水平/
# 科技/
# 能源/
# 资源与环境/
'''
    
    if not os.path.exists('.gitignore'):
        with open('.gitignore', 'w', encoding='utf-8') as f:
            f.write(gitignore_content)
        print("✓ 创建: .gitignore")
    
    readme_content = '''# 统计建模项目

## 项目简介

本项目是一个统计建模数据分析项目，包含数据处理、综合评价分析等功能。

## 项目结构

```
统计建模/
├── code/                          # 代码文件夹
├── docs/                          # 文档文件夹
├── reports/                       # 报告文件夹
├── 国民消费水平/                   # 国民消费水平数据
├── 科技/                           # 科技数据
├── 能源/                           # 能源数据
├── 资源与环境/                     # 资源与环境数据
├── .gitignore                     # Git忽略文件
├── README.md                      # 本文档
└── 文件夹查看指南.md               # 详细的文件夹查看指南
```

## 快速开始

### 数据处理

```bash
# 快速模式
python code/final_processor.py

# 或选择模式
python code/run_processor.py
```

### 综合评价

```bash
python code/check_and_topsis.py
```

## 文档

详细说明请查看 `文件夹查看指南.md`。

## 许可证

本项目仅供学习和研究使用。
'''
    
    if not os.path.exists('README.md'):
        with open('README.md', 'w', encoding='utf-8') as f:
            f.write(readme_content)
        print("✓ 创建: README.md")
    
    print("\n" + "="*80)
    print("整理完成！")
    print("="*80)
    print(f"\n创建文件夹: {len(folders_to_create)}个")
    print(f"移动报告文件: {moved_reports}个")
    print(f"移动文档文件: {moved_docs}个")
    print("\n项目结构已优化，具备可扩展性！")

if __name__ == '__main__':
    organize_project()
