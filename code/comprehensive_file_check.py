import os
import json
import pandas as pd
import numpy as np
import ast
import re
from datetime import datetime
from typing import List, Dict, Any, Tuple

class ComprehensiveFileChecker:
    def __init__(self):
        self.check_results = {
            'summary': {},
            'issues': [],
            'file_checks': {},
            'coupling_analysis': {}
        }
        self.severity_levels = {
            'critical': 3,
            'high': 2,
            'medium': 1,
            'low': 0
        }
    
    def log_issue(self, severity: str, category: str, file_path: str, description: str, details: str = None):
        issue = {
            'timestamp': datetime.now().isoformat(),
            'severity': severity,
            'category': category,
            'file_path': file_path,
            'description': description,
            'details': details
        }
        self.check_results['issues'].append(issue)
    
    def check_json_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'issues': []
        }
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if not isinstance(data, (dict, list)):
                result['is_valid'] = False
                result['issues'].append('根节点应该是对象或数组')
                self.log_issue('medium', 'JSON格式', file_path, 'JSON根节点类型不正确', '应该是对象或数组')
            
            if 'results' in locals() and isinstance(data, dict):
                for key in ['report_title', 'created_at', 'results']:
                    if key not in data:
                        result['issues'].append(f'缺少必要字段: {key}')
                        self.log_issue('medium', '内容完整性', file_path, f'缺少必要字段: {key}')
            
        except json.JSONDecodeError as e:
            result['is_valid'] = False
            result['issues'].append(f'JSON解析错误: {str(e)}')
            self.log_issue('critical', 'JSON格式', file_path, 'JSON解析失败', str(e))
        except Exception as e:
            result['is_valid'] = False
            result['issues'].append(f'读取错误: {str(e)}')
            self.log_issue('high', '文件损坏', file_path, '文件读取失败', str(e))
        
        return result['is_valid'], result
    
    def check_python_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'issues': [],
            'imports': [],
            'defined_functions': [],
            'defined_classes': []
        }
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            try:
                tree = ast.parse(content)
                
                for node in ast.walk(tree):
                    if isinstance(node, ast.Import):
                        for name in node.names:
                            result['imports'].append(name.name)
                    elif isinstance(node, ast.ImportFrom):
                        result['imports'].append(node.module or '')
                    elif isinstance(node, ast.FunctionDef):
                        result['defined_functions'].append(node.name)
                    elif isinstance(node, ast.ClassDef):
                        result['defined_classes'].append(node.name)
                
                main_found = any(
                    isinstance(node, ast.If) and
                    isinstance(node.test, ast.Compare) and
                    isinstance(node.test.left, ast.Name) and
                    node.test.left.id == '__name__'
                    for node in ast.walk(tree)
                )
                
                if not main_found and len(result['defined_functions']) > 0:
                    result['issues'].append('缺少__main__入口点')
                    self.log_issue('low', '代码结构', file_path, '缺少__main__入口点')
            
            except SyntaxError as e:
                result['is_valid'] = False
                result['issues'].append(f'语法错误: 第{e.lineno}行, {e.text}')
                self.log_issue('critical', 'Python语法', file_path, '语法错误', f'第{e.lineno}行')
            
        except Exception as e:
            result['is_valid'] = False
            result['issues'].append(f'读取错误: {str(e)}')
            self.log_issue('high', '文件损坏', file_path, '文件读取失败', str(e))
        
        return result['is_valid'], result
    
    def check_excel_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'issues': [],
            'sheets': [],
            'shape': None
        }
        
        try:
            xl = pd.ExcelFile(file_path)
            result['sheets'] = xl.sheet_names
            
            df = pd.read_excel(file_path)
            result['shape'] = df.shape
            
            if len(df) == 0:
                result['issues'].append('Excel文件为空')
                self.log_issue('medium', '内容完整性', file_path, 'Excel文件为空')
            
            if len(df.columns) == 0:
                result['issues'].append('没有列数据')
                self.log_issue('medium', '内容完整性', file_path, '没有列数据')
            
        except Exception as e:
            result['is_valid'] = False
            result['issues'].append(f'读取错误: {str(e)}')
            self.log_issue('high', '文件损坏', file_path, 'Excel读取失败', str(e))
        
        return result['is_valid'], result
    
    def check_csv_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'issues': [],
            'shape': None
        }
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            result['shape'] = df.shape
            
            if len(df) == 0:
                result['issues'].append('CSV文件为空')
                self.log_issue('medium', '内容完整性', file_path, 'CSV文件为空')
            
        except Exception as e:
            result['is_valid'] = False
            result['issues'].append(f'读取错误: {str(e)}')
            self.log_issue('high', '文件损坏', file_path, 'CSV读取失败', str(e))
        
        return result['is_valid'], result
    
    def check_markdown_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'issues': [],
            'has_title': False,
            'sections': 0
        }
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            if content.strip().startswith('#'):
                result['has_title'] = True
            
            result['sections'] = len(re.findall(r'^#{1,6}\s', content, re.MULTILINE))
            
            if len(content.strip()) == 0:
                result['is_valid'] = False
                result['issues'].append('Markdown文件为空')
                self.log_issue('medium', '内容完整性', file_path, 'Markdown文件为空')
            
        except Exception as e:
            result['is_valid'] = False
            result['issues'].append(f'读取错误: {str(e)}')
            self.log_issue('high', '文件损坏', file_path, 'Markdown读取失败', str(e))
        
        return result['is_valid'], result
    
    def analyze_coupling(self, code_files: List[str]) -> Dict:
        coupling_result = {
            'file_dependencies': {},
            'coupling_score': 0.0,
            'recommendation': ''
        }
        
        all_imports = {}
        
        for file_path in code_files:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                imports = []
                try:
                    tree = ast.parse(content)
                    for node in ast.walk(tree):
                        if isinstance(node, ast.Import):
                            for name in node.names:
                                imports.append(name.name)
                        elif isinstance(node, ast.ImportFrom):
                            imports.append(node.module or '')
                except:
                    pass
                
                all_imports[file_path] = imports
                coupling_result['file_dependencies'][file_path] = imports
                
            except:
                pass
        
        total_files = len(code_files)
        if total_files > 0:
            external_imports = sum(1 for imps in all_imports.values() 
                                   for imp in imps 
                                   if not any(imp.startswith(prefix) for prefix in ['os', 'sys', 'json', 'pandas', 'numpy', 'datetime', 'typing', 'ast', 're']))
            
            coupling_score = external_imports / (total_files * 10) if total_files > 0 else 0
            coupling_result['coupling_score'] = min(coupling_score, 1.0)
            
            if coupling_result['coupling_score'] < 0.3:
                coupling_result['recommendation'] = '耦合度低，模块独立性好'
            elif coupling_result['coupling_score'] < 0.6:
                coupling_result['recommendation'] = '耦合度适中，结构合理'
            else:
                coupling_result['recommendation'] = '耦合度较高，建议考虑模块化重构'
        
        return coupling_result
    
    def run_check(self):
        print("="*80)
        print("全面文件完整性与功能性检查")
        print("="*80)
        
        folders_to_check = ['.', 'code', '国民消费水平', '科技', '能源', '资源与环境']
        files_to_check = []
        
        for folder in folders_to_check:
            folder_path = os.path.join('.', folder)
            if os.path.exists(folder_path):
                for root, dirs, files in os.walk(folder_path):
                    for file in files:
                        if not file.startswith('~$'):
                            files_to_check.append(os.path.join(root, file))
        
        print(f"\n找到 {len(files_to_check)} 个文件进行检查\n")
        
        code_files = []
        valid_count = 0
        invalid_count = 0
        
        for file_path in files_to_check:
            ext = os.path.splitext(file_path)[1].lower()
            is_valid = True
            check_details = {}
            
            print(f"检查: {file_path}")
            
            if ext == '.json':
                is_valid, check_details = self.check_json_file(file_path)
            elif ext == '.py':
                is_valid, check_details = self.check_python_file(file_path)
                code_files.append(file_path)
            elif ext == '.xlsx':
                is_valid, check_details = self.check_excel_file(file_path)
            elif ext == '.csv':
                is_valid, check_details = self.check_csv_file(file_path)
            elif ext == '.md':
                is_valid, check_details = self.check_markdown_file(file_path)
            
            self.check_results['file_checks'][file_path] = {
                'is_valid': is_valid,
                'details': check_details
            }
            
            if is_valid:
                valid_count += 1
                print(f"  ✓ 通过")
            else:
                invalid_count += 1
                print(f"  ✗ 发现问题")
        
        print("\n" + "="*80)
        print("系统耦合协调度分析")
        print("="*80)
        
        if code_files:
            coupling_result = self.analyze_coupling(code_files)
            self.check_results['coupling_analysis'] = coupling_result
            print(f"\n耦合度评分: {coupling_result['coupling_score']:.2f}")
            print(f"评价: {coupling_result['recommendation']}")
        
        self.check_results['summary'] = {
            'total_files': len(files_to_check),
            'valid_files': valid_count,
            'invalid_files': invalid_count,
            'issues_count': len(self.check_results['issues']),
            'check_time': datetime.now().isoformat()
        }
        
        severity_counts = {}
        for issue in self.check_results['issues']:
            sev = issue['severity']
            severity_counts[sev] = severity_counts.get(sev, 0) + 1
        
        self.check_results['summary']['issues_by_severity'] = severity_counts
        
        report_path = '全面文件检查报告.json'
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(self.check_results, f, ensure_ascii=False, indent=2)
        
        print(f"\n检查报告已保存到: {report_path}")
        
        self.generate_markdown_report()
        
        return self.check_results
    
    def generate_markdown_report(self):
        md_report = '# 全面文件完整性与功能性检查报告\n\n'
        md_report += f'**检查时间**: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n\n'
        
        summary = self.check_results['summary']
        md_report += '## 检查摘要\n\n'
        md_report += '| 指标 | 数值 |\n'
        md_report += '|------|------|\n'
        md_report += f'| 检查文件总数 | {summary["total_files"]} |\n'
        md_report += f'| 通过文件数 | {summary["valid_files"]} |\n'
        md_report += f'| 有问题文件数 | {summary["invalid_files"]} |\n'
        md_report += f'| 发现问题总数 | {summary["issues_count"]} |\n'
        
        if 'issues_by_severity' in summary:
            md_report += '\n### 问题按严重程度分布\n\n'
            for sev, count in summary['issues_by_severity'].items():
                md_report += f'- **{sev}**: {count}个\n'
        
        if self.check_results['issues']:
            md_report += '\n## 发现的问题\n\n'
            
            for sev in ['critical', 'high', 'medium', 'low']:
                sev_issues = [i for i in self.check_results['issues'] if i['severity'] == sev]
                if sev_issues:
                    md_report += f'\n### {sev.upper()} 严重程度\n\n'
                    for issue in sev_issues:
                        md_report += f'**文件**: {issue["file_path"]}\n\n'
                        md_report += f'**类别**: {issue["category"]}\n\n'
                        md_report += f'**描述**: {issue["description"]}\n\n'
                        if issue['details']:
                            md_report += f'**详情**: {issue["details"]}\n\n'
                        md_report += '---\n\n'
        
        if 'coupling_analysis' in self.check_results and self.check_results['coupling_analysis']:
            coupling = self.check_results['coupling_analysis']
            md_report += '\n## 系统耦合协调度分析\n\n'
            md_report += f'- **耦合度评分**: {coupling["coupling_score"]:.2f}\n'
            md_report += f'- **评价**: {coupling["recommendation"]}\n'
        
        with open('全面文件检查报告.md', 'w', encoding='utf-8') as f:
            f.write(md_report)
        
        print("Markdown报告已保存到: 全面文件检查报告.md")

def main():
    checker = ComprehensiveFileChecker()
    checker.run_check()

if __name__ == '__main__':
    main()
