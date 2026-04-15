import os
import json
import pandas as pd
import numpy as np
import ast
import re
import subprocess
from datetime import datetime
from typing import List, Dict, Any, Tuple

class ComprehensiveAuditor:
    def __init__(self):
        self.audit_results = {
            'summary': {},
            'integrity_checks': {},
            'functional_checks': {},
            'issues': [],
            'passing_items': [],
            'recommendations': []
        }
        self.severity_levels = {
            'critical': 3,
            'high': 2,
            'medium': 1,
            'low': 0
        }
    
    def log_issue(self, severity: str, category: str, file_path: str, 
                  description: str, location: str = None, 
                  impact_scope: str = None, details: str = None):
        issue = {
            'timestamp': datetime.now().isoformat(),
            'severity': severity,
            'category': category,
            'file_path': file_path,
            'location': location,
            'description': description,
            'impact_scope': impact_scope,
            'details': details
        }
        self.audit_results['issues'].append(issue)
    
    def log_passing_item(self, category: str, file_path: str, description: str):
        self.audit_results['passing_items'].append({
            'category': category,
            'file_path': file_path,
            'description': description
        })
    
    def add_recommendation(self, priority: str, description: str, actionable_steps: List[str]):
        self.audit_results['recommendations'].append({
            'priority': priority,
            'description': description,
            'actionable_steps': actionable_steps
        })
    
    def check_json_format(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'has_required_fields': True,
            'data_quality': 'good'
        }
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if not isinstance(data, (dict, list)):
                result['is_valid'] = False
                self.log_issue('medium', 'JSON格式', file_path, 
                              '根节点应该是对象或数组', 
                              impact_scope='文件结构')
            
            if isinstance(data, dict):
                required_fields = ['report_title', 'created_at', 'results']
                missing_fields = [f for f in required_fields if f not in data]
                if missing_fields and '报告' in os.path.basename(file_path):
                    result['has_required_fields'] = False
                    self.log_issue('medium', '内容完整性', file_path,
                                  f'缺少必要字段: {", ".join(missing_fields)}',
                                  impact_scope='报告内容')
            
            self.log_passing_item('JSON格式', file_path, 'JSON格式验证通过')
            
        except json.JSONDecodeError as e:
            result['is_valid'] = False
            self.log_issue('critical', 'JSON格式', file_path,
                          f'JSON解析错误: {str(e)}',
                          location=f'第{e.lineno}行',
                          impact_scope='整个文件')
        except Exception as e:
            result['is_valid'] = False
            self.log_issue('high', '文件损坏', file_path,
                          f'文件读取失败: {str(e)}',
                          impact_scope='整个文件')
        
        return result['is_valid'], result
    
    def check_python_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'has_syntax': True,
            'has_main': True,
            'imports_complete': True
        }
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            try:
                tree = ast.parse(content)
                
                imports = []
                functions = []
                classes = []
                has_main = False
                
                for node in ast.walk(tree):
                    if isinstance(node, ast.Import):
                        for name in node.names:
                            imports.append(name.name)
                    elif isinstance(node, ast.ImportFrom):
                        imports.append(node.module or '')
                    elif isinstance(node, ast.FunctionDef):
                        functions.append(node.name)
                    elif isinstance(node, ast.ClassDef):
                        classes.append(node.name)
                    elif isinstance(node, ast.If):
                        if (isinstance(node.test, ast.Compare) and
                            isinstance(node.test.left, ast.Name) and
                            node.test.left.id == '__name__'):
                            has_main = True
                
                if not functions and not classes:
                    result['is_valid'] = False
                    self.log_issue('low', '代码结构', file_path,
                                  '没有定义函数或类',
                                  impact_scope='代码功能')
                
                if not has_main and len(functions) > 0:
                    result['has_main'] = False
                    self.log_issue('low', '代码结构', file_path,
                                  '缺少__main__入口点',
                                  impact_scope='独立运行')
                
                self.log_passing_item('Python语法', file_path, 'Python语法检查通过')
                
            except SyntaxError as e:
                result['is_valid'] = False
                result['has_syntax'] = False
                self.log_issue('critical', 'Python语法', file_path,
                              f'语法错误: {e.msg}',
                              location=f'第{e.lineno}行, 列{e.offset}',
                              impact_scope='整个文件')
            
        except Exception as e:
            result['is_valid'] = False
            self.log_issue('high', '文件损坏', file_path,
                          f'文件读取失败: {str(e)}',
                          impact_scope='整个文件')
        
        return result['is_valid'], result
    
    def check_excel_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'has_data': True,
            'no_corruption': True,
            'data_quality': 'good'
        }
        
        try:
            xl = pd.ExcelFile(file_path)
            
            df = pd.read_excel(file_path)
            
            if len(df) == 0:
                result['has_data'] = False
                self.log_issue('medium', '内容完整性', file_path,
                              'Excel文件为空',
                              impact_scope='数据内容')
            
            if len(df.columns) == 0:
                result['has_data'] = False
                self.log_issue('medium', '内容完整性', file_path,
                              '没有列数据',
                              impact_scope='数据结构')
            
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            for col in numeric_cols:
                if df[col].isnull().sum() > len(df) * 0.5:
                    result['data_quality'] = 'poor'
                    self.log_issue('medium', '数据质量', file_path,
                                  f'列"{col}"缺失值超过50%',
                                  impact_scope='数据可用性')
            
            self.log_passing_item('Excel格式', file_path, 'Excel文件格式验证通过')
            
        except Exception as e:
            result['is_valid'] = False
            result['no_corruption'] = False
            self.log_issue('high', '文件损坏', file_path,
                          f'Excel读取失败: {str(e)}',
                          impact_scope='整个文件')
        
        return result['is_valid'], result
    
    def check_csv_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'has_data': True,
            'encoding_correct': True
        }
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            
            if len(df) == 0:
                result['has_data'] = False
                self.log_issue('medium', '内容完整性', file_path,
                              'CSV文件为空',
                              impact_scope='数据内容')
            
            self.log_passing_item('CSV格式', file_path, 'CSV文件格式验证通过')
            
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, encoding='gbk')
                result['encoding_correct'] = False
                self.log_issue('low', '文件格式', file_path,
                              '使用GBK编码，建议使用UTF-8-BOM',
                              impact_scope='跨平台兼容性')
            except:
                result['is_valid'] = False
                self.log_issue('high', '文件损坏', file_path,
                              'CSV编码识别失败',
                              impact_scope='整个文件')
        except Exception as e:
            result['is_valid'] = False
            self.log_issue('high', '文件损坏', file_path,
                          f'CSV读取失败: {str(e)}',
                          impact_scope='整个文件')
        
        return result['is_valid'], result
    
    def check_markdown_file(self, file_path: str) -> Tuple[bool, Dict]:
        result = {
            'is_valid': True,
            'has_structure': True,
            'content_complete': True
        }
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            if len(content.strip()) == 0:
                result['is_valid'] = False
                result['content_complete'] = False
                self.log_issue('medium', '内容完整性', file_path,
                              'Markdown文件为空',
                              impact_scope='文档内容')
            
            if not content.strip().startswith('#'):
                result['has_structure'] = False
                self.log_issue('low', '文档结构', file_path,
                              '没有一级标题',
                              impact_scope='文档可读性')
            
            sections = len(re.findall(r'^#{1,6}\s', content, re.MULTILINE))
            if sections == 0 and len(content) > 100:
                result['has_structure'] = False
                self.log_issue('low', '文档结构', file_path,
                              '没有章节标题',
                              impact_scope='文档可读性')
            
            self.log_passing_item('Markdown格式', file_path, 'Markdown文件格式验证通过')
            
        except Exception as e:
            result['is_valid'] = False
            self.log_issue('high', '文件损坏', file_path,
                          f'Markdown读取失败: {str(e)}',
                          impact_scope='整个文件')
        
        return result['is_valid'], result
    
    def test_functional_independent(self, file_path: str) -> bool:
        if not file_path.endswith('.py'):
            return True
        
        try:
            result = subprocess.run(
                ['python', '-m', 'py_compile', file_path],
                capture_output=True,
                timeout=30
            )
            if result.returncode == 0:
                self.log_passing_item('功能性-独立运行', file_path, 
                                    'Python语法编译检查通过')
                return True
            else:
                self.log_issue('high', '功能性-独立运行', file_path,
                              'Python编译失败',
                              details=result.stderr.decode('utf-8', errors='ignore'))
                return False
        except Exception as e:
            self.log_issue('medium', '功能性-独立运行', file_path,
                          f'功能测试异常: {str(e)}')
            return False
    
    def run_integrity_checks(self):
        print("="*80)
        print("执行完整性检查")
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
        
        valid_count = 0
        invalid_count = 0
        
        for file_path in files_to_check:
            ext = os.path.splitext(file_path)[1].lower()
            is_valid = True
            
            print(f"检查: {file_path}")
            
            if ext == '.json':
                is_valid, _ = self.check_json_format(file_path)
            elif ext == '.py':
                is_valid, _ = self.check_python_file(file_path)
            elif ext == '.xlsx':
                is_valid, _ = self.check_excel_file(file_path)
            elif ext == '.csv':
                is_valid, _ = self.check_csv_file(file_path)
            elif ext == '.md':
                is_valid, _ = self.check_markdown_file(file_path)
            
            if is_valid:
                valid_count += 1
            else:
                invalid_count += 1
        
        self.audit_results['integrity_checks'] = {
            'total_files': len(files_to_check),
            'valid_files': valid_count,
            'invalid_files': invalid_count
        }
        
        return self.audit_results['integrity_checks']
    
    def run_functional_checks(self):
        print("\n" + "="*80)
        print("执行功能性检查")
        print("="*80)
        
        code_folder = 'code'
        code_files = []
        
        if os.path.exists(code_folder):
            for root, dirs, files in os.walk(code_folder):
                for file in files:
                    if file.endswith('.py'):
                        code_files.append(os.path.join(root, file))
        
        print(f"\n找到 {len(code_files)} 个Python文件进行功能测试\n")
        
        functional_pass = 0
        functional_fail = 0
        
        for file_path in code_files:
            print(f"功能测试: {file_path}")
            if self.test_functional_independent(file_path):
                functional_pass += 1
            else:
                functional_fail += 1
        
        self.audit_results['functional_checks'] = {
            'total_code_files': len(code_files),
            'functional_pass': functional_pass,
            'functional_fail': functional_fail
        }
        
        return self.audit_results['functional_checks']
    
    def generate_recommendations(self):
        print("\n" + "="*80)
        print("生成改进建议")
        print("="*80)
        
        severity_counts = {}
        for issue in self.audit_results['issues']:
            sev = issue['severity']
            severity_counts[sev] = severity_counts.get(sev, 0) + 1
        
        if severity_counts.get('critical', 0) > 0:
            self.add_recommendation(
                'high',
                '紧急修复严重问题',
                ['优先处理所有critical级别的问题', '验证修复后的文件功能', '进行回归测试']
            )
        
        if severity_counts.get('high', 0) > 0:
            self.add_recommendation(
                'medium',
                '解决高优先级问题',
                ['处理所有high级别的问题', '检查数据文件的完整性', '确保代码可正常运行']
            )
        
        if not self.audit_results['issues']:
            self.add_recommendation(
                'low',
                '持续改进',
                ['建立定期检查机制', '考虑添加单元测试', '完善文档注释']
            )
        
        self.add_recommendation(
            'medium',
            '建立版本控制',
            ['初始化Git仓库', '提交当前版本', '建立分支策略']
        )
    
    def generate_report(self):
        print("\n" + "="*80)
        print("生成审计报告")
        print("="*80)
        
        summary = {
            'audit_date': datetime.now().isoformat(),
            'integrity': self.audit_results.get('integrity_checks', {}),
            'functional': self.audit_results.get('functional_checks', {}),
            'total_issues': len(self.audit_results['issues']),
            'total_passing': len(self.audit_results['passing_items']),
            'issues_by_severity': {}
        }
        
        for issue in self.audit_results['issues']:
            sev = issue['severity']
            summary['issues_by_severity'][sev] = summary['issues_by_severity'].get(sev, 0) + 1
        
        self.audit_results['summary'] = summary
        
        report_path = '全面审计报告.json'
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(self.audit_results, f, ensure_ascii=False, indent=2)
        
        print(f"JSON报告已保存到: {report_path}")
        
        self.generate_markdown_report()
    
    def generate_markdown_report(self):
        md_report = '# 全面完整性与功能性审计报告\n\n'
        md_report += f'**审计时间**: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n\n'
        
        summary = self.audit_results['summary']
        
        md_report += '## 执行摘要\n\n'
        md_report += '### 检查概览\n\n'
        md_report += '| 指标 | 数值 |\n'
        md_report += '|------|------|\n'
        md_report += f'| 检查文件总数 | {summary["integrity"].get("total_files", 0)} |\n'
        md_report += f'| 通过文件数 | {summary["integrity"].get("valid_files", 0)} |\n'
        md_report += f'| 有问题文件数 | {summary["integrity"].get("invalid_files", 0)} |\n'
        md_report += f'| 发现问题总数 | {summary["total_issues"]} |\n'
        md_report += f'| 通过项总数 | {summary["total_passing"]} |\n'
        
        if summary['issues_by_severity']:
            md_report += '\n### 问题按严重程度分布\n\n'
            for sev, count in summary['issues_by_severity'].items():
                md_report += f'- **{sev.upper()}**: {count}个\n'
        
        md_report += '\n### 功能性检查\n\n'
        md_report += '| 指标 | 数值 |\n'
        md_report += '|------|------|\n'
        md_report += f'| 代码文件总数 | {summary["functional"].get("total_code_files", 0)} |\n'
        md_report += f'| 功能通过数 | {summary["functional"].get("functional_pass", 0)} |\n'
        md_report += f'| 功能失败数 | {summary["functional"].get("functional_fail", 0)} |\n'
        
        if self.audit_results['passing_items']:
            md_report += '\n## 通过项清单\n\n'
            for item in self.audit_results['passing_items'][:50]:
                md_report += f'- **{item["category"]}**: {item["file_path"]} - {item["description"]}\n'
            if len(self.audit_results['passing_items']) > 50:
                md_report += f'- ... 还有 {len(self.audit_results["passing_items"]) - 50} 项通过\n'
        
        if self.audit_results['issues']:
            md_report += '\n## 问题项详情\n\n'
            
            for sev in ['critical', 'high', 'medium', 'low']:
                sev_issues = [i for i in self.audit_results['issues'] if i['severity'] == sev]
                if sev_issues:
                    md_report += f'\n### {sev.upper()} 严重程度\n\n'
                    for issue in sev_issues:
                        md_report += f'**文件**: {issue["file_path"]}\n\n'
                        md_report += f'**类别**: {issue["category"]}\n\n'
                        md_report += f'**描述**: {issue["description"]}\n\n'
                        if issue.get('location'):
                            md_report += f'**位置**: {issue["location"]}\n\n'
                        if issue.get('impact_scope'):
                            md_report += f'**影响范围**: {issue["impact_scope"]}\n\n'
                        if issue.get('details'):
                            md_report += f'**详情**: {issue["details"]}\n\n'
                        md_report += '---\n\n'
        
        if self.audit_results['recommendations']:
            md_report += '\n## 改进建议\n\n'
            for rec in self.audit_results['recommendations']:
                md_report += f'### 优先级: {rec["priority"].upper()}\n\n'
                md_report += f'**{rec["description"]}**\n\n'
                md_report += '**可操作步骤**:\n'
                for step in rec['actionable_steps']:
                    md_report += f'1. {step}\n'
                md_report += '\n'
        
        with open('全面审计报告.md', 'w', encoding='utf-8') as f:
            f.write(md_report)
        
        print("Markdown报告已保存到: 全面审计报告.md")
    
    def run_audit(self):
        print("="*100)
        print("全面完整性与功能性审计工具")
        print("="*100)
        
        self.run_integrity_checks()
        self.run_functional_checks()
        self.generate_recommendations()
        self.generate_report()
        
        print("\n" + "="*100)
        print("审计完成！")
        print("="*100)
        
        return self.audit_results

def main():
    auditor = ComprehensiveAuditor()
    auditor.run_audit()

if __name__ == '__main__':
    main()
