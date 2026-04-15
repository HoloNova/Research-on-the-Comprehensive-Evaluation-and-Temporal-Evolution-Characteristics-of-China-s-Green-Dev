"""
数据处理与分析标准化流程 - 精简核心版

包含核心功能：数据预处理、熵权法、TOPSIS、耦合协调度、可视化
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False
sns.set_style('whitegrid')


class DataProcessor:
    """数据预处理模块"""
    
    @staticmethod
    def load_data(file_path: str) -> pd.DataFrame:
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path, encoding='utf-8-sig')
        elif file_path.endswith('.xlsx'):
            return pd.read_excel(file_path)
        raise ValueError("不支持的文件格式")
    
    @staticmethod
    def clean_data(df: pd.DataFrame) -> pd.DataFrame:
        df_clean = df.copy()
        numeric_cols = df_clean.select_dtypes(include=[np.number]).columns
        
        for col in numeric_cols:
            if df_clean[col].isnull().sum() > 0:
                if '年份' in df_clean.columns and '省份' in df_clean.columns:
                    df_clean = df_clean.sort_values('年份')
                    df_clean[col] = df_clean.groupby('省份')[col].transform(
                        lambda x: x.interpolate(method='linear')
                    )
                if df_clean[col].isnull().sum() > 0:
                    df_clean[col] = df_clean[col].fillna(df_clean[col].mean())
        
        return df_clean
    
    @staticmethod
    def normalize(df: pd.DataFrame, positive_cols: list, negative_cols: list) -> pd.DataFrame:
        df_norm = df.copy()
        all_cols = positive_cols + negative_cols
        
        for col in all_cols:
            if col in df.columns:
                min_val = df[col].min()
                max_val = df[col].max()
                if max_val == min_val:
                    df_norm[col] = 0.5
                else:
                    if col in positive_cols:
                        df_norm[col] = (df[col] - min_val) / (max_val - min_val)
                    else:
                        df_norm[col] = (max_val - df[col]) / (max_val - min_val)
        
        return df_norm


class EntropyWeight:
    """熵权法模块"""
    
    @staticmethod
    def calculate(matrix: np.ndarray) -> np.ndarray:
        m, n = matrix.shape
        P = matrix / (matrix.sum(axis=0) + 1e-10)
        k = 1 / np.log(m) if m > 1 else 1
        E = -k * (P * np.log(P + 1e-10)).sum(axis=0)
        g = 1 - E
        W = g / (g.sum() + 1e-10)
        return W


class TOPSIS:
    """TOPSIS模块"""
    
    @staticmethod
    def calculate(matrix: np.ndarray, weights: np.ndarray, positive_mask: np.ndarray) -> np.ndarray:
        V = matrix * weights
        V_plus = np.max(V * positive_mask.reshape(1, -1), axis=0)
        V_minus = np.min(V * positive_mask.reshape(1, -1), axis=0)
        D_plus = np.sqrt(((V - V_plus) ** 2).sum(axis=1))
        D_minus = np.sqrt(((V - V_minus) ** 2).sum(axis=1))
        S = D_minus / (D_plus + D_minus + 1e-10)
        return S


class CouplingCoordination:
    """耦合协调度模块"""
    
    @staticmethod
    def calculate(u1: np.ndarray, u2: np.ndarray) -> dict:
        C = 2 * np.sqrt((u1 * u2) / (u1 + u2 + 1e-10))
        T = 0.5 * u1 + 0.5 * u2
        D = np.sqrt(C * T)
        
        def get_level(d):
            if d >= 0.8: return '优质协调'
            elif d >= 0.6: return '中度协调'
            elif d >= 0.4: return '勉强协调'
            elif d >= 0.2: return '中度失调'
            else: return '极度失调'
        
        levels = [get_level(d) for d in D]
        return {'coupling': C, 'development': T, 'coordination': D, 'level': levels}


class Visualizer:
    """可视化模块"""
    
    @staticmethod
    def plot_results(df: pd.DataFrame, save_dir: str = 'visualizations'):
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        if '年份' in df.columns:
            df_yearly = df.groupby('年份').mean().reset_index()
            
            fig, axes = plt.subplots(2, 1, figsize=(12, 8))
            
            axes[0].plot(df_yearly['年份'], df_yearly['综合得分'], 
                       marker='o', linewidth=2, label='综合得分')
            axes[0].set_title('综合得分时间趋势', fontsize=14, fontweight='bold')
            axes[0].set_xlabel('年份')
            axes[0].set_ylabel('得分')
            axes[0].legend()
            axes[0].grid(True, alpha=0.3)
            
            axes[1].plot(df_yearly['年份'], df_yearly['耦合协调度'], 
                       marker='s', linewidth=2, color='orange', label='耦合协调度')
            axes[1].set_title('耦合协调度时间趋势', fontsize=14, fontweight='bold')
            axes[1].set_xlabel('年份')
            axes[1].set_ylabel('协调度')
            axes[1].legend()
            axes[1].grid(True, alpha=0.3)
            
            plt.tight_layout()
            plt.savefig(os.path.join(save_dir, '时间趋势图.png'), dpi=300, bbox_inches='tight')
            plt.close()
        
        if '年份' in df.columns and '省份' in df.columns:
            pivot_df = df.pivot(index='省份', columns='年份', values='综合得分')
            fig, ax = plt.subplots(figsize=(14, 10))
            sns.heatmap(pivot_df, annot=True, fmt='.4f', cmap='RdYlGn', center=0.5, ax=ax)
            ax.set_title('各省份综合得分热力图', fontsize=14, fontweight='bold')
            plt.tight_layout()
            plt.savefig(os.path.join(save_dir, '综合得分热力图.png'), dpi=300, bbox_inches='tight')
            plt.close()


class StandardAnalysis:
    """标准化流程主类"""
    
    def __init__(self):
        self.processor = DataProcessor()
        self.entropy = EntropyWeight()
        self.topsis = TOPSIS()
        self.coupling = CouplingCoordination()
        self.visualizer = Visualizer()
    
    def run(self, input_file: str, output_dir: str = 'results', 
            positive_cols: list = None, negative_cols: list = None):
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        print("="*80)
        print("步骤1：数据预处理")
        print("="*80)
        df = self.processor.load_data(input_file)
        print(f"数据加载完成: {df.shape}")
        
        df_clean = self.processor.clean_data(df)
        print("数据清洗完成")
        
        if positive_cols is None:
            positive_cols = [col for col in df.columns if col not in ['年份', '省份']]
        if negative_cols is None:
            negative_cols = []
        
        df_norm = self.processor.normalize(df_clean, positive_cols, negative_cols)
        print("数据标准化完成")
        
        print("\n" + "="*80)
        print("步骤2：熵权法计算权重")
        print("="*80)
        all_cols = positive_cols + negative_cols
        matrix = df_norm[all_cols].values
        weights = self.entropy.calculate(matrix)
        print(f"权重计算完成: {np.round(weights, 4)}")
        
        print("\n" + "="*80)
        print("步骤3：TOPSIS计算综合得分")
        print("="*80)
        positive_mask = np.array([1 if col in positive_cols else -1 for col in all_cols])
        scores = self.topsis.calculate(matrix, weights, positive_mask)
        print("综合得分计算完成")
        
        print("\n" + "="*80)
        print("步骤4：耦合协调度计算")
        print("="*80)
        eco_cols = positive_cols[:3] if len(positive_cols) >=3 else positive_cols
        env_cols = positive_cols[3:] + negative_cols
        
        eco_indices = [all_cols.index(col) for col in eco_cols if col in all_cols]
        env_indices = [all_cols.index(col) for col in env_cols if col in all_cols]
        
        if eco_indices and env_indices:
            u1 = self.topsis.calculate(matrix[:, eco_indices], weights[eco_indices]/weights[eco_indices].sum(),
                                       np.array([1]*len(eco_indices)))
            u2 = self.topsis.calculate(matrix[:, env_indices], weights[env_indices]/weights[env_indices].sum(),
                                       np.array([1 if col in positive_cols else -1 for col in env_cols if col in all_cols]))
            coupling_result = self.coupling.calculate(u1, u2)
        else:
            u1 = scores * 0.5
            u2 = scores * 0.5
            coupling_result = self.coupling.calculate(u1, u2)
        
        print("耦合协调度计算完成")
        
        df_result = df_clean.copy()
        df_result['综合得分'] = scores
        df_result['经济子系统得分'] = u1
        df_result['环境子系统得分'] = u2
        df_result['耦合度'] = coupling_result['coupling']
        df_result['综合发展指数'] = coupling_result['development']
        df_result['耦合协调度'] = coupling_result['coordination']
        df_result['协调等级'] = coupling_result['level']
        
        result_file = os.path.join(output_dir, '分析结果.xlsx')
        df_result.to_excel(result_file, index=False)
        print(f"结果已保存: {result_file}")
        
        print("\n" + "="*80)
        print("步骤5：可视化")
        print("="*80)
        vis_dir = os.path.join(output_dir, 'visualizations')
        self.visualizer.plot_results(df_result, vis_dir)
        print(f"可视化图表已保存: {vis_dir}")
        
        self.generate_report(df_result, weights, all_cols, output_dir)
        
        return df_result
    
    def generate_report(self, df_result: pd.DataFrame, weights: np.ndarray, 
                       cols: list, output_dir: str):
        report_file = os.path.join(output_dir, '分析报告.md')
        
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("# 数据分析报告\n\n")
            f.write(f"**生成时间**: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            f.write("## 数据概览\n\n")
            f.write(f"- 样本数量: {len(df_result)}\n")
            f.write(f"- 指标数量: {len(cols)}\n")
            if '年份' in df_result.columns:
                f.write(f"- 时间范围: {df_result['年份'].min()} - {df_result['年份'].max()}\n")
            
            f.write("\n## 指标权重\n\n")
            f.write("| 指标 | 权重 |\n")
            f.write("|------|------|\n")
            for col, w in zip(cols, weights):
                f.write(f"| {col} | {w:.4f} |\n")
            
            f.write("\n## 关键指标统计\n\n")
            key_metrics = ['综合得分', '耦合协调度']
            for metric in key_metrics:
                if metric in df_result.columns:
                    f.write(f"### {metric}\n\n")
                    f.write(f"- 平均值: {df_result[metric].mean():.4f}\n")
                    f.write(f"- 最大值: {df_result[metric].max():.4f}\n")
                    f.write(f"- 最小值: {df_result[metric].min():.4f}\n")
                    f.write(f"- 标准差: {df_result[metric].std():.4f}\n\n")
            
            if '协调等级' in df_result.columns:
                f.write("## 协调等级分布\n\n")
                level_counts = df_result['协调等级'].value_counts()
                for level, count in level_counts.items():
                    f.write(f"- {level}: {count}个\n")
        
        print(f"分析报告已保存: {report_file}")


def main():
    print("="*80)
    print("数据处理与分析标准化流程 - 精简核心版")
    print("="*80)
    print("\n使用示例:")
    print("from code.standard_analysis import StandardAnalysis")
    print("analysis = StandardAnalysis()")
    print("results = analysis.run('data.xlsx', 'results')")
    print("\n流程已准备就绪！")


if __name__ == '__main__':
    main()
