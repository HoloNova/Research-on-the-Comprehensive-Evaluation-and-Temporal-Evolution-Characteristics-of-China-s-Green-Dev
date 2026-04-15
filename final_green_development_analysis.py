"""
我国绿色发展水平综合评价与时间演化特征研究
完整分析脚本 - 使用已有分析结果
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False
sns.set_style('whitegrid')


def main():
    print("="*80)
    print("我国绿色发展水平综合评价与时间演化特征研究")
    print("="*80)
    
    # 创建输出目录
    output_dir = '绿色发展研究结果'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 步骤1：读取已有分析结果
    print("\n" + "="*80)
    print("步骤1：读取已有分析结果")
    print("="*80)
    
    results = pd.read_excel('analysis_results/分析结果.xlsx')
    print(f"数据形状: {results.shape}")
    print("\n数据列名:")
    print(results.columns.tolist())
    
    # 步骤2：生成详细分析报告
    print("\n" + "="*80)
    print("步骤2：生成详细分析报告")
    print("="*80)
    
    generate_detailed_report(results, output_dir)
    
    # 步骤3：创建额外的可视化图表
    print("\n" + "="*80)
    print("步骤3：创建额外的可视化图表")
    print("="*80)
    
    create_additional_visualizations(results, output_dir)
    
    print("\n" + "="*80)
    print("分析完成！")
    print("="*80)


def generate_detailed_report(results, output_dir):
    """生成详细的分析报告"""
    
    report_file = os.path.join(output_dir, '我国绿色发展水平综合评价研究报告.md')
    
    # 定义指标
    positive_cols = ['人均GDP', '第三产业占比', 'R&D占比', '污水处理率', 
                     '垃圾处理率', '森林覆盖率', '优良天数比例', '环保投资占比']
    negative_cols = ['单位GDP能耗', '单位GDP水耗']
    all_cols = positive_cols + negative_cols
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("# 我国绿色发展水平综合评价与时间演化特征研究\n\n")
        f.write(f"**生成时间**: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # 1. 研究背景与方法
        f.write("## 1. 研究背景与方法\n\n")
        f.write("### 1.1 研究背景\n\n")
        f.write("绿色发展是实现经济高质量发展的重要途径，本研究基于熵权法-TOPSIS-耦合协调度模型，")
        f.write("对我国31个省份2018-2022年的绿色发展水平进行综合评价。\n\n")
        
        f.write("### 1.2 研究方法\n\n")
        f.write("#### 1.2.1 数据标准化\n\n")
        f.write("- **正向指标**: $X'_{ij} = \\frac{X_{ij} - \\min(X_j)}{\\max(X_j) - \\min(X_j)}$\n")
        f.write("- **负向指标**: $X'_{ij} = \\frac{\\max(X_j) - X_{ij}}{\\max(X_j) - \\min(X_j)}$\n\n")
        
        f.write("#### 1.2.2 熵权法\n\n")
        f.write("1. 比重矩阵: $P_{ij} = \\frac{X_{ij}}{\\sum_{i=1}^{m} X_{ij}}$\n")
        f.write("2. 信息熵: $E_j = -k \\sum_{i=1}^{m} P_{ij} \\ln(P_{ij})$\n")
        f.write("3. 差异系数: $G_j = 1 - E_j$\n")
        f.write("4. 权重: $W_j = \\frac{G_j}{\\sum_{j=1}^{n} G_j}$\n\n")
        
        f.write("#### 1.2.3 TOPSIS方法\n\n")
        f.write("1. 加权标准化矩阵: $V_{ij} = W_j \\times X_{ij}$\n")
        f.write("2. 正/负理想解: 确定各指标最优值和最差值\n")
        f.write("3. 欧氏距离: $D_i^+ = \\sqrt{\\sum_{j=1}^{n} (V_{ij} - V_j^+)^2}$\n")
        f.write("4. 综合得分: $S_i = \\frac{D_i^-}{D_i^+ + D_i^-}$\n\n")
        
        f.write("#### 1.2.4 耦合协调度\n\n")
        f.write("1. 耦合度: $C = 2 \\times \\sqrt{\\frac{u_1 \\times u_2}{u_1 + u_2}}$\n")
        f.write("2. 综合发展指数: $T = 0.5u_1 + 0.5u_2$\n")
        f.write("3. 耦合协调度: $D = \\sqrt{C \\times T}$\n\n")
        
        # 2. 数据概况
        f.write("## 2. 数据概况\n\n")
        f.write(f"- **样本数量**: {len(results)}个（31省份×5年）\n")
        f.write(f"- **时间范围**: {results['年份'].min()}-{results['年份'].max()}年\n")
        f.write(f"- **评价指标**: 10个\n\n")
        
        f.write("### 2.1 评价指标体系\n\n")
        f.write("| 子系统 | 指标名称 | 指标类型 |\n")
        f.write("|--------|---------|---------|\n")
        f.write("| **经济子系统** | 人均GDP | 正向 |\n")
        f.write("| **经济子系统** | 第三产业占比 | 正向 |\n")
        f.write("| **经济子系统** | R&D占比 | 正向 |\n")
        f.write("| **环境子系统** | 单位GDP能耗 | 负向 |\n")
        f.write("| **环境子系统** | 单位GDP水耗 | 负向 |\n")
        f.write("| **环境子系统** | 污水处理率 | 正向 |\n")
        f.write("| **环境子系统** | 垃圾处理率 | 正向 |\n")
        f.write("| **环境子系统** | 森林覆盖率 | 正向 |\n")
        f.write("| **环境子系统** | 优良天数比例 | 正向 |\n")
        f.write("| **环境子系统** | 环保投资占比 | 正向 |\n\n")
        
        # 3. 熵权法权重结果
        f.write("## 3. 熵权法权重结果\n\n")
        
        # 重新计算权重用于展示
        from standard_analysis import DataProcessor, EntropyWeight
        
        df_raw = pd.read_excel('sample_data.xlsx')
        
        processor = DataProcessor()
        df_norm = processor.normalize(df_raw, positive_cols, negative_cols)
        matrix = df_norm[all_cols].values
        
        entropy = EntropyWeight()
        weights = entropy.calculate(matrix)
        
        f.write("### 3.1 指标权重表\n\n")
        f.write("| 指标名称 | 信息熵 | 差异系数 | 权重 | 权重排序 |\n")
        f.write("|---------|--------|---------|------|---------|\n")
        
        # 计算信息熵等用于展示
        m, n = matrix.shape
        P = matrix / (matrix.sum(axis=0) + 1e-10)
        k = 1 / np.log(m) if m > 1 else 1
        E = -k * (P * np.log(P + 1e-10)).sum(axis=0)
        g = 1 - E
        
        weight_df = pd.DataFrame({
            '指标': all_cols,
            '信息熵': E,
            '差异系数': g,
            '权重': weights
        }).sort_values('权重', ascending=False).reset_index(drop=True)
        
        for idx, row in weight_df.iterrows():
            f.write(f"| {row['指标']} | {row['信息熵']:.4f} | {row['差异系数']:.4f} | {row['权重']:.4f} | {int(idx+1)} |\n")
        
        f.write("\n### 3.2 权重分析\n\n")
        f.write(f"- **权重最大指标**: {weight_df.iloc[0]['指标']} ({weight_df.iloc[0]['权重']:.4f})\n")
        f.write(f"- **权重最小指标**: {weight_df.iloc[-1]['指标']} ({weight_df.iloc[-1]['权重']:.4f})\n")
        eco_weight = weight_df[weight_df['指标'].isin(positive_cols[:3])]['权重'].sum()
        env_weight = weight_df[~weight_df['指标'].isin(positive_cols[:3])]['权重'].sum()
        f.write(f"- **经济子系统权重**: {eco_weight:.4f}\n")
        f.write(f"- **环境子系统权重**: {env_weight:.4f}\n\n")
        
        # 4. TOPSIS综合得分
        f.write("## 4. TOPSIS综合得分\n\n")
        
        f.write("### 4.1 综合得分统计\n\n")
        f.write("| 统计量 | 综合得分 | 经济子系统得分 | 环境子系统得分 |\n")
        f.write("|--------|---------|---------------|---------------|\n")
        f.write(f"| 平均值 | {results['综合得分'].mean():.4f} | {results['经济子系统得分'].mean():.4f} | {results['环境子系统得分'].mean():.4f} |\n")
        f.write(f"| 最大值 | {results['综合得分'].max():.4f} | {results['经济子系统得分'].max():.4f} | {results['环境子系统得分'].max():.4f} |\n")
        f.write(f"| 最小值 | {results['综合得分'].min():.4f} | {results['经济子系统得分'].min():.4f} | {results['环境子系统得分'].min():.4f} |\n")
        f.write(f"| 标准差 | {results['综合得分'].std():.4f} | {results['经济子系统得分'].std():.4f} | {results['环境子系统得分'].std():.4f} |\n\n")
        
        f.write("### 4.2 综合得分排名前10省份\n\n")
        top10 = results.sort_values('综合得分', ascending=False).head(10)
        f.write("| 排名 | 省份 | 年份 | 综合得分 | 经济子系统 | 环境子系统 |\n")
        f.write("|------|------|------|---------|-----------|-----------|\n")
        for idx, (_, row) in enumerate(top10.iterrows(), 1):
            f.write(f"| {idx} | {row['省份']} | {row['年份']} | {row['综合得分']:.4f} | {row['经济子系统得分']:.4f} | {row['环境子系统得分']:.4f} |\n")
        
        # 5. 耦合协调度分析
        f.write("\n## 5. 耦合协调度分析\n\n")
        
        f.write("### 5.1 耦合协调度统计\n\n")
        f.write("| 统计量 | 耦合度 | 综合发展指数 | 耦合协调度 |\n")
        f.write("|--------|--------|-------------|-----------|\n")
        f.write(f"| 平均值 | {results['耦合度'].mean():.4f} | {results['综合发展指数'].mean():.4f} | {results['耦合协调度'].mean():.4f} |\n")
        f.write(f"| 最大值 | {results['耦合度'].max():.4f} | {results['综合发展指数'].max():.4f} | {results['耦合协调度'].max():.4f} |\n")
        f.write(f"| 最小值 | {results['耦合度'].min():.4f} | {results['综合发展指数'].min():.4f} | {results['耦合协调度'].min():.4f} |\n")
        f.write(f"| 标准差 | {results['耦合度'].std():.4f} | {results['综合发展指数'].std():.4f} | {results['耦合协调度'].std():.4f} |\n\n")
        
        f.write("### 5.2 协调等级分布\n\n")
        level_counts = results['协调等级'].value_counts().sort_index()
        for level, count in level_counts.items():
            f.write(f"- **{level}**: {count}个样本 ({count/len(results)*100:.1f}%)\n")
        
        # 6. 时间演化特征
        f.write("\n## 6. 时间演化特征\n\n")
        
        f.write("### 6.1 年度平均趋势\n\n")
        yearly_avg = results.groupby('年份').agg({
            '综合得分': 'mean',
            '经济子系统得分': 'mean',
            '环境子系统得分': 'mean',
            '耦合协调度': 'mean'
        }).round(4)
        
        f.write("| 年份 | 综合得分 | 经济子系统 | 环境子系统 | 耦合协调度 |\n")
        f.write("|------|---------|-----------|-----------|-----------|\n")
        for year, row in yearly_avg.iterrows():
            f.write(f"| {year} | {row['综合得分']:.4f} | {row['经济子系统得分']:.4f} | {row['环境子系统得分']:.4f} | {row['耦合协调度']:.4f} |\n")
        
        f.write("\n### 6.2 趋势分析\n\n")
        for col in ['综合得分', '经济子系统得分', '环境子系统得分', '耦合协调度']:
            trend = yearly_avg[col].iloc[-1] - yearly_avg[col].iloc[0]
            trend_dir = "上升" if trend > 0 else "下降"
            f.write(f"- **{col}**: 整体呈{trend_dir}趋势，变化量为{trend:+.4f}\n")
        
        # 7. 区域差异分析
        f.write("\n## 7. 区域差异分析\n\n")
        
        f.write("### 7.1 各省份平均得分\n\n")
        province_avg = results.groupby('省份').agg({
            '综合得分': 'mean',
            '经济子系统得分': 'mean',
            '环境子系统得分': 'mean',
            '耦合协调度': 'mean'
        }).round(4).sort_values('综合得分', ascending=False)
        
        f.write("| 排名 | 省份 | 综合得分 | 经济子系统 | 环境子系统 | 耦合协调度 |\n")
        f.write("|------|------|---------|-----------|-----------|-----------|\n")
        for idx, (province, row) in enumerate(province_avg.iterrows(), 1):
            f.write(f"| {idx} | {province} | {row['综合得分']:.4f} | {row['经济子系统得分']:.4f} | {row['环境子系统得分']:.4f} | {row['耦合协调度']:.4f} |\n")
        
        # 8. 研究结论
        f.write("\n## 8. 研究结论\n\n")
        
        f.write("### 8.1 主要发现\n\n")
        trend_text = "上升" if yearly_avg['综合得分'].iloc[-1] > yearly_avg['综合得分'].iloc[0] else "下降"
        f.write(f"1. **绿色发展水平**: 我国绿色发展水平整体呈{trend_text}趋势，综合得分平均值为{results['综合得分'].mean():.4f}\n")
        f.write(f"2. **系统协调**: 经济与环境子系统耦合协调度平均值为{results['耦合协调度'].mean():.4f}，主要处于{level_counts.index[0]}水平\n")
        f.write(f"3. **区域差异**: 各省份绿色发展水平存在显著差异，{province_avg.index[0]}地区表现较好\n")
        
        f.write("\n### 8.2 政策建议\n\n")
        f.write("1. **加大环保投入**: 提高环保投资占比，加强环境基础设施建设\n")
        f.write("2. **推动产业升级**: 提高第三产业占比，促进经济结构优化\n")
        f.write("3. **加强技术创新**: 增加R&D投入，推动绿色技术创新\n")
        f.write("4. **促进区域协调**: 针对不同区域制定差异化的绿色发展政策\n")
        f.write("5. **完善考核机制**: 建立绿色发展考核评价体系，引导高质量发展\n")
        
        f.write("\n---\n\n")
        f.write("*报告生成时间: {}*\n".format(pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    print(f"详细分析报告已保存: {report_file}")


def create_additional_visualizations(results, output_dir):
    """创建额外的可视化图表"""
    
    vis_dir = os.path.join(output_dir, '可视化图表')
    if not os.path.exists(vis_dir):
        os.makedirs(vis_dir)
    
    # 图表1: 时间趋势图
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    
    numeric_cols = results.select_dtypes(include=[np.number]).columns
    yearly_avg = results.groupby('年份')[numeric_cols].mean()
    
    axes[0, 0].plot(yearly_avg.index, yearly_avg['综合得分'], 
                     marker='o', linewidth=2, markersize=8)
    axes[0, 0].set_title('绿色发展综合得分时间趋势', fontsize=14, fontweight='bold')
    axes[0, 0].set_xlabel('年份')
    axes[0, 0].set_ylabel('得分')
    axes[0, 0].grid(True, alpha=0.3)
    
    axes[0, 1].plot(yearly_avg.index, yearly_avg['经济子系统得分'], 
                     marker='s', linewidth=2, markersize=8, label='经济子系统', color='blue')
    axes[0, 1].plot(yearly_avg.index, yearly_avg['环境子系统得分'], 
                     marker='^', linewidth=2, markersize=8, label='环境子系统', color='green')
    axes[0, 1].set_title('子系统得分时间趋势', fontsize=14, fontweight='bold')
    axes[0, 1].set_xlabel('年份')
    axes[0, 1].set_ylabel('得分')
    axes[0, 1].legend()
    axes[0, 1].grid(True, alpha=0.3)
    
    axes[1, 0].plot(yearly_avg.index, yearly_avg['耦合度'], 
                     marker='D', linewidth=2, markersize=8, label='耦合度', color='orange')
    axes[1, 0].plot(yearly_avg.index, yearly_avg['耦合协调度'], 
                     marker='*', linewidth=2, markersize=10, label='耦合协调度', color='red')
    axes[1, 0].set_title('耦合协调度时间趋势', fontsize=14, fontweight='bold')
    axes[1, 0].set_xlabel('年份')
    axes[1, 0].set_ylabel('协调度')
    axes[1, 0].legend()
    axes[1, 0].grid(True, alpha=0.3)
    
    # 协调等级分布饼图
    level_counts = results['协调等级'].value_counts()
    axes[1, 1].pie(level_counts.values, labels=level_counts.index, autopct='%1.1f%%', startangle=90)
    axes[1, 1].set_title('协调等级分布', fontsize=14, fontweight='bold')
    
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '时间演化特征总图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    
    # 图表2: 省份得分热力图
    pivot_df = results.pivot(index='省份', columns='年份', values='综合得分')
    fig, ax = plt.subplots(figsize=(16, 12))
    sns.heatmap(pivot_df, annot=True, fmt='.4f', cmap='RdYlGn', center=0.5, ax=ax)
    ax.set_title('各省份绿色发展综合得分热力图', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '省份得分热力图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    
    # 图表3: 箱线图
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    sns.boxplot(x='年份', y='综合得分', data=results, ax=axes[0, 0])
    axes[0, 0].set_title('综合得分年度分布', fontsize=12, fontweight='bold')
    axes[0, 0].set_xlabel('年份')
    axes[0, 0].set_ylabel('综合得分')
    
    sns.boxplot(x='年份', y='经济子系统得分', data=results, ax=axes[0, 1])
    axes[0, 1].set_title('经济子系统得分年度分布', fontsize=12, fontweight='bold')
    axes[0, 1].set_xlabel('年份')
    axes[0, 1].set_ylabel('经济子系统得分')
    
    sns.boxplot(x='年份', y='环境子系统得分', data=results, ax=axes[1, 0])
    axes[1, 0].set_title('环境子系统得分年度分布', fontsize=12, fontweight='bold')
    axes[1, 0].set_xlabel('年份')
    axes[1, 0].set_ylabel('环境子系统得分')
    
    sns.boxplot(x='年份', y='耦合协调度', data=results, ax=axes[1, 1])
    axes[1, 1].set_title('耦合协调度年度分布', fontsize=12, fontweight='bold')
    axes[1, 1].set_xlabel('年份')
    axes[1, 1].set_ylabel('耦合协调度')
    
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '年度分布箱线图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    
    # 图表4: 散点图 - 经济与环境子系统关系
    fig, ax = plt.subplots(figsize=(10, 8))
    scatter = ax.scatter(results['经济子系统得分'], results['环境子系统得分'], 
                         c=results['年份'], cmap='viridis', alpha=0.6, s=100)
    plt.colorbar(scatter, label='年份')
    ax.set_xlabel('经济子系统得分', fontsize=12)
    ax.set_ylabel('环境子系统得分', fontsize=12)
    ax.set_title('经济与环境子系统关系散点图', fontsize=14, fontweight='bold')
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '经济环境关系散点图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"额外可视化图表已保存: {vis_dir}")


if __name__ == '__main__':
    main()
