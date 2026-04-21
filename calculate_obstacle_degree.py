"""
障碍度分析脚本
基于标准化指标数据和熵权法计算障碍度因子，生成表9和图8
"""

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import warnings
import os
from datetime import datetime

warnings.filterwarnings('ignore')

# 设置中文字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
matplotlib.rcParams['figure.dpi'] = 300
matplotlib.rcParams['savefig.dpi'] = 300


# ============================================================
# 1. 指标元数据
# ============================================================

INDICATOR_INFO = {
    '人均GDP': {'子系统': '经济', '方向': '正向', '单位': '元'},
    '第三产业占比': {'子系统': '经济', '方向': '正向', '单位': '%'},
    'R&D占比': {'子系统': '经济', '方向': '正向', '单位': '%'},
    '单位GDP能耗': {'子系统': '环境', '方向': '负向', '单位': '吨标准煤/万元'},
    '单位GDP水耗': {'子系统': '环境', '方向': '负向', '单位': '立方米/万元'},
    '污水处理率': {'子系统': '环境', '方向': '正向', '单位': '%'},
    '垃圾处理率': {'子系统': '环境', '方向': '正向', '单位': '%'},
    '森林覆盖率': {'子系统': '环境', '方向': '正向', '单位': '%'},
    '优良天数比例': {'子系统': '环境', '方向': '正向', '单位': '%'},
    '环保投资占比': {'子系统': '环境', '方向': '正向', '单位': '%'},
}

POSITIVE_COLS = ['人均GDP', '第三产业占比', 'R&D占比', '污水处理率', '垃圾处理率', '森林覆盖率', '优良天数比例', '环保投资占比']
NEGATIVE_COLS = ['单位GDP能耗', '单位GDP水耗']
ALL_INDICATORS = POSITIVE_COLS + NEGATIVE_COLS


# ============================================================
# 2. 障碍度计算
# ============================================================

def calculate_obstacle_degree(df_master):
    """
    计算障碍度
    
    障碍度模型:
    1. 标准化数据已存在 (X_ij')
    2. 计算偏离度: I_ij = 1 - X_ij' (标准化值与最优值的差距)
    3. 计算障碍度: O_ij = (W_j × I_ij) / Σ(W_j × I_ij) × 100%
    4. 综合障碍度: O_i = Σ(W_j × I_ij)
    
    其中 W_j 为熵权法计算的权重
    """
    print("=" * 60)
    print("障碍度计算")
    print("=" * 60)
    
    # 提取标准化数据
    norm_cols = [f'{col}_标准化' for col in ALL_INDICATORS]
    X = df_master[norm_cols].values.copy()
    
    # 熵权法计算权重
    m, n = X.shape
    P = X / (X.sum(axis=0) + 1e-10)
    k = 1 / np.log(m)
    E = -k * (P * np.log(P + 1e-10)).sum(axis=0)
    g = 1 - E
    weights = g / g.sum()
    
    # 计算偏离度 (标准化值与最优值1的差距)
    # I_ij = 1 - X_ij'
    I = 1 - X
    
    # 计算各指标障碍贡献值
    # W_j × I_ij
    obstacle_contribution = weights * I  # broadcasting: (n,) * (m, n) -> (m, n)
    
    # 计算综合障碍度 (每个样本的总障碍度)
    comprehensive_obstacle = obstacle_contribution.sum(axis=1)
    
    # 计算各指标的障碍度占比 (对于每个样本)
    # O_ij = (W_j × I_ij) / Σ(W_j × I_ij) × 100%
    row_sums = obstacle_contribution.sum(axis=1, keepdims=True)
    row_sums = np.where(row_sums == 0, 1e-10, row_sums)  # 避免除零
    obstacle_ratio = obstacle_contribution / row_sums * 100
    
    print(f"\n计算完成: {m}条记录, {n}个指标")
    print(f"\n综合障碍度统计:")
    print(f"  均值: {comprehensive_obstacle.mean():.4f}")
    print(f"  标准差: {comprehensive_obstacle.std():.4f}")
    print(f"  范围: [{comprehensive_obstacle.min():.4f}, {comprehensive_obstacle.max():.4f}]")
    
    print(f"\n各指标平均障碍度占比:")
    for i, col in enumerate(ALL_INDICATORS):
        avg_ratio = obstacle_ratio[:, i].mean()
        print(f"  {col}: {avg_ratio:.2f}%")
    
    # 添加到数据框
    df_master['综合障碍度'] = comprehensive_obstacle
    
    for i, col in enumerate(ALL_INDICATORS):
        df_master[f'{col}_障碍度'] = obstacle_ratio[:, i]
    
    return df_master, weights, I, obstacle_contribution, obstacle_ratio


# ============================================================
# 3. 表格生成
# ============================================================

def generate_table9(df_master, weights, obstacle_ratio, output_path):
    """
    表9：障碍度因子排序
    按平均障碍度降序排列
    """
    print("\n生成 表9：障碍度因子排序...")
    
    # 计算各指标的平均障碍度
    avg_obstacle = []
    for i, col in enumerate(ALL_INDICATORS):
        avg_ratio = obstacle_ratio[:, i].mean()
        std_ratio = obstacle_ratio[:, i].std()
        max_ratio = obstacle_ratio[:, i].max()
        min_ratio = obstacle_ratio[:, i].min()
        avg_obstacle.append({
            '指标名称': col,
            '所属子系统': INDICATOR_INFO[col]['子系统'],
            '指标方向': INDICATOR_INFO[col]['方向'],
            '权重': weights[i],
            '平均障碍度(%)': avg_ratio,
            '标准差': std_ratio,
            '最大值': max_ratio,
            '最小值': min_ratio,
            '排序': 0
        })
    
    df_obstacle = pd.DataFrame(avg_obstacle)
    df_obstacle = df_obstacle.sort_values('平均障碍度(%)', ascending=False).reset_index(drop=True)
    df_obstacle['排序'] = range(1, len(df_obstacle) + 1)
    
    # 重新排序列
    df_obstacle = df_obstacle[['排序', '指标名称', '所属子系统', '指标方向', '权重', '平均障碍度(%)', '标准差', '最大值', '最小值']]
    
    # Excel保存
    df_obstacle.to_excel(output_path, index=False)
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式
    md_lines = [
        "## 表9 障碍度因子排序\n\n",
        "| 排序 | 指标名称 | 所属子系统 | 指标方向 | 权重 | 平均障碍度(%) | 标准差 | 最大值 | 最小值 |\n",
        "|------|---------|-----------|---------|------|-------------|--------|--------|--------|\n"
    ]
    
    for _, row in df_obstacle.iterrows():
        md_lines.append(f"| {row['排序']} | {row['指标名称']} | {row['所属子系统']} | {row['指标方向']} | "
                        f"{row['权重']:.4f} | {row['平均障碍度(%)']:.2f} | {row['标准差']:.2f} | "
                        f"{row['最大值']:.2f} | {row['最小值']:.2f} |\n")
    
    # 添加说明
    md_lines.append(f"\n**排序规则**: 按平均障碍度降序排列，障碍度越高表明该指标对绿色发展水平的制约作用越大\n\n")
    md_lines.append(f"**主要障碍因子**: {', '.join(df_obstacle.head(3)['指标名称'].tolist())}是制约绿色发展水平提升的主要障碍因子\n")
    
    return ''.join(md_lines), df_obstacle


# ============================================================
# 4. 图表生成
# ============================================================

def plot_fig8_obstacle_bar(df_obstacle, output_path):
    """
    图8：障碍因子柱状图
    展示各指标障碍度大小对比
    """
    print("\n生成 图8：障碍因子柱状图...")
    
    # 按障碍度排序
    df_sorted = df_obstacle.sort_values('平均障碍度(%)', ascending=True)
    
    # 设置颜色（按子系统）
    subsystem_colors = {
        '经济': '#2E86AB',
        '环境': '#F18F01'
    }
    
    bar_colors = [subsystem_colors[subsys] for subsys in df_sorted['所属子系统']]
    
    fig, ax = plt.subplots(figsize=(12, 7))
    
    bars = ax.barh(range(len(df_sorted)), df_sorted['平均障碍度(%)'].values,
                   color=bar_colors, edgecolor='white', linewidth=0.5, height=0.6)
    
    # 数值标注
    for i, (bar, val) in enumerate(zip(bars, df_sorted['平均障碍度(%)'])):
        ax.text(val + 0.2, i, f'{val:.2f}%', va='center', ha='left',
                fontsize=10, fontweight='bold')
    
    ax.set_yticks(range(len(df_sorted)))
    ax.set_yticklabels(df_sorted['指标名称'], fontsize=11)
    ax.set_xlabel('平均障碍度 (%)', fontsize=13, fontweight='bold')
    ax.set_title('图8 各指标障碍度对比分析', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlim(0, max(df_sorted['平均障碍度(%)']) * 1.25)
    ax.grid(True, axis='x', alpha=0.3, linestyle='--')
    
    # 图例
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=color, label=f'{subsys}子系统') 
                       for subsys, color in subsystem_colors.items()]
    ax.legend(handles=legend_elements, loc='lower right', fontsize=11, framealpha=0.9)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


def plot_fig8_grouped_bar(df_master, obstacle_ratio, output_path):
    """
    图8补充：按子系统分组的障碍度对比图（按年份）
    """
    print("\n生成 图8补充：障碍度年度对比图...")
    
    # 按年份计算各指标平均障碍度
    yearly_obstacle = df_master.groupby('年份')[[f'{col}_障碍度' for col in ALL_INDICATORS]].mean()
    
    fig, ax = plt.subplots(figsize=(12, 7))
    
    # 分经济/环境两组
    econ_cols = [f'{col}_障碍度' for col in POSITIVE_COLS[:3]]
    env_cols = [f'{col}_障碍度' for col in ALL_INDICATORS if col not in POSITIVE_COLS[:3]]
    
    years = sorted(df_master['年份'].unique())
    x = np.arange(len(years))
    width = 0.35
    
    econ_avg = yearly_obstacle[econ_cols].mean(axis=1).values
    env_avg = yearly_obstacle[env_cols].mean(axis=1).values
    
    bars1 = ax.bar(x - width/2, econ_avg, width, label='经济子系统障碍度', color='#2E86AB')
    bars2 = ax.bar(x + width/2, env_avg, width, label='环境子系统障碍度', color='#F18F01')
    
    # 数值标注
    for bar in bars1:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 0.3,
                f'{height:.2f}%', ha='center', va='bottom', fontsize=9, fontweight='bold')
    for bar in bars2:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 0.3,
                f'{height:.2f}%', ha='center', va='bottom', fontsize=9, fontweight='bold')
    
    ax.set_xticks(x)
    ax.set_xticklabels([str(y) for y in years], fontsize=11)
    ax.set_xlabel('年份', fontsize=13, fontweight='bold')
    ax.set_ylabel('平均障碍度 (%)', fontsize=13, fontweight='bold')
    ax.set_title('图8 经济与环境子系统障碍度年度对比', fontsize=14, fontweight='bold', pad=15)
    ax.legend(fontsize=11, framealpha=0.9)
    ax.set_ylim(0, max(max(econ_avg), max(env_avg)) * 1.2)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    plt.savefig(output_path.replace('.png', '_年度对比.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path.replace('.png', '_年度对比.png')}")


# ============================================================
# 5. 主程序
# ============================================================

def main():
    print("=" * 60)
    print("障碍度分析工具")
    print("=" * 60)
    
    # 加载主数据
    print("\n加载综合分析主表...")
    master_file = '综合数据分析/02_面板数据/综合分析主表.xlsx'
    df_master = pd.read_excel(master_file)
    print(f"  数据形状: {df_master.shape}")
    
    # 计算障碍度
    df_master, weights, I, obstacle_contribution, obstacle_ratio = calculate_obstacle_degree(df_master)
    
    # 输出目录
    coord_output_dir = '综合数据分析/06_耦合协调度'
    chart_output_dir = '综合数据分析/07_可视化图表'
    report_output_dir = '综合数据分析/08_分析报告'
    
    # 生成表9
    md_table9, df_obstacle = generate_table9(
        df_master, weights, obstacle_ratio,
        f'{coord_output_dir}/表9_障碍度因子排序.xlsx'
    )
    
    # 生成图8
    plot_fig8_obstacle_bar(df_obstacle, f'{chart_output_dir}/图8_障碍因子柱状图.png')
    plot_fig8_grouped_bar(df_master, obstacle_ratio, f'{chart_output_dir}/图8_障碍因子柱状图.png')
    
    # 生成汇总报告
    print("\n生成障碍度分析报告...")
    report_path = f'{report_output_dir}/障碍度分析结果汇总.md'
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("# 障碍度分析结果汇总\n\n")
        f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"**数据来源**: 综合分析主表 (31省份 × 5年 × 10指标)\n\n")
        f.write(f"**计算方法**: 障碍度模型 (Obstacle Degree Model)\n\n")
        
        # 计算公式
        f.write("## 计算方法\n\n")
        f.write("### 障碍度计算公式\n\n")
        f.write("1. **偏离度计算**: $I_{ij} = 1 - X'_{ij}$\n\n")
        f.write("   其中 $X'_{ij}$ 为标准化后的指标值\n\n")
        f.write("2. **障碍度计算**: $O_{ij} = \\frac{W_j \\times I_{ij}}{\\sum_{j=1}^{n}(W_j \\times I_{ij})} \\times 100\\%$\n\n")
        f.write("   其中 $W_j$ 为熵权法计算的指标权重\n\n")
        f.write("3. **综合障碍度**: $O_i = \\sum_{j=1}^{n}(W_j \\times I_{ij})$\n\n")
        f.write("   表示各评价指标对绿色发展水平的综合阻碍程度\n\n")
        
        f.write("---\n\n")
        f.write(md_table9)
        f.write("\n\n---\n\n")
        
        # 障碍度统计摘要
        f.write("## 障碍度统计摘要\n\n")
        f.write("| 统计量 | 综合障碍度 |\n")
        f.write("|--------|-----------|\n")
        f.write(f"| 均值 | {df_master['综合障碍度'].mean():.4f} |\n")
        f.write(f"| 最大值 | {df_master['综合障碍度'].max():.4f} |\n")
        f.write(f"| 最小值 | {df_master['综合障碍度'].min():.4f} |\n")
        f.write(f"| 标准差 | {df_master['综合障碍度'].std():.4f} |\n\n")
        
        # 年度障碍度趋势
        f.write("## 年度障碍度变化趋势\n\n")
        f.write("| 年份 | 综合障碍度 | 经济子系统障碍度 | 环境子系统障碍度 |\n")
        f.write("|------|-----------|----------------|----------------|\n")
        
        for year in sorted(df_master['年份'].unique()):
            df_year = df_master[df_master['年份'] == year]
            overall = df_year['综合障碍度'].mean()
            econ_cols = [f'{col}_障碍度' for col in POSITIVE_COLS[:3]]
            env_cols = [f'{col}_障碍度' for col in ALL_INDICATORS if col not in POSITIVE_COLS[:3]]
            econ_obs = df_year[econ_cols].mean().mean()
            env_obs = df_year[env_cols].mean().mean()
            f.write(f"| {year} | {overall:.2f}% | {econ_obs:.2f}% | {env_obs:.2f}% |\n")
        
        f.write("\n---\n\n")
        f.write("*注：障碍度越高，表明该指标对绿色发展水平的制约作用越大*\n")
    
    print(f"  ✓ 报告已保存: {report_path}")
    
    print("\n" + "=" * 60)
    print("障碍度分析完成！")
    print("=" * 60)
    print(f"\n输出文件:")
    print(f"  表9: {coord_output_dir}/表9_障碍度因子排序.xlsx")
    print(f"  图8: {chart_output_dir}/图8_障碍因子柱状图.png")
    print(f"  图8补充: {chart_output_dir}/图8_障碍因子柱状图_年度对比.png")
    print(f"  汇总报告: {report_path}")
    
    # 打印TOP3障碍因子
    print("\n" + "=" * 60)
    print("TOP3障碍因子:")
    print("=" * 60)
    for i, row in df_obstacle.head(3).iterrows():
        print(f"  第{row['排序']}名: {row['指标名称']} (障碍度: {row['平均障碍度(%)']:.2f}%)")


if __name__ == '__main__':
    main()
