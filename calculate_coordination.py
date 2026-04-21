"""
耦合协调度分析脚本
基于四大子系统得分计算耦合度、协调度，生成表7-表8、图6-图7
"""

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import warnings
import os

warnings.filterwarnings('ignore')

# 设置中文字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
matplotlib.rcParams['figure.dpi'] = 300
matplotlib.rcParams['savefig.dpi'] = 300


# ============================================================
# 1. 耦合协调度等级标准
# ============================================================

COORDINATION_LEVELS = [
    (0.0, 0.1, '极度失调'),
    (0.1, 0.2, '严重失调'),
    (0.2, 0.3, '中度失调'),
    (0.3, 0.4, '轻度失调'),
    (0.4, 0.5, '濒临失调'),
    (0.5, 0.6, '勉强协调'),
    (0.6, 0.7, '初级协调'),
    (0.7, 0.8, '中级协调'),
    (0.8, 0.9, '良好协调'),
    (0.9, 1.0, '优质协调'),
]


def get_coordination_level(D):
    """根据协调度值获取等级"""
    for lower, upper, level in COORDINATION_LEVELS:
        if lower <= D < upper:
            return level
    return '优质协调'


# ============================================================
# 2. 耦合协调度计算
# ============================================================

def calculate_coupling_coordination(df_master):
    """
    计算耦合协调度
    基于经济子系统和环境子系统得分
    """
    print("=" * 60)
    print("耦合协调度计算")
    print("=" * 60)
    
    # 确保有子系统得分
    if '经济子系统得分' not in df_master.columns or '环境子系统得分' not in df_master.columns:
        print("错误: 缺少子系统得分数据")
        return None
    
    u1 = df_master['经济子系统得分'].values
    u2 = df_master['环境子系统得分'].values
    
    # 计算耦合度 C
    # C = 2 * sqrt(u1 * u2) / (u1 + u2)
    # 避免除零
    denominator = u1 + u2
    mask = denominator > 0
    C = np.zeros_like(u1)
    C[mask] = 2 * np.sqrt(u1[mask] * u2[mask]) / denominator[mask]
    
    # 限制C在[0,1]范围内
    C = np.clip(C, 0, 1)
    
    # 计算综合发展指数 T
    # T = alpha * u1 + beta * u2, 通常取 alpha = beta = 0.5
    alpha, beta = 0.5, 0.5
    T = alpha * u1 + beta * u2
    
    # 计算耦合协调度 D
    # D = sqrt(C * T)
    D = np.sqrt(C * T)
    
    # 限制D在[0,1]范围内
    D = np.clip(D, 0, 1)
    
    # 划分等级
    levels = [get_coordination_level(d) for d in D]
    
    # 添加到数据框
    df_master['耦合度'] = C
    df_master['综合发展指数'] = T
    df_master['耦合协调度'] = D
    df_master['协调等级'] = levels
    
    print(f"\n计算完成: {len(df_master)}条记录")
    print(f"\n耦合度统计: 均值={C.mean():.4f}, 标准差={C.std():.4f}")
    print(f"综合发展指数统计: 均值={T.mean():.4f}, 标准差={T.std():.4f}")
    print(f"耦合协调度统计: 均值={D.mean():.4f}, 标准差={D.std():.4f}")
    
    # 等级分布
    print("\n协调等级分布:")
    level_counts = pd.Series(levels).value_counts().sort_index()
    for level, count in level_counts.items():
        pct = count / len(df_master) * 100
        print(f"  {level}: {count}条 ({pct:.1f}%)")
    
    return df_master


# ============================================================
# 3. 表格生成
# ============================================================

def generate_table7(output_path):
    """表7：协调等级标准表"""
    print("\n生成 表7：协调等级标准表...")
    
    # Excel格式
    table_data = []
    for lower, upper, level in COORDINATION_LEVELS:
        table_data.append({
            '等级序号': len(table_data) + 1,
            '协调等级': level,
            '协调度区间': f'[{lower:.1f}, {upper:.1f})',
            '协调度下限': lower,
            '协调度上限': upper,
            '等级描述': get_level_description(level)
        })
    
    df = pd.DataFrame(table_data)
    df.to_excel(output_path, index=False)
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式
    md_lines = [
        "## 表7 耦合协调度等级划分标准\n\n",
        "| 等级序号 | 协调等级 | 协调度区间 | 等级描述 |\n",
        "|---------|---------|-----------|--------|\n"
    ]
    
    for i, (lower, upper, level) in enumerate(COORDINATION_LEVELS, 1):
        md_lines.append(f"| {i} | {level} | [{lower:.1f}, {upper:.1f}) | {get_level_description(level)} |\n")
    
    return ''.join(md_lines)


def get_level_description(level):
    """获取等级描述"""
    descriptions = {
        '极度失调': '经济与环境系统严重不协调，发展极不平衡',
        '严重失调': '经济与环境系统高度不协调，发展严重失衡',
        '中度失调': '经济与环境系统中度不协调，发展存在明显失衡',
        '轻度失调': '经济与环境系统轻度不协调，发展略有失衡',
        '濒临失调': '经济与环境系统处于协调与失调的临界状态',
        '勉强协调': '经济与环境系统基本协调，但协调程度较低',
        '初级协调': '经济与环境系统初步实现协调发展',
        '中级协调': '经济与环境系统协调发展水平中等',
        '良好协调': '经济与环境系统协调发展水平较高',
        '优质协调': '经济与环境系统高度协调，发展非常均衡',
    }
    return descriptions.get(level, '')


def generate_table8(df_master, output_path):
    """表8：各省协调度与等级表"""
    print("\n生成 表8：各省协调度与等级表...")
    
    # 按省份计算年均值
    province_avg = df_master.groupby('省份').agg({
        '耦合度': 'mean',
        '综合发展指数': 'mean',
        '耦合协调度': 'mean'
    }).reset_index()
    
    # 添加等级
    province_avg['协调等级'] = province_avg['耦合协调度'].apply(get_coordination_level)
    
    # 添加区域信息
    region_map = df_master.groupby('省份')['区域'].first().reset_index()
    province_avg = province_avg.merge(region_map, on='省份', how='left')
    
    # 按协调度排序
    province_avg = province_avg.sort_values('耦合协调度', ascending=False).reset_index(drop=True)
    province_avg['排名'] = range(1, len(province_avg) + 1)
    
    # 重排顺序
    province_avg = province_avg[['排名', '省份', '区域', '耦合度', '综合发展指数', '耦合协调度', '协调等级']]
    
    # Excel保存
    province_avg.to_excel(output_path, index=False)
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式
    md_lines = [
        "## 表8 各省份耦合协调度及等级（年均值）\n\n",
        "| 排名 | 省份 | 区域 | 耦合度 | 综合发展指数 | 耦合协调度 | 协调等级 |\n",
        "|------|------|------|--------|-------------|-----------|--------|\n"
    ]
    
    for _, row in province_avg.iterrows():
        md_lines.append(f"| {row['排名']} | {row['省份']} | {row['区域']} | "
                        f"{row['耦合度']:.4f} | {row['综合发展指数']:.4f} | "
                        f"{row['耦合协调度']:.4f} | {row['协调等级']} |\n")
    
    return ''.join(md_lines)


# ============================================================
# 4. 图表生成
# ============================================================

def plot_fig6_coordination_trend(df, output_path):
    """
    图6：协调度时序图
    展示各省份协调度随时间变化趋势
    """
    print("\n生成 图6：协调度时序图...")
    
    # 按区域分组展示（31条线太密集，按区域均值展示更清晰）
    region_avg = df.groupby(['年份', '区域'])['耦合协调度'].mean().unstack()
    regions = ['东部', '中部', '西部', '东北']
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D']
    markers = ['o', 's', '^', 'D']
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
    
    # 上图：四大区域趋势
    for i, region in enumerate(regions):
        if region in region_avg.columns:
            scores = region_avg[region].values
            ax1.plot(sorted(df['年份'].unique()), scores, color=colors[i], 
                     marker=markers[i], linewidth=2.5, markersize=8, label=f'{region}地区')
            
            # 数值标注
            for year, score in zip(sorted(df['年份'].unique()), scores):
                ax1.annotate(f'{score:.3f}', (year, score), textcoords="offset points",
                           xytext=(0, 10), ha='center', fontsize=8, color=colors[i], fontweight='bold')
    
    ax1.set_ylabel('耦合协调度', fontsize=12, fontweight='bold')
    ax1.set_title('图6 耦合协调度时序演化特征', fontsize=14, fontweight='bold', pad=15)
    ax1.legend(loc='best', fontsize=11, framealpha=0.9)
    ax1.grid(True, alpha=0.3, linestyle='--')
    ax1.axhline(y=0.6, color='green', linestyle=':', alpha=0.5, label='初级协调阈值')
    ax1.axhline(y=0.8, color='blue', linestyle=':', alpha=0.5, label='良好协调阈值')
    
    # 下图：全国均值+标准差范围
    national_avg = df.groupby('年份')['耦合协调度'].agg(['mean', 'std']).reset_index()
    years = sorted(df['年份'].unique())
    
    ax2.fill_between(years, national_avg['mean'] - national_avg['std'], 
                     national_avg['mean'] + national_avg['std'], 
                     alpha=0.3, color='#2E86AB', label='±1标准差')
    ax2.plot(years, national_avg['mean'], 'ko-', linewidth=2.5, markersize=8, label='全国均值')
    
    for year, mean, std in zip(years, national_avg['mean'], national_avg['std']):
        ax2.annotate(f'{mean:.3f}', (year, mean), textcoords="offset points",
                    xytext=(0, 12), ha='center', fontsize=9, fontweight='bold')
    
    ax2.set_xlabel('年份', fontsize=12, fontweight='bold')
    ax2.set_ylabel('耦合协调度', fontsize=12, fontweight='bold')
    ax2.set_title('全国耦合协调度变化趋势', fontsize=13, fontweight='bold')
    ax2.legend(loc='best', fontsize=11, framealpha=0.9)
    ax2.grid(True, alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


def plot_fig7_coordination_map(df, output_path, year=2021):
    """
    图7：协调度空间分布图
    使用分级柱状图/散点图展示各省份协调度
    注：无GIS地图数据时使用替代方案
    """
    print(f"\n生成 图7：协调度空间分布图 ({year}年)...")
    
    df_year = df[df['年份'] == year].copy()
    
    # 按区域和协调度排序
    df_year = df_year.sort_values(['区域', '耦合协调度'], ascending=[True, True])
    
    # 定义颜色映射
    color_map = {
        '极度失调': '#8B0000',
        '严重失调': '#FF4500',
        '中度失调': '#FF6347',
        '轻度失调': '#FFA07A',
        '濒临失调': '#FFD700',
        '勉强协调': '#90EE90',
        '初级协调': '#3CB371',
        '中级协调': '#2E8B57',
        '良好协调': '#228B22',
        '优质协调': '#006400',
    }
    
    fig, ax = plt.subplots(figsize=(14, 10))
    
    # 创建水平柱状图
    province_names = df_year['省份'].values
    coord_values = df_year['耦合协调度'].values
    coord_levels = df_year['协调等级'].values
    
    bar_colors = [color_map.get(level, '#808080') for level in coord_levels]
    
    bars = ax.barh(range(len(province_names)), coord_values, color=bar_colors,
                   edgecolor='white', linewidth=0.5, height=0.7)
    
    # 添加数值和等级标注
    for i, (val, level) in enumerate(zip(coord_values, coord_levels)):
        ax.text(val + 0.005, i, f'{val:.3f} ({level})', va='center', ha='left',
                fontsize=8, fontweight='bold')
    
    ax.set_yticks(range(len(province_names)))
    ax.set_yticklabels(province_names, fontsize=10)
    ax.set_xlabel('耦合协调度', fontsize=13, fontweight='bold')
    ax.set_title(f'图7 {year}年省域耦合协调度空间分布', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlim(0, max(coord_values) * 1.35)
    ax.grid(True, axis='x', alpha=0.3, linestyle='--')
    
    # 添加图例
    from matplotlib.patches import Patch
    legend_elements = []
    for level, color in color_map.items():
        if level in coord_levels:
            legend_elements.append(Patch(facecolor=color, label=level))
    
    ax.legend(handles=legend_elements, loc='lower right', fontsize=8, framealpha=0.9,
              title='协调等级', title_fontsize=9)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


# ============================================================
# 5. 主程序
# ============================================================

def main():
    print("=" * 60)
    print("耦合协调度分析工具")
    print("=" * 60)
    
    # 加载主数据
    print("\n加载综合分析主表...")
    master_file = '综合数据分析/02_面板数据/综合分析主表.xlsx'
    df_master = pd.read_excel(master_file)
    print(f"  数据形状: {df_master.shape}")
    
    # 计算TOPSIS子系统得分（与之前一致）
    POSITIVE_COLS = ['人均GDP', '第三产业占比', 'R&D占比', '污水处理率', '垃圾处理率', '森林覆盖率', '优良天数比例', '环保投资占比']
    NEGATIVE_COLS = ['单位GDP能耗', '单位GDP水耗']
    ALL_INDICATORS = POSITIVE_COLS + NEGATIVE_COLS
    
    norm_cols = [f'{col}_标准化' for col in ALL_INDICATORS]
    normalized_matrix = df_master[norm_cols].values
    
    # 熵权法
    m, n = normalized_matrix.shape
    P = normalized_matrix / (normalized_matrix.sum(axis=0) + 1e-10)
    k = 1 / np.log(m)
    E = -k * (P * np.log(P + 1e-10)).sum(axis=0)
    g = 1 - E
    weights = g / g.sum()
    
    # TOPSIS - 经济子系统
    econ_indices = [ALL_INDICATORS.index(col) for col in POSITIVE_COLS[:3]]
    env_indices = [ALL_INDICATORS.index(col) for col in ALL_INDICATORS if col not in POSITIVE_COLS[:3]]
    
    econ_w = weights[econ_indices]
    econ_w = econ_w / econ_w.sum()
    env_w = weights[env_indices]
    env_w = env_w / env_w.sum()
    
    econ_matrix = normalized_matrix[:, econ_indices]
    V_econ_plus = (econ_matrix * econ_w).max(axis=0)
    V_econ_minus = (econ_matrix * econ_w).min(axis=0)
    D_plus_econ = np.sqrt(((econ_matrix * econ_w - V_econ_plus) ** 2).sum(axis=1))
    D_minus_econ = np.sqrt(((econ_matrix * econ_w - V_econ_minus) ** 2).sum(axis=1))
    df_master['经济子系统得分'] = D_minus_econ / (D_plus_econ + D_minus_econ)
    
    env_matrix = normalized_matrix[:, env_indices]
    V_env_plus = (env_matrix * env_w).max(axis=0)
    V_env_minus = (env_matrix * env_w).min(axis=0)
    D_plus_env = np.sqrt(((env_matrix * env_w - V_env_plus) ** 2).sum(axis=1))
    D_minus_env = np.sqrt(((env_matrix * env_w - V_env_minus) ** 2).sum(axis=1))
    df_master['环境子系统得分'] = D_minus_env / (D_plus_env + D_minus_env)
    
    # 计算耦合协调度
    df_master = calculate_coupling_coordination(df_master)
    
    if df_master is None:
        return
    
    # 保存耦合协调度结果
    coord_output_dir = '综合数据分析/06_耦合协调度'
    coord_result_file = f'{coord_output_dir}/耦合协调度计算结果.xlsx'
    df_master.to_excel(coord_result_file, index=False)
    print(f"\n  耦合协调度结果已保存: {coord_result_file}")
    
    # 生成表7
    md_table7 = generate_table7(f'{coord_output_dir}/表7_协调等级标准表.xlsx')
    
    # 生成表8
    md_table8 = generate_table8(df_master, f'{coord_output_dir}/表8_各省协调度与等级表.xlsx')
    
    # 生成图6
    chart_output_dir = '综合数据分析/07_可视化图表'
    plot_fig6_coordination_trend(df_master, f'{chart_output_dir}/图6_协调度时序图.png')
    
    # 生成图7
    plot_fig7_coordination_map(df_master, f'{chart_output_dir}/图7_协调度空间分布图.png', year=2021)
    
    # 生成汇总报告
    print("\n生成耦合协调度分析报告...")
    report_path = '综合数据分析/08_分析报告/耦合协调度分析结果汇总.md'
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("# 耦合协调度分析结果汇总\n\n")
        f.write(f"**生成时间**: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"**数据来源**: 综合分析主表 (31省份 × 5年)\n\n")
        f.write(f"**计算方法**: 耦合协调度模型 (Coupling Coordination Degree Model)\n\n")
        
        # 计算公式
        f.write("## 计算方法\n\n")
        f.write("### 耦合度计算公式\n\n")
        f.write("$$C = 2 \\times \\sqrt{\\frac{u_1 \\times u_2}{u_1 + u_2}}$$\n\n")
        f.write("其中：$u_1$ 为经济子系统得分，$u_2$ 为环境子系统得分\n\n")
        
        f.write("### 综合发展指数计算公式\n\n")
        f.write("$$T = \\alpha \\times u_1 + \\beta \\times u_2$$\n\n")
        f.write("其中：$\\alpha = \\beta = 0.5$\n\n")
        
        f.write("### 耦合协调度计算公式\n\n")
        f.write("$$D = \\sqrt{C \\times T}$$\n\n")
        
        f.write("---\n\n")
        f.write(md_table7)
        f.write("\n\n---\n\n")
        f.write(md_table8)
        f.write("\n\n---\n\n")
        
        # 统计摘要
        f.write("## 协调度统计摘要\n\n")
        f.write("| 统计量 | 耦合度 | 综合发展指数 | 耦合协调度 |\n")
        f.write("|--------|--------|-------------|-----------|\n")
        f.write(f"| 均值 | {df_master['耦合度'].mean():.4f} | {df_master['综合发展指数'].mean():.4f} | {df_master['耦合协调度'].mean():.4f} |\n")
        f.write(f"| 最大值 | {df_master['耦合度'].max():.4f} | {df_master['综合发展指数'].max():.4f} | {df_master['耦合协调度'].max():.4f} |\n")
        f.write(f"| 最小值 | {df_master['耦合度'].min():.4f} | {df_master['综合发展指数'].min():.4f} | {df_master['耦合协调度'].min():.4f} |\n")
        f.write(f"| 标准差 | {df_master['耦合度'].std():.4f} | {df_master['综合发展指数'].std():.4f} | {df_master['耦合协调度'].std():.4f} |\n\n")
        
        # 年度趋势
        f.write("## 年度协调度趋势\n\n")
        yearly = df_master.groupby('年份')[['耦合度', '综合发展指数', '耦合协调度']].mean()
        f.write("| 年份 | 耦合度 | 综合发展指数 | 耦合协调度 |\n")
        f.write("|------|--------|-------------|-----------|\n")
        for year, row in yearly.iterrows():
            f.write(f"| {year} | {row['耦合度']:.4f} | {row['综合发展指数']:.4f} | {row['耦合协调度']:.4f} |\n")
        
        f.write("\n---\n\n")
        f.write("*注：所有计算结果均基于熵权TOPSIS子系统得分，采用标准耦合协调度模型*\n")
    
    print(f"  ✓ 报告已保存: {report_path}")
    
    print("\n" + "=" * 60)
    print("耦合协调度分析完成！")
    print("=" * 60)
    print(f"\n输出文件:")
    print(f"  表7: {coord_output_dir}/表7_协调等级标准表.xlsx")
    print(f"  表8: {coord_output_dir}/表8_各省协调度与等级表.xlsx")
    print(f"  图6: {chart_output_dir}/图6_协调度时序图.png")
    print(f"  图7: {chart_output_dir}/图7_协调度空间分布图.png")
    print(f"  汇总报告: {report_path}")


if __name__ == '__main__':
    main()
