"""
学术风格图表生成脚本
基于TOPSIS得分数据生成图2-图5
"""

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.ticker import MaxNLocator
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# Set font at module level
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False
matplotlib.rcParams['figure.dpi'] = 300
matplotlib.rcParams['savefig.dpi'] = 300

# ============================================================
# 1. 加载数据
# ============================================================

def load_analysis_data():
    """加载TOPSIS分析结果数据"""
    master_file = '综合数据分析/02_面板数据/综合分析主表.xlsx'
    df_master = pd.read_excel(master_file)
    
    # 重新计算TOPSIS得分（与generate_tables.py一致）
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
    
    # TOPSIS
    V = normalized_matrix * weights
    V_plus = V.max(axis=0)
    V_minus = V.min(axis=0)
    D_plus = np.sqrt(((V - V_plus) ** 2).sum(axis=1))
    D_minus = np.sqrt(((V - V_minus) ** 2).sum(axis=1))
    composite_scores = D_minus / (D_plus + D_minus)
    
    # 子系统得分
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
    econ_scores = D_minus_econ / (D_plus_econ + D_minus_econ)
    
    env_matrix = normalized_matrix[:, env_indices]
    V_env_plus = (env_matrix * env_w).max(axis=0)
    V_env_minus = (env_matrix * env_w).min(axis=0)
    D_plus_env = np.sqrt(((env_matrix * env_w - V_env_plus) ** 2).sum(axis=1))
    D_minus_env = np.sqrt(((env_matrix * env_w - V_env_minus) ** 2).sum(axis=1))
    env_scores = D_minus_env / (D_plus_env + D_minus_env)
    
    # 添加到数据框
    df_master['综合得分'] = composite_scores
    df_master['经济子系统得分'] = econ_scores
    df_master['环境子系统得分'] = env_scores
    
    print(f"数据加载完成: {df_master.shape}")
    print(f"年份: {sorted(df_master['年份'].unique())}")
    print(f"区域: {sorted(df_master['区域'].unique())}")
    
    return df_master


# ============================================================
# 2. 图表生成函数
# ============================================================

def get_chinese_font_prop(size=12, weight='normal'):
    """获取中文字体属性"""
    return fm.FontProperties(family='SimHei', size=size, weight=weight)


def plot_fig2_national_trend(df, output_path):
    """
    图2：全国时序趋势图
    展示全国及各子系统得分的时间演化趋势
    """
    print("\n生成 图2：全国时序趋势图...")
    
    years = sorted(df['年份'].unique())
    national_avg = df.groupby('年份')[['综合得分', '经济子系统得分', '环境子系统得分']].mean()
    
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # 绘制三条趋势线
    ax.plot(years, national_avg['综合得分'], 'ko-', linewidth=2.5, markersize=8, 
            label='综合得分', zorder=3)
    ax.plot(years, national_avg['经济子系统得分'], 'bs-', linewidth=2, markersize=7,
            label='经济子系统', zorder=2)
    ax.plot(years, national_avg['环境子系统得分'], 'g^-', linewidth=2, markersize=7,
            label='环境子系统', zorder=1)
    
    # 添加数值标注
    for year in years:
        y1 = national_avg.loc[year, '综合得分']
        ax.annotate(f'{y1:.3f}', (year, y1), textcoords="offset points", 
                    xytext=(0, 12), ha='center', fontsize=9, fontweight='bold')
    
    # 设置坐标轴
    ax.set_xlabel('年份', fontsize=13, fontweight='bold')
    ax.set_ylabel('TOPSIS得分', fontsize=13, fontweight='bold')
    ax.set_title('图2 全国绿色发展水平时序趋势（2018-2022年）', fontsize=14, fontweight='bold', pad=15)
    ax.set_xticks(years)
    ax.set_xticklabels([str(y) for y in years])
    ax.yaxis.set_major_locator(MaxNLocator(5))
    ax.set_ylim(bottom=min(national_avg.min()) - 0.02, top=max(national_avg.max()) + 0.05)
    ax.legend(loc='lower right', fontsize=11, framealpha=0.9, edgecolor='gray')
    ax.grid(True, alpha=0.3, linestyle='--')
    
    # 添加趋势线
    z = np.polyfit(years, national_avg['综合得分'], 1)
    p = np.poly1d(z)
    ax.plot(years, p(years), 'r--', alpha=0.6, linewidth=1.5, label=f'趋势线 (斜率={z[0]:.4f})')
    ax.legend(loc='lower right', fontsize=11, framealpha=0.9, edgecolor='gray')
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


def plot_fig3_region_comparison(df, output_path):
    """
    图3：东中西部及东北地区对比图
    展示四大区域年度得分对比
    """
    print("\n生成 图3：东中西部及东北地区对比图...")
    
    region_avg = df.groupby(['年份', '区域'])['综合得分'].mean().unstack()
    regions = ['东部', '中部', '西部', '东北']
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D']
    markers = ['o', 's', '^', 'D']
    
    fig, ax = plt.subplots(figsize=(11, 6))
    
    years = sorted(df['年份'].unique())
    
    for i, region in enumerate(regions):
        if region in region_avg.columns:
            scores = region_avg[region].values
            ax.plot(years, scores, color=colors[i], marker=markers[i], 
                    linewidth=2.5, markersize=8, label=f'{region}地区', zorder=3)
            
            # 数值标注
            for year, score in zip(years, scores):
                ax.annotate(f'{score:.3f}', (year, score), textcoords="offset points",
                           xytext=(0, 10), ha='center', fontsize=8, color=colors[i], fontweight='bold')
    
    ax.set_xlabel('年份', fontsize=13, fontweight='bold')
    ax.set_ylabel('TOPSIS综合得分', fontsize=13, fontweight='bold')
    ax.set_title('图3 四大区域绿色发展水平对比（2018-2022年）', fontsize=14, fontweight='bold', pad=15)
    ax.set_xticks(years)
    ax.set_xticklabels([str(y) for y in years])
    ax.legend(loc='best', fontsize=11, framealpha=0.9, edgecolor='gray')
    ax.grid(True, alpha=0.3, linestyle='--')
    
    # 添加区域均值标注
    for i, region in enumerate(regions):
        if region in region_avg.columns:
            mean_score = region_avg[region].mean()
            ax.axhline(y=mean_score, color=colors[i], linestyle=':', alpha=0.4)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


def plot_fig4_province_distribution(df, output_path, year=2021):
    """
    图4：2021年省域空间分布图（横向柱状图）
    按区域分组的省域得分排序展示
    """
    print(f"\n生成 图4：{year}年省域空间分布图...")
    
    df_year = df[df['年份'] == year].copy()
    df_year = df_year.sort_values('综合得分', ascending=True)
    
    # 按区域设置颜色
    region_colors = {
        '东部': '#2E86AB',
        '中部': '#A23B72',
        '西部': '#F18F01',
        '东北': '#C73E1D'
    }
    
    fig, ax = plt.subplots(figsize=(12, 10))
    
    bars = ax.barh(df_year['省份'], df_year['综合得分'], 
                   color=[region_colors[r] for r in df_year['区域']],
                   edgecolor='white', linewidth=0.5)
    
    # 添加数值标注
    for bar, score in zip(bars, df_year['综合得分']):
        ax.text(bar.get_width() + 0.005, bar.get_y() + bar.get_height()/2,
                f'{score:.3f}', va='center', ha='left', fontsize=8, fontweight='bold')
    
    ax.set_xlabel('TOPSIS综合得分', fontsize=13, fontweight='bold')
    ax.set_title(f'图4 {year}年省域绿色发展水平空间分布', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlim(0, max(df_year['综合得分']) * 1.15)
    
    # 添加图例
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=color, label=f'{region}地区') 
                       for region, color in region_colors.items()]
    ax.legend(handles=legend_elements, loc='lower right', fontsize=10, framealpha=0.9)
    
    ax.grid(True, axis='x', alpha=0.3, linestyle='--')
    ax.set_yticklabels(df_year['省份'], fontsize=9)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


def plot_fig5_subsystem_comparison(df, output_path):
    """
    图5：四大子系统对比图
    展示31省份经济与各环境子指标得分对比
    """
    print("\n生成 图5：四大子系统对比图...")
    
    # 计算更细的子维度
    # 经济子系统：人均GDP、第三产业占比、R&D占比
    # 环境子系统：污水处理率、垃圾处理率、森林覆盖率、优良天数比例、环保投资占比、单位GDP能耗、单位GDP水耗
    
    # 这里使用经济子系统得分和环境子系统得分做分组柱状图
    province_avg = df.groupby('省份')[['经济子系统得分', '环境子系统得分']].mean()
    province_avg = province_avg.sort_values('经济子系统得分', ascending=True)
    
    fig, ax = plt.subplots(figsize=(14, 8))
    
    y_pos = np.arange(len(province_avg))
    width = 0.35
    
    bars1 = ax.barh(y_pos - width/2, province_avg['经济子系统得分'], width,
                    label='经济子系统', color='#2E86AB', edgecolor='white', linewidth=0.5)
    bars2 = ax.barh(y_pos + width/2, province_avg['环境子系统得分'], width,
                    label='环境子系统', color='#F18F01', edgecolor='white', linewidth=0.5)
    
    ax.set_yticks(y_pos)
    ax.set_yticklabels(province_avg.index, fontsize=9)
    ax.set_xlabel('TOPSIS得分', fontsize=13, fontweight='bold')
    ax.set_title('图5 各省份经济与子系统得分对比（年均值）', fontsize=14, fontweight='bold', pad=15)
    ax.legend(loc='lower right', fontsize=11, framealpha=0.9)
    ax.grid(True, axis='x', alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


# ============================================================
# 3. 主程序
# ============================================================

def main():
    print("=" * 60)
    print("学术风格图表生成工具")
    print("=" * 60)
    
    # 加载数据
    df = load_analysis_data()
    
    # 输出目录
    output_dir = '综合数据分析/07_可视化图表'
    
    # 生成图2
    plot_fig2_national_trend(df, f'{output_dir}/图2_全国时序趋势图.png')
    
    # 生成图3
    plot_fig3_region_comparison(df, f'{output_dir}/图3_东中西部及东北地区对比图.png')
    
    # 生成图4
    plot_fig4_province_distribution(df, f'{output_dir}/图4_2021年省域空间分布图.png', year=2021)
    
    # 生成图5
    plot_fig5_subsystem_comparison(df, f'{output_dir}/图5_四大子系统对比图.png')
    
    print("\n" + "=" * 60)
    print("所有图表生成完成！")
    print("=" * 60)
    print(f"\n输出目录: {output_dir}")
    print("  - 图2_全国时序趋势图.png")
    print("  - 图3_东中西部及东北地区对比图.png")
    print("  - 图4_2021年省域空间分布图.png")
    print("  - 图5_四大子系统对比图.png")


if __name__ == '__main__':
    main()
