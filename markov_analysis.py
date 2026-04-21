"""
马尔可夫转移矩阵分析脚本
基于绿色发展得分数据构建传统和空间马尔可夫链，分析等级转移规律
"""

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
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
# 1. 省份邻接关系
# ============================================================

PROVINCE_NEIGHBORS = {
    '北京': ['天津', '河北'],
    '天津': ['北京', '河北'],
    '河北': ['北京', '天津', '辽宁', '内蒙古', '山西', '河南', '山东'],
    '山西': ['河北', '内蒙古', '河南', '陕西'],
    '内蒙古': ['辽宁', '吉林', '黑龙江', '河北', '山西', '陕西', '宁夏', '甘肃'],
    '辽宁': ['内蒙古', '吉林', '河北'],
    '吉林': ['辽宁', '黑龙江', '内蒙古'],
    '黑龙江': ['吉林', '内蒙古'],
    '上海': ['江苏', '浙江'],
    '江苏': ['山东', '安徽', '浙江', '上海'],
    '浙江': ['江苏', '安徽', '江西', '福建', '上海'],
    '安徽': ['江苏', '山东', '河南', '湖北', '江西', '浙江'],
    '福建': ['浙江', '江西', '广东'],
    '江西': ['安徽', '浙江', '福建', '广东', '湖南', '湖北'],
    '山东': ['河北', '河南', '安徽', '江苏'],
    '河南': ['河北', '山西', '陕西', '湖北', '安徽', '山东'],
    '湖北': ['河南', '陕西', '重庆', '湖南', '江西', '安徽'],
    '湖南': ['湖北', '重庆', '贵州', '广西', '广东', '江西'],
    '广东': ['福建', '江西', '湖南', '广西', '海南'],
    '广西': ['湖南', '贵州', '云南', '广东'],
    '海南': ['广东'],
    '重庆': ['四川', '贵州', '湖南', '湖北', '陕西'],
    '四川': ['重庆', '贵州', '云南', '西藏', '青海', '甘肃', '陕西'],
    '贵州': ['重庆', '四川', '云南', '广西', '湖南'],
    '云南': ['四川', '贵州', '广西', '西藏'],
    '西藏': ['四川', '云南', '青海', '新疆'],
    '陕西': ['内蒙古', '山西', '河南', '湖北', '重庆', '四川', '甘肃', '宁夏'],
    '甘肃': ['内蒙古', '宁夏', '陕西', '四川', '青海', '新疆'],
    '青海': ['甘肃', '四川', '西藏', '新疆'],
    '宁夏': ['内蒙古', '陕西', '甘肃'],
    '新疆': ['西藏', '青海', '甘肃']
}


# ============================================================
# 2. 等级划分
# ============================================================

def classify_by_quantiles(scores, n_types=4):
    """
    基于分位数划分等级
    默认4类：低(0-25%)、中低(25-50%)、中高(50-75%)、高(75-100%)
    """
    quantiles = np.percentile(scores, [i * 100 / n_types for i in range(n_types + 1)])
    
    # 确保分位数唯一且严格递增
    quantiles = np.unique(quantiles)
    if len(quantiles) < n_types + 1:
        quantiles = np.linspace(scores.min() - 0.001, scores.max() + 0.001, n_types + 1)
    
    # 扩展范围以确保所有值都被分配
    quantiles[0] = scores.min() - 0.001
    quantiles[-1] = scores.max() + 0.001
    
    grade_names = {0: '低水平', 1: '中低水平', 2: '中高水平', 3: '高水平'}
    
    # 使用np.digitize避免NaN问题
    labels_int = np.digitize(scores, quantiles[1:-1], right=True)
    
    return labels_int, quantiles, grade_names


# ============================================================
# 3. 传统马尔可夫转移矩阵
# ============================================================

def build_traditional_markov(df_with_types, years):
    """
    构建传统马尔可夫转移矩阵
    """
    # 获取最大类型数
    n_types = df_with_types['type'].max() + 1
    
    # 转移计数矩阵
    transition_count = np.zeros((n_types, n_types))
    
    for province in df_with_types['省份'].unique():
        prov_data = df_with_types[df_with_types['省份'] == province].sort_values('年份')
        types = prov_data['type'].values
        
        for t in range(len(types) - 1):
            if types[t] >= 0 and types[t+1] >= 0:
                transition_count[types[t], types[t+1]] += 1
    
    # 转移概率矩阵（行标准化）
    transition_prob = np.zeros_like(transition_count, dtype=float)
    for i in range(n_types):
        row_sum = transition_count[i].sum()
        if row_sum > 0:
            transition_prob[i] = transition_count[i] / row_sum
    
    return transition_count, transition_prob


# ============================================================
# 4. 空间马尔可夫转移矩阵
# ============================================================

def build_spatial_markov(df_with_types, years, neighbor_dict):
    """
    构建空间马尔可夫转移矩阵
    根据邻域背景类型分层构建转移矩阵
    """
    n_types = df_with_types['type'].max() + 1
    
    # 为每个时间点构建空间滞后类型
    spatial_matrices = []
    
    for year in years:
        df_year = df_with_types[df_with_types['年份'] == year].set_index('省份')
        
        # 计算每个省份的邻域平均类型（空间滞后）
        spatial_lag = {}
        for province in df_year.index:
            neighbors = neighbor_dict.get(province, [])
            neighbor_types = []
            for nb in neighbors:
                if nb in df_year.index:
                    neighbor_types.append(df_year.loc[nb, 'type'])
            
            if neighbor_types:
                spatial_lag[province] = np.mean(neighbor_types)
            else:
                spatial_lag[province] = df_year.loc[province, 'type']  # 无邻居则用自身
        
        # 将空间滞后也划分为类型
        lag_values = pd.Series(spatial_lag)
        lag_bins = lag_values.quantile([i/n_types for i in range(n_types+1)]).unique()
        if len(lag_bins) < n_types + 1:
            lag_bins = np.linspace(lag_values.min() - 0.001, lag_values.max() + 0.001, n_types + 1)
        lag_bins[0] = lag_values.min() - 0.001
        lag_bins[-1] = lag_values.max() + 0.001
        lag_types_arr = np.digitize(lag_values.values, lag_bins[1:-1], right=True)
        lag_types = pd.Series(lag_types_arr, index=lag_values.index)
        
        spatial_matrices.append(lag_types)
    
    # 构建空间马尔可夫转移矩阵（按邻域背景分层）
    spatial_transition = {}
    spatial_count = {}
    
    for lag_type in range(n_types):
        spatial_transition[lag_type] = np.zeros((n_types, n_types))
        spatial_count[lag_type] = np.zeros((n_types, n_types))
    
    for i, year in enumerate(years[:-1]):  # t到t+1
        df_t = df_with_types[df_with_types['年份'] == year].set_index('省份')
        df_t1 = df_with_types[df_with_types['年份'] == years[i+1]].set_index('省份')
        lag_types_i = spatial_matrices[i]
        
        for province in df_t.index:
            if province in df_t1.index and province in lag_types_i.index:
                type_t = df_t.loc[province, 'type']
                type_t1 = df_t1.loc[province, 'type']
                lag_type = lag_types_i[province]
                
                if type_t >= 0 and type_t1 >= 0 and lag_type >= 0:
                    spatial_count[lag_type][type_t, type_t1] += 1
    
    # 计算概率
    for lag_type in range(n_types):
        for i in range(n_types):
            row_sum = spatial_count[lag_type][i].sum()
            if row_sum > 0:
                spatial_transition[lag_type][i] = spatial_count[lag_type][i] / row_sum
    
    return spatial_count, spatial_transition


# ============================================================
# 5. 表格生成
# ============================================================

def generate_table11(transition_count, transition_prob, grade_names, output_path):
    """
    表11：传统马尔可夫转移矩阵
    """
    print("\n生成 表11：传统马尔可夫转移概率矩阵...")
    
    n_types = len(grade_names)
    
    # Markdown格式
    md_lines = [
        "## 表11 传统马尔可夫转移概率矩阵\n\n",
        f"**转移概率（P_ij）**：表示从状态i转移到状态j的概率\n\n",
        f"**样本期**：2018-2022年，共4期转移\n\n",
        "### (a) 转移概率矩阵\n\n",
        "| 年初状态\\年末状态 | " + " | ".join([grade_names.get(i, f'类型{i+1}') for i in range(n_types)]) + " |\n",
        "|-------------------|" + "|".join(["------" for _ in range(n_types)]) + "|\n"
    ]
    
    for i in range(n_types):
        row = [f"{transition_prob[i, j]:.4f}" for j in range(n_types)]
        md_lines.append(f"| {grade_names.get(i, f'类型{i+1}')} | " + " | ".join(row) + " |\n")
    
    md_lines.append("\n### (b) 转移频数矩阵\n\n")
    md_lines.append("| 年初状态\\年末状态 | " + " | ".join([grade_names.get(i, f'类型{i+1}') for i in range(n_types)]) + " |\n")
    md_lines.append("|-------------------|" + "|".join(["------" for _ in range(n_types)]) + "|\n")
    
    for i in range(n_types):
        row = [str(int(transition_count[i, j])) for j in range(n_types)]
        md_lines.append(f"| {grade_names.get(i, f'类型{i+1}')} | " + " | ".join(row) + " |\n")
    
    # Excel格式
    # 概率矩阵
    prob_data = {}
    prob_data['年初状态'] = [grade_names.get(i, f'类型{i+1}') for i in range(n_types)]
    for j in range(n_types):
        prob_data[grade_names.get(j, f'类型{j+1}')] = [transition_prob[i, j] for i in range(n_types)]
    pd.DataFrame(prob_data).to_excel(output_path, index=False, sheet_name='转移概率')
    
    print(f"  ✓ 已保存: {output_path}")
    return ''.join(md_lines)


def generate_table12(spatial_count, spatial_transition, grade_names, output_path):
    """
    表12：空间马尔可夫转移矩阵
    """
    print("\n生成 表12：空间马尔可夫转移概率矩阵...")
    
    n_types = len(grade_names)
    lag_names = {0: '低水平邻域', 1: '中低水平邻域', 2: '中高水平邻域', 3: '高水平邻域'}
    
    md_lines = [
        "## 表12 空间马尔可夫转移概率矩阵\n\n",
        "**空间马尔可夫转移概率**：考虑邻域背景条件下的状态转移概率\n\n",
    ]
    
    # Excel writer
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for lag_type in range(n_types):
            sheet_name = f'邻域{lag_names.get(lag_type, lag_type+1)}'[:31]
            
            prob_data = {}
            prob_data['年初状态'] = [grade_names.get(i, f'类型{i+1}') for i in range(n_types)]
            for j in range(n_types):
                prob_data[grade_names.get(j, f'类型{j+1}')] = [spatial_transition[lag_type][i, j] for i in range(n_types)]
            pd.DataFrame(prob_data).to_excel(writer, sheet_name=sheet_name, index=False)
            
            md_lines.append(f"### {lag_names.get(lag_type, f'邻域背景{lag_type+1}')}\n\n")
            md_lines.append("| 年初状态\\年末状态 | " + " | ".join([grade_names.get(i, f'类型{i+1}') for i in range(n_types)]) + " |\n")
            md_lines.append("|-------------------|" + "|".join(["------" for _ in range(n_types)]) + "|\n")
            
            for i in range(n_types):
                row = [f"{spatial_transition[lag_type][i, j]:.4f}" for j in range(n_types)]
                md_lines.append(f"| {grade_names.get(i, f'类型{i+1}')} | " + " | ".join(row) + " |\n")
            
            md_lines.append("\n")
    
    print(f"  ✓ 已保存: {output_path}")
    return ''.join(md_lines)


# ============================================================
# 6. 热力图生成
# ============================================================

def plot_heatmap12(transition_prob, grade_names, output_path):
    """
    热力图12：传统马尔可夫转移矩阵热力图
    """
    print("\n生成 热力图12：传统马尔可夫转移矩阵热力图...")
    
    n_types = len(grade_names)
    labels = [grade_names.get(i, f'类型{i+1}') for i in range(n_types)]
    
    fig, ax = plt.subplots(figsize=(8, 6))
    
    # 使用seaborn绘制热力图
    sns.heatmap(transition_prob, annot=True, fmt='.4f', cmap='YlOrRd',
                xticklabels=labels, yticklabels=labels,
                cbar_kws={'label': '转移概率'}, ax=ax,
                linewidths=0.5, linecolor='white',
                vmin=0, vmax=1)
    
    ax.set_xlabel('年末状态', fontsize=13, fontweight='bold')
    ax.set_ylabel('年初状态', fontsize=13, fontweight='bold')
    ax.set_title('热力图12 传统马尔可夫转移概率矩阵', fontsize=14, fontweight='bold', pad=15)
    ax.set_xticklabels(labels, rotation=30, ha='right')
    ax.set_yticklabels(labels, rotation=0)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


def plot_heatmap13(spatial_transition, grade_names, output_path):
    """
    热力图13：空间马尔可夫转移矩阵热力图（多子图）
    """
    print("\n生成 热力图13：空间马尔可夫转移矩阵热力图...")
    
    n_types = len(grade_names)
    labels = [grade_names.get(i, f'类型{i+1}') for i in range(n_types)]
    lag_names = {0: '低水平邻域', 1: '中低水平邻域', 2: '中高水平邻域', 3: '高水平邻域'}
    
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    axes = axes.flatten()
    
    for lag_type in range(n_types):
        ax = axes[lag_type]
        sns.heatmap(spatial_transition[lag_type], annot=True, fmt='.4f', 
                   cmap='YlOrRd', xticklabels=labels, yticklabels=labels,
                   cbar_kws={'label': '转移概率'}, ax=ax,
                   linewidths=0.5, linecolor='white',
                   vmin=0, vmax=1)
        
        ax.set_xlabel('年末状态', fontsize=11, fontweight='bold')
        ax.set_ylabel('年初状态', fontsize=11, fontweight='bold')
        ax.set_title(f'{lag_names.get(lag_type, f"邻域{lag_type+1}")}', 
                    fontsize=12, fontweight='bold')
        ax.set_xticklabels(labels, rotation=30, ha='right')
        ax.set_yticklabels(labels, rotation=0)
    
    fig.suptitle('热力图13 空间马尔可夫转移概率矩阵', fontsize=14, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


# ============================================================
# 7. 主程序
# ============================================================

def main():
    print("=" * 60)
    print("马尔可夫转移矩阵分析工具")
    print("=" * 60)
    
    # 加载数据
    print("\n加载分析数据...")
    master_file = '综合数据分析/02_面板数据/综合分析主表.xlsx'
    df_master = pd.read_excel(master_file)
    
    # 计算TOPSIS得分
    POSITIVE_COLS = ['人均GDP', '第三产业占比', 'R&D占比', '污水处理率', '垃圾处理率', '森林覆盖率', '优良天数比例', '环保投资占比']
    NEGATIVE_COLS = ['单位GDP能耗', '单位GDP水耗']
    ALL_INDICATORS = POSITIVE_COLS + NEGATIVE_COLS
    
    norm_cols = [f'{col}_标准化' for col in ALL_INDICATORS]
    normalized_matrix = df_master[norm_cols].values
    
    m, n = normalized_matrix.shape
    P = normalized_matrix / (normalized_matrix.sum(axis=0) + 1e-10)
    k = 1 / np.log(m)
    E = -k * (P * np.log(P + 1e-10)).sum(axis=0)
    g = 1 - E
    weights = g / g.sum()
    
    V = normalized_matrix * weights
    V_plus = V.max(axis=0)
    V_minus = V.min(axis=0)
    D_plus = np.sqrt(((V - V_plus) ** 2).sum(axis=1))
    D_minus = np.sqrt(((V - V_minus) ** 2).sum(axis=1))
    df_master['综合得分'] = D_minus / (D_plus + D_minus)
    
    years = sorted(df_master['年份'].unique())
    print(f"  数据形状: {df_master.shape}")
    print(f"  年份: {years}")
    
    # 等级划分（4分位）
    print("\n执行等级划分（4分位法）...")
    all_scores = df_master['综合得分'].values
    types, quantiles, grade_names = classify_by_quantiles(all_scores, n_types=4)
    df_master['type'] = types
    
    print(f"  分位数阈值: {quantiles}")
    print(f"  等级分布:")
    for i in range(4):
        count = sum(types == i)
        print(f"    {grade_names.get(i, f'类型{i+1}')}: {count}条 ({count/len(types)*100:.1f}%)")
    
    # 传统马尔可夫
    print("\n构建传统马尔可夫转移矩阵...")
    transition_count, transition_prob = build_traditional_markov(df_master, years)
    
    # 空间马尔可夫
    print("\n构建空间马尔可夫转移矩阵...")
    spatial_count, spatial_transition = build_spatial_markov(df_master, years, PROVINCE_NEIGHBORS)
    
    # 输出目录
    coord_output_dir = '综合数据分析/06_耦合协调度'
    chart_output_dir = '综合数据分析/07_可视化图表'
    report_output_dir = '综合数据分析/08_分析报告'
    
    # 生成表11
    md_table11 = generate_table11(transition_count, transition_prob, grade_names,
                                  f'{coord_output_dir}/表11_传统马尔可夫转移矩阵.xlsx')
    
    # 生成表12
    md_table12 = generate_table12(spatial_count, spatial_transition, grade_names,
                                  f'{coord_output_dir}/表12_空间马尔可夫转移矩阵.xlsx')
    
    # 生成热力图12
    plot_heatmap12(transition_prob, grade_names,
                   f'{chart_output_dir}/热力图12_传统马尔可夫转移矩阵.png')
    
    # 生成热力图13
    plot_heatmap13(spatial_transition, grade_names,
                   f'{chart_output_dir}/热力图13_空间马尔可夫转移矩阵.png')
    
    # 汇总报告
    print("\n生成马尔可夫分析报告...")
    report_path = f'{report_output_dir}/马尔可夫转移矩阵分析汇总.md'
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("# 马尔可夫转移矩阵分析汇总\n\n")
        f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"**数据来源**: 综合分析主表 (31省份 × 5年)\n\n")
        f.write(f"**等级划分**: 4分位法（低水平、中低水平、中高水平、高水平）\n\n")
        
        f.write("## 等级划分阈值\n\n")
        f.write(f"| 等级 | 下限 | 上限 |\n")
        f.write(f"|------|------|------|\n")
        for i in range(4):
            f.write(f"| {grade_names.get(i, f'类型{i+1}')} | {quantiles[i]:.4f} | {quantiles[i+1]:.4f} |\n")
        f.write("\n")
        
        f.write("---\n\n")
        f.write(md_table11)
        f.write("\n\n---\n\n")
        f.write(md_table12)
        f.write("\n\n---\n\n")
        
        f.write("*注：空间马尔可夫基于Queen邻接矩阵构建的空间滞后类型分层*\n")
    
    print(f"  ✓ 报告已保存: {report_path}")
    
    print("\n" + "=" * 60)
    print("马尔可夫分析完成！")
    print("=" * 60)
    print(f"\n输出文件:")
    print(f"  表11: {coord_output_dir}/表11_传统马尔可夫转移矩阵.xlsx")
    print(f"  表12: {coord_output_dir}/表12_空间马尔可夫转移矩阵.xlsx")
    print(f"  热力图12: {chart_output_dir}/热力图12_传统马尔可夫转移矩阵.png")
    print(f"  热力图13: {chart_output_dir}/热力图13_空间马尔可夫转移矩阵.png")
    print(f"  汇总报告: {report_path}")


if __name__ == '__main__':
    main()
