"""
空间分析脚本
基于综合得分数据执行核密度估计(KDE)、全局Moran's I、莫兰散点图和LISA分析
"""

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from scipy import stats
from scipy.stats import gaussian_kde
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
# 1. 省份坐标（用于空间权重矩阵）
# ============================================================

# 各省会城市近似经纬度坐标
PROVINCE_COORDS = {
    '北京': (116.4, 39.9), '天津': (117.2, 39.1), '河北': (115.5, 38.0),
    '山西': (112.5, 37.9), '内蒙古': (111.7, 40.8),
    '辽宁': (123.4, 41.8), '吉林': (125.3, 43.9), '黑龙江': (126.6, 45.8),
    '上海': (121.5, 31.2), '江苏': (118.8, 32.1), '浙江': (120.2, 30.3),
    '安徽': (117.3, 31.8), '福建': (119.3, 26.1), '江西': (115.9, 28.7),
    '山东': (117.0, 36.7), '河南': (113.7, 34.8), '湖北': (114.3, 30.6),
    '湖南': (112.9, 28.2), '广东': (113.3, 23.1), '广西': (108.3, 22.8),
    '海南': (110.3, 20.0), '重庆': (106.5, 29.6), '四川': (104.1, 30.6),
    '贵州': (106.7, 26.6), '云南': (102.7, 25.0), '西藏': (91.1, 29.7),
    '陕西': (108.9, 34.3), '甘肃': (103.8, 36.1), '青海': (101.8, 36.6),
    '宁夏': (106.3, 38.5), '新疆': (87.6, 43.8)
}

# 省份邻接关系（简化版空间权重）
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
# 2. 空间权重矩阵构建
# ============================================================

def build_spatial_weights(provinces):
    """
    构建空间权重矩阵（queen邻接）
    """
    n = len(provinces)
    W = np.zeros((n, n))
    
    for i, p1 in enumerate(provinces):
        if p1 in PROVINCE_NEIGHBORS:
            for j, p2 in enumerate(provinces):
                if p2 in PROVINCE_NEIGHBORS[p1]:
                    W[i, j] = 1
    
    # 行标准化
    row_sums = W.sum(axis=1, keepdims=True)
    row_sums = np.where(row_sums == 0, 1, row_sums)
    W = W / row_sums
    
    return W


# ============================================================
# 3. 核密度估计(KDE)
# ============================================================

def calculate_kde(scores, output_path):
    """
    计算并绘制核密度估计图
    """
    print("\n生成 图9：核密度估计结果图...")
    
    years = list(range(2018, 2023))
    all_scores_by_year = []
    
    fig, ax = plt.subplots(figsize=(10, 6))
    
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#006400']
    
    for idx, year in enumerate(years):
        scores_year = scores[scores['年份'] == year]['综合得分'].values
        if len(scores_year) > 1:
            kde = gaussian_kde(scores_year, bw_method='silverman')
            x_vals = np.linspace(scores['综合得分'].min() - 0.05, 
                                scores['综合得分'].max() + 0.05, 200)
            y_vals = kde(x_vals)
            ax.plot(x_vals, y_vals, color=colors[idx], linewidth=2.5, 
                   label=f'{year}年', alpha=0.85)
            ax.fill_between(x_vals, y_vals, alpha=0.15, color=colors[idx])
            all_scores_by_year.extend(scores_year.tolist())
    
    # 总体KDE
    all_scores = scores['综合得分'].values
    kde_all = gaussian_kde(all_scores, bw_method='silverman')
    x_all = np.linspace(all_scores.min() - 0.05, all_scores.max() + 0.05, 200)
    y_all = kde_all(x_all)
    ax.plot(x_all, y_all, 'k-', linewidth=3, label='总体', alpha=0.9, zorder=5)
    ax.fill_between(x_all, y_all, alpha=0.1, color='black', zorder=4)
    
    # 统计信息标注
    mean_score = all_scores.mean()
    std_score = all_scores.std()
    ax.axvline(x=mean_score, color='red', linestyle='--', linewidth=1.5, alpha=0.6,
               label=f'均值={mean_score:.3f}')
    ax.axvline(x=mean_score - std_score, color='gray', linestyle=':', linewidth=1, alpha=0.4)
    ax.axvline(x=mean_score + std_score, color='gray', linestyle=':', linewidth=1, alpha=0.4)
    
    ax.set_xlabel('绿色发展综合得分', fontsize=13, fontweight='bold')
    ax.set_ylabel('概率密度', fontsize=13, fontweight='bold')
    ax.set_title('图9 绿色发展水平核密度估计分布（2018-2022年）', fontsize=14, fontweight='bold', pad=15)
    ax.legend(loc='best', fontsize=10, framealpha=0.9)
    ax.grid(True, alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


# ============================================================
# 4. 全局Moran's I
# ============================================================

def calculate_moran_I(y, W):
    """
    计算全局Moran's I指数
    
    I = (n / W_sum) * (y' * W * y) / (y' * y)
    
    其中y为标准化后的变量
    """
    n = len(y)
    
    # 标准化变量
    y_mean = y.mean()
    y_std = y.std()
    if y_std == 0:
        return 0, 0, 1, 0
    
    z = (y - y_mean) / y_std
    
    # 计算分子和分母
    W_sum = W.sum()
    numerator = z @ W @ z
    denominator = z @ z
    
    I = (n / W_sum) * (numerator / denominator)
    
    # 期望值
    E_I = -1 / (n - 1)
    
    # 方差（正态性假设）
    S1 = 0.5 * ((W + W.T) ** 2).sum()
    S2 = (W.sum(axis=1) + W.sum(axis=0)) ** 2
    S2 = S2.sum()
    S3 = ((W ** 2).sum(axis=1) + (W.T ** 2).sum(axis=1)) ** 2
    S3 = S3.sum()
    
    # 使用简化方差公式
    n_sq = n * n
    variance = (n_sq * S1 - n * S2 + 3 * W_sum ** 2) / (W_sum ** 2 * (n_sq - 1)) - E_I ** 2
    variance = max(variance, 1e-10)  # 防止负值
    
    # Z值
    Z = (I - E_I) / np.sqrt(variance)
    
    # P值（双尾检验）
    p_value = 2 * (1 - stats.norm.cdf(abs(Z)))
    
    return I, Z, p_value, E_I


def generate_table10(moran_results, output_path):
    """
    表10：全局Moran's I指数
    """
    print("\n生成 表10：全局Moran's I指数表...")
    
    table_data = []
    for year, result in moran_results.items():
        I, Z, p, E_I = result
        significance = ''
        if p < 0.01:
            significance = '*** (1%水平显著)'
        elif p < 0.05:
            significance = '** (5%水平显著)'
        elif p < 0.1:
            significance = '* (10%水平显著)'
        else:
            significance = '不显著'
        
        table_data.append({
            '年份': year,
            "Moran's I值": round(I, 4),
            '期望值E(I)': round(E_I, 4),
            '标准化Z值': round(Z, 4),
            'P值': round(p, 4),
            '显著性': significance
        })
    
    df = pd.DataFrame(table_data)
    df.to_excel(output_path, index=False)
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式
    md_lines = [
        "## 表10 全局Moran's I空间自相关检验\n\n",
        "| 年份 | Moran's I值 | 期望值E(I) | 标准化Z值 | P值 | 显著性 |\n",
        "|------|------------|-----------|-----------|-----|-------|\n"
    ]
    
    for _, row in df.iterrows():
        md_lines.append(f"| {row['年份']} | {row['Moran\'s I值']:.4f} | "
                        f"{row['期望值E(I)']:.4f} | {row['标准化Z值']:.4f} | "
                        f"{row['P值']:.4f} | {row['显著性']} |\n")
    
    return ''.join(md_lines)


# ============================================================
# 5. 莫兰散点图
# ============================================================

def plot_moran_scatter(scores, W, output_path, year=2021):
    """
    图10：莫兰散点图
    """
    print(f"\n生成 图10：莫兰散点图 ({year}年)...")
    
    df_year = scores[scores['年份'] == year].copy()
    provinces = df_year['省份'].values
    z = (df_year['综合得分'].values - df_year['综合得分'].mean()) / df_year['综合得分'].std()
    
    # 空间滞后值
    W_z = W @ z
    
    fig, ax = plt.subplots(figsize=(8, 8))
    
    # 散点
    scatter = ax.scatter(z, W_z, c='steelblue', alpha=0.7, s=80, edgecolors='white', linewidth=1, zorder=3)
    
    # 标注省份名称
    for i, province in enumerate(provinces):
        ax.annotate(province, (z[i], W_z[i]), textcoords="offset points",
                    xytext=(5, 5), fontsize=8, alpha=0.8)
    
    # 回归线
    slope, intercept, r_value, p_value, std_err = stats.linregress(z, W_z)
    x_fit = np.linspace(z.min() - 0.1, z.max() + 0.1, 100)
    y_fit = slope * x_fit + intercept
    ax.plot(x_fit, y_fit, 'r--', linewidth=2, alpha=0.8, 
            label=f'回归线 (斜率={slope:.3f}, R²={r_value**2:.3f})', zorder=2)
    
    # 象限分割线
    ax.axhline(y=0, color='gray', linestyle='-', alpha=0.3, linewidth=0.5)
    ax.axvline(x=0, color='gray', linestyle='-', alpha=0.3, linewidth=0.5)
    
    # 象限标签
    ax.text(0.1, ax.get_ylim()[1] * 0.8, 'H-H', fontsize=12, fontweight='bold', alpha=0.5)
    ax.text(ax.get_xlim()[0] * 0.7, ax.get_ylim()[1] * 0.8, 'L-H', fontsize=12, fontweight='bold', alpha=0.5)
    ax.text(ax.get_xlim()[0] * 0.7, ax.get_ylim()[0] * 0.7, 'L-L', fontsize=12, fontweight='bold', alpha=0.5)
    ax.text(0.1, ax.get_ylim()[0] * 0.7, 'H-L', fontsize=12, fontweight='bold', alpha=0.5)
    
    ax.set_xlabel('标准化综合得分 (Z)', fontsize=13, fontweight='bold')
    ax.set_ylabel('空间滞后值 (W×Z)', fontsize=13, fontweight='bold')
    ax.set_title(f'图10 {year}年省域绿色发展Moran散点图', fontsize=14, fontweight='bold', pad=15)
    ax.legend(loc='best', fontsize=10, framealpha=0.9)
    ax.grid(True, alpha=0.2, linestyle='--')
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


# ============================================================
# 6. LISA分析
# ============================================================

def calculate_LISA(y, W):
    """
    计算局部Moran's I (LISA)
    
    I_i = z_i * Σ(W_ij * z_j)
    """
    n = len(y)
    y_mean = y.mean()
    y_std = y.std()
    if y_std == 0:
        return np.zeros(n), np.zeros(n, dtype=int), np.zeros(n)
    
    z = (y - y_mean) / y_std
    W_z = W @ z
    
    # 局部Moran's I
    I_local = z * W_z
    
    # 判定象限类型
    # 1: H-H (高-高), 2: L-H (低-高), 3: L-L (低-低), 4: H-L (高-低)
    quadrant = np.zeros(n, dtype=int)
    for i in range(n):
        if z[i] > 0 and W_z[i] > 0:
            quadrant[i] = 1  # H-H
        elif z[i] < 0 and W_z[i] > 0:
            quadrant[i] = 2  # L-H
        elif z[i] < 0 and W_z[i] < 0:
            quadrant[i] = 3  # L-L
        elif z[i] > 0 and W_z[i] < 0:
            quadrant[i] = 4  # H-L
    
    return I_local, quadrant, z


def plot_lisa_cluster(scores, W, output_path, year=2021):
    """
    图11：LISA集聚图
    """
    print(f"\n生成 图11：LISA集聚图 ({year}年)...")
    
    df_year = scores[scores['年份'] == year].copy()
    provinces = df_year['省份'].values
    scores_vals = df_year['综合得分'].values
    
    # 计算LISA
    I_local, quadrant, z = calculate_LISA(scores_vals, W)
    
    # 象限标签
    quadrant_labels = {1: 'H-H', 2: 'L-H', 3: 'L-L', 4: 'H-L'}
    quadrant_names = {1: '高-高集聚', 2: '低-高集聚', 3: '低-低集聚', 4: '高-低集聚'}
    quadrant_colors = {1: '#E41A1C', 2: '#FFBBB5', 3: '#4DAF4A', 4: '#B5E0B3'}
    
    fig, ax = plt.subplots(figsize=(10, 10))
    
    # 按区域排序（从北到南）
    province_order = sorted(provinces, key=lambda p: PROVINCE_COORDS.get(p, (0, 0))[1], reverse=True)
    indices = [list(provinces).index(p) for p in province_order]
    
    # 创建颜色映射
    colors = [quadrant_colors.get(quadrant[i], '#999999') for i in indices]
    
    # 使用气泡图展示
    y_pos = range(len(province_order))
    bars = ax.barh(y_pos, [scores_vals[i] for i in indices], color=colors,
                   edgecolor='white', linewidth=0.5, height=0.7)
    
    # 标注省份名称和象限类型
    for i, (idx, pos) in enumerate(zip(indices, y_pos)):
        ax.text(scores_vals[idx] + 0.01, pos, f'{province_order[i]} ({quadrant_labels[quadrant[idx]]})',
                va='center', ha='left', fontsize=9)
    
    ax.set_yticks(y_pos)
    ax.set_yticklabels([])
    ax.set_xlabel('绿色发展综合得分', fontsize=13, fontweight='bold')
    ax.set_title(f'图11 {year}年省域绿色发展LISA集聚图', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlim(0, max(scores_vals) * 1.25)
    ax.grid(True, axis='x', alpha=0.3, linestyle='--')
    
    # 图例
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=color, label=f'{quadrant_names[q]} (n={sum(quadrant == q)})')
                       for q, color in quadrant_colors.items() if sum(quadrant == q) > 0]
    ax.legend(handles=legend_elements, loc='lower right', fontsize=10, framealpha=0.9)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  ✓ 已保存: {output_path}")


# ============================================================
# 7. 主程序
# ============================================================

def main():
    print("=" * 60)
    print("空间分析工具 (KDE + Moran's I + LISA)")
    print("=" * 60)
    
    # 加载综合得分数据
    print("\n加载分析结果数据...")
    
    # 使用generate_tables.py计算得分
    master_file = '综合数据分析/02_面板数据/综合分析主表.xlsx'
    df_master = pd.read_excel(master_file)
    
    # 重新计算TOPSIS得分
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
    composite_scores = D_minus / (D_plus + D_minus)
    
    df_master['综合得分'] = composite_scores
    
    # 构建空间权重矩阵
    provinces = df_master[df_master['年份'] == 2021]['省份'].values
    W = build_spatial_weights(provinces)
    print(f"\n空间权重矩阵: {W.shape[0]} × {W.shape[1]}")
    print(f"  邻接省份对数: {int(W.sum() * W.shape[0])}")
    
    # 输出目录
    chart_output_dir = '综合数据分析/07_可视化图表'
    report_output_dir = '综合数据分析/08_分析报告'
    coord_output_dir = '综合数据分析/06_耦合协调度'
    
    # 1. 图9：核密度估计
    calculate_kde(df_master, f'{chart_output_dir}/图9_核密度估计结果图.png')
    
    # 2. 表10：全局Moran's I
    moran_results = {}
    for year in sorted(df_master['年份'].unique()):
        df_year = df_master[df_master['年份'] == year]
        provinces_year = df_year['省份'].values
        W_year = build_spatial_weights(provinces_year)
        scores_year = df_year['综合得分'].values
        
        I, Z, p, E_I = calculate_moran_I(scores_year, W_year)
        moran_results[year] = (I, Z, p, E_I)
        print(f"\n{year}年 Moran's I = {I:.4f}, Z = {Z:.4f}, p = {p:.4f}")
    
    md_table10 = generate_table10(
        moran_results,
        f'{coord_output_dir}/表10_全局MoransI指数.xlsx'
    )
    
    # 3. 图10：莫兰散点图 (2021年)
    provinces_2021 = df_master[df_master['年份'] == 2021]['省份'].values
    W_2021 = build_spatial_weights(provinces_2021)
    
    plot_moran_scatter(df_master, W_2021, f'{chart_output_dir}/图10_莫兰散点图.png', year=2021)
    
    # 4. 图11：LISA集聚图 (2021年)
    plot_lisa_cluster(df_master, W_2021, f'{chart_output_dir}/图11_LISA集聚图.png', year=2021)
    
    # 汇总报告
    print("\n生成空间分析报告...")
    report_path = f'{report_output_dir}/空间分析结果汇总.md'
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("# 空间分析结果汇总\n\n")
        f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"**数据来源**: 综合分析主表 (31省份 × 5年)\n\n")
        f.write(f"**空间权重**: Queen邻接矩阵(省份邻接关系)\n\n")
        
        # 核密度估计
        f.write("## 核密度估计(KDE)\n\n")
        f.write("核密度估计是一种非参数估计方法，用于描述数据的概率密度分布。\n\n")
        f.write("使用Silverman带宽选择器，对各年度综合得分进行核密度估计。\n\n")
        
        f.write("---\n\n")
        f.write(md_table10)
        f.write("\n\n---\n\n")
        
        # LISA统计
        f.write("## LISA局部空间关联统计\n\n")
        df_2021 = df_master[df_master['年份'] == 2021].copy()
        provinces = df_2021['省份'].values
        I_local, quadrant, z = calculate_LISA(df_2021['综合得分'].values, W_2021)
        
        quadrant_counts = pd.Series(quadrant).value_counts().sort_index()
        quadrant_names = {1: 'H-H (高-高集聚)', 2: 'L-H (低-高集聚)', 3: 'L-L (低-低集聚)', 4: 'H-L (高-低集聚)'}
        
        f.write("| 集聚类型 | 数量 | 占比 |\n")
        f.write("|---------|------|------|\n")
        for q, count in quadrant_counts.items():
            f.write(f"| {quadrant_names.get(q, '未知')} | {count} | {count/len(provinces)*100:.1f}% |\n")
        
        f.write("\n---\n\n")
        f.write("*注：空间权重基于省份邻接关系构建的Queen邻接矩阵*\n")
    
    print(f"  ✓ 报告已保存: {report_path}")
    
    print("\n" + "=" * 60)
    print("空间分析完成！")
    print("=" * 60)
    print(f"\n输出文件:")
    print(f"  图9: {chart_output_dir}/图9_核密度估计结果图.png")
    print(f"  表10: {coord_output_dir}/表10_全局MoransI指数.xlsx")
    print(f"  图10: {chart_output_dir}/图10_莫兰散点图.png")
    print(f"  图11: {chart_output_dir}/图11_LISA集聚图.png")
    print(f"  汇总报告: {report_path}")


if __name__ == '__main__':
    main()
