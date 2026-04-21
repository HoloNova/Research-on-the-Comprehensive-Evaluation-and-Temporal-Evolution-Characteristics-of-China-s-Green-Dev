"""
基于主数据生成规范化分析表格
包含：指标体系及权重、熵权结果、区域年均得分、省份分年得分、子系统省域结果
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# ============================================================
# 1. 配置参数
# ============================================================

# 指标元数据
INDICATOR_INFO = {
    '人均GDP': {'子系统': '经济子系统', '方向': '正向', '单位': '元'},
    '第三产业占比': {'子系统': '经济子系统', '方向': '正向', '单位': '%'},
    'R&D占比': {'子系统': '经济子系统', '方向': '正向', '单位': '%'},
    '单位GDP能耗': {'子系统': '环境子系统', '方向': '负向', '单位': '吨标准煤/万元'},
    '单位GDP水耗': {'子系统': '环境子系统', '方向': '负向', '单位': '立方米/万元'},
    '污水处理率': {'子系统': '环境子系统', '方向': '正向', '单位': '%'},
    '垃圾处理率': {'子系统': '环境子系统', '方向': '正向', '单位': '%'},
    '森林覆盖率': {'子系统': '环境子系统', '方向': '正向', '单位': '%'},
    '优良天数比例': {'子系统': '环境子系统', '方向': '正向', '单位': '%'},
    '环保投资占比': {'子系统': '环境子系统', '方向': '正向', '单位': '%'},
}

POSITIVE_COLS = ['人均GDP', '第三产业占比', 'R&D占比', '污水处理率', '垃圾处理率', '森林覆盖率', '优良天数比例', '环保投资占比']
NEGATIVE_COLS = ['单位GDP能耗', '单位GDP水耗']
ALL_INDICATORS = POSITIVE_COLS + NEGATIVE_COLS


# ============================================================
# 2. 熵权TOPSIS计算
# ============================================================

def entropy_weight_method(normalized_matrix):
    """熵权法计算权重"""
    m, n = normalized_matrix.shape
    
    # 计算比重
    P = normalized_matrix / (normalized_matrix.sum(axis=0) + 1e-10)
    
    # 计算信息熵
    k = 1 / np.log(m)
    E = -k * (P * np.log(P + 1e-10)).sum(axis=0)
    
    # 计算差异系数和权重
    g = 1 - E
    w = g / g.sum()
    
    return w, E, g


def topsis_method(normalized_matrix, weights):
    """TOPSIS方法计算综合得分"""
    m, n = normalized_matrix.shape
    
    # 加权标准化矩阵
    V = normalized_matrix * weights
    
    # 正理想解和负理想解
    V_plus = V.max(axis=0)
    V_minus = V.min(axis=0)
    
    # 计算距离
    D_plus = np.sqrt(((V - V_plus) ** 2).sum(axis=1))
    D_minus = np.sqrt(((V - V_minus) ** 2).sum(axis=1))
    
    # 综合得分
    S = D_minus / (D_plus + D_minus)
    
    return S, D_plus, D_minus, V_plus, V_minus


def compute_subsystem_scores(normalized_matrix, weights, indicators, positive_cols):
    """计算经济/环境子系统得分"""
    m = normalized_matrix.shape[0]
    
    # 经济子系统指标索引
    econ_indices = [indicators.index(col) for col in positive_cols[:3]]
    env_indices = [indicators.index(col) for col in indicators if col not in positive_cols[:3]]
    
    # 子系统权重归一化
    econ_w = weights[econ_indices]
    econ_w = econ_w / econ_w.sum()
    env_w = weights[env_indices]
    env_w = env_w / env_w.sum()
    
    # 经济子系统得分
    econ_matrix = normalized_matrix[:, econ_indices]
    V_econ = econ_matrix * econ_w
    V_econ_plus = V_econ.max(axis=0)
    V_econ_minus = V_econ.min(axis=0)
    D_plus_econ = np.sqrt(((V_econ - V_econ_plus) ** 2).sum(axis=1))
    D_minus_econ = np.sqrt(((V_econ - V_econ_minus) ** 2).sum(axis=1))
    econ_scores = D_minus_econ / (D_plus_econ + D_minus_econ)
    
    # 环境子系统得分
    env_matrix = normalized_matrix[:, env_indices]
    V_env = env_matrix * env_w
    V_env_plus = V_env.max(axis=0)
    V_env_minus = V_env.min(axis=0)
    D_plus_env = np.sqrt(((V_env - V_env_plus) ** 2).sum(axis=1))
    D_minus_env = np.sqrt(((V_env - V_env_minus) ** 2).sum(axis=1))
    env_scores = D_minus_env / (D_plus_env + D_minus_env)
    
    return econ_scores, env_scores


def run_full_analysis(df_master):
    """运行完整熵权TOPSIS分析"""
    print("=" * 60)
    print("运行熵权TOPSIS分析")
    print("=" * 60)
    
    # 提取标准化数据
    norm_cols = [f'{col}_标准化' for col in ALL_INDICATORS]
    normalized_matrix = df_master[norm_cols].values
    
    # 熵权法
    weights, entropy_vals, diff_coeffs = entropy_weight_method(normalized_matrix)
    
    print("\n熵权法计算结果:")
    for i, col in enumerate(ALL_INDICATORS):
        print(f"  {col}: 熵值={entropy_vals[i]:.4f}, 差异系数={diff_coeffs[i]:.4f}, 权重={weights[i]:.4f}")
    
    # TOPSIS
    composite_scores, D_plus, D_minus, V_plus, V_minus = topsis_method(normalized_matrix, weights)
    
    # 子系统得分
    econ_scores, env_scores = compute_subsystem_scores(
        normalized_matrix, weights, ALL_INDICATORS, POSITIVE_COLS
    )
    
    print(f"\nTOPSIS结果统计:")
    print(f"  综合得分: 均值={composite_scores.mean():.4f}, 标准差={composite_scores.std():.4f}")
    print(f"  经济子系统: 均值={econ_scores.mean():.4f}, 标准差={econ_scores.std():.4f}")
    print(f"  环境子系统: 均值={env_scores.mean():.4f}, 标准差={env_scores.std():.4f}")
    
    return {
        'weights': weights,
        'entropy': entropy_vals,
        'diff_coeffs': diff_coeffs,
        'composite_scores': composite_scores,
        'econ_scores': econ_scores,
        'env_scores': env_scores,
        'D_plus': D_plus,
        'D_minus': D_minus,
    }


# ============================================================
# 3. 表格生成
# ============================================================

def generate_table1(weights, output_path):
    """表1 指标体系及权重"""
    print("\n生成 表1 指标体系及权重...")
    
    table_data = []
    current_subsystem = None
    
    for col in ALL_INDICATORS:
        info = INDICATOR_INFO[col]
        if info['子系统'] != current_subsystem:
            current_subsystem = info['子系统']
            table_data.append({
                '一级指标': current_subsystem,
                '二级指标': '',
                '指标方向': '',
                '单位': '',
                '熵权权重': ''
            })
        
        table_data.append({
            '一级指标': '',
            '二级指标': col,
            '指标方向': info['方向'],
            '单位': info['单位'],
            '熵权权重': f'{weights[ALL_INDICATORS.index(col)]:.4f}'
        })
    
    df = pd.DataFrame(table_data)
    df.to_excel(output_path, index=False)
    print(f"  ✓ 已保存: {output_path}")
    
    # 同时生成Markdown格式
    md_lines = [
        "## 表1 指标体系及权重\n\n",
        "| 一级指标 | 二级指标 | 指标方向 | 单位 | 熵权权重 |\n",
        "|---------|---------|---------|------|--------|\n"
    ]
    
    for col in ALL_INDICATORS:
        info = INDICATOR_INFO[col]
        w = weights[ALL_INDICATORS.index(col)]
        md_lines.append(f"| {info['子系统']} | {col} | {info['方向']} | {info['单位']} | {w:.4f} |\n")
    
    return ''.join(md_lines)


def generate_table3(weights, entropy_vals, diff_coeffs, output_path):
    """表3 各指标熵权结果"""
    print("\n生成 表3 各指标熵权结果...")
    
    table_data = []
    for i, col in enumerate(ALL_INDICATORS):
        table_data.append({
            '指标名称': col,
            '信息熵 (E)': f'{entropy_vals[i]:.4f}',
            '差异系数 (G=1-E)': f'{diff_coeffs[i]:.4f}',
            '熵权权重 (W)': f'{weights[i]:.4f}',
            '权重排序': ''
        })
    
    # 添加排序
    df = pd.DataFrame(table_data)
    df['权重数值'] = weights
    df = df.sort_values('权重数值', ascending=False).reset_index(drop=True)
    df['权重排序'] = range(1, len(df) + 1)
    df = df.drop('权重数值', axis=1)
    
    df.to_excel(output_path, index=False)
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式
    md_lines = [
        "## 表3 各指标熵权结果\n\n",
        "| 排名 | 指标名称 | 信息熵 (E) | 差异系数 (G=1-E) | 熵权权重 (W) |\n",
        "|------|---------|-----------|----------------|-------------|\n"
    ]
    
    for _, row in df.iterrows():
        md_lines.append(f"| {row['权重排序']} | {row['指标名称']} | {row['信息熵 (E)']} | {row['差异系数 (G=1-E)']} | {row['熵权权重 (W)']} |\n")
    
    return ''.join(md_lines)


def generate_table4(df_master, composite_scores, econ_scores, env_scores, output_path):
    """表4 全国及四大区域年均得分"""
    print("\n生成 表4 全国及四大区域年均得分...")
    
    df = df_master.copy()
    df['综合得分'] = composite_scores
    df['经济子系统得分'] = econ_scores
    df['环境子系统得分'] = env_scores
    
    # 全国年均
    national_avg = df.groupby('年份')[['综合得分', '经济子系统得分', '环境子系统得分']].mean()
    
    # 区域年均
    region_avg = df.groupby(['年份', '区域'])[['综合得分', '经济子系统得分', '环境子系统得分']].mean()
    
    # Excel输出
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 全国年均得分
        national_df = pd.DataFrame({
            '年份': national_avg.index,
            '综合得分': national_avg['综合得分'].values,
            '经济子系统得分': national_avg['经济子系统得分'].values,
            '环境子系统得分': national_avg['环境子系统得分'].values,
        })
        national_df.to_excel(writer, sheet_name='全国年均得分', index=False)
        
        # 区域年均得分
        for region in ['东部', '中部', '西部', '东北']:
            region_data = region_avg.xs(region, level='区域')
            region_df = pd.DataFrame({
                '年份': region_data.index,
                f'{region}综合得分': region_data['综合得分'].values,
                f'{region}经济子系统': region_data['经济子系统得分'].values,
                f'{region}环境子系统': region_data['环境子系统得分'].values,
            })
            region_df.to_excel(writer, sheet_name=f'{region}地区年均', index=False)
    
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式 - 全国
    md_lines = ["## 表4 全国及四大区域年均得分\n\n", "### 4.1 全国年均得分\n\n"]
    md_lines.append("| 年份 | 综合得分 | 经济子系统得分 | 环境子系统得分 |\n")
    md_lines.append("|------|---------|---------------|---------------|\n")
    for year, row in national_avg.iterrows():
        md_lines.append(f"| {year} | {row['综合得分']:.4f} | {row['经济子系统得分']:.4f} | {row['环境子系统得分']:.4f} |\n")
    
    # 区域年均
    md_lines.append("\n### 4.2 四大区域年均综合得分\n\n")
    md_lines.append("| 区域 | 2018年 | 2019年 | 2020年 | 2021年 | 2022年 | 总体均值 |\n")
    md_lines.append("|------|--------|--------|--------|--------|--------|--------|\n")
    
    for region in ['东部', '中部', '西部', '东北']:
        region_data = region_avg.xs(region, level='区域')['综合得分']
        scores = [region_data.get(y, 0) for y in range(2018, 2023)]
        md_lines.append(f"| {region} | {scores[0]:.4f} | {scores[1]:.4f} | {scores[2]:.4f} | {scores[3]:.4f} | {scores[4]:.4f} | {np.mean(scores):.4f} |\n")
    
    return ''.join(md_lines)


def generate_table5(df_master, composite_scores, output_path, top_n=10):
    """表5 省份分年综合得分（节选）"""
    print("\n生成 表5 省份分年综合得分（节选）...")
    
    df = df_master.copy()
    df['综合得分'] = composite_scores
    
    # 选择代表性省份：按年均得分排名选前10
    province_avg = df.groupby('省份')['综合得分'].mean().sort_values(ascending=False)
    selected_provinces = province_avg.head(top_n).index.tolist()
    
    # 按得分排序
    df_selected = df[df['省份'].isin(selected_provinces)].copy()
    
    # Excel输出
    pivot = df_selected.pivot(index='省份', columns='年份', values='综合得分')
    # 按年均得分排序
    pivot = pivot.reindex(selected_provinces)
    pivot.to_excel(output_path)
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式
    md_lines = ["## 表5 省份分年综合得分（节选TOP10）\n\n"]
    md_lines.append("| 排名 | 省份 | 2018年 | 2019年 | 2020年 | 2021年 | 2022年 | 年均得分 |\n")
    md_lines.append("|------|------|--------|--------|--------|--------|--------|--------|\n")
    
    for rank, province in enumerate(selected_provinces, 1):
        scores = []
        for year in range(2018, 2023):
            mask = (df['省份'] == province) & (df['年份'] == year)
            score = df[mask]['综合得分'].values[0] if mask.sum() > 0 else 0
            scores.append(score)
        
        avg_score = np.mean(scores)
        score_strs = ' | '.join([f'{s:.4f}' for s in scores])
        md_lines.append(f"| {rank} | {province} | {score_strs} | {avg_score:.4f} |\n")
    
    return ''.join(md_lines)


def generate_table6(df_master, econ_scores, env_scores, output_path):
    """表6 四大子系统省域结果"""
    print("\n生成 表6 四大子系统省域结果...")
    
    df = df_master.copy()
    df['经济子系统得分'] = econ_scores
    df['环境子系统得分'] = env_scores
    
    # 按省份和年份计算平均
    province_avg = df.groupby('省份')[['经济子系统得分', '环境子系统得分']].mean()
    
    # 增加综合得分和排名
    province_avg = province_avg.sort_values('经济子系统得分', ascending=False)
    province_avg['经济子系统排名'] = range(1, len(province_avg) + 1)
    province_avg = province_avg.sort_values('环境子系统得分', ascending=False)
    province_avg['环境子系统排名'] = range(1, len(province_avg) + 1)
    province_avg = province_avg.sort_index()
    
    # Excel输出
    province_avg.to_excel(output_path)
    print(f"  ✓ 已保存: {output_path}")
    
    # Markdown格式
    md_lines = ["## 表6 四大子系统省域结果（年均值）\n\n"]
    md_lines.append("| 省份 | 经济子系统得分 | 经济排名 | 环境子系统得分 | 环境排名 |\n")
    md_lines.append("|------|---------------|---------|---------------|--------|\n")
    
    # 按经济子系统排序
    econ_sorted = province_avg.sort_values('经济子系统得分', ascending=False)
    for province in econ_sorted.index:
        row = econ_sorted.loc[province]
        md_lines.append(f"| {province} | {row['经济子系统得分']:.4f} | {row['经济子系统排名']} | ")
        md_lines.append(f"{province_avg.loc[province, '环境子系统得分']:.4f} | {province_avg.loc[province, '环境子系统排名']} |\n")
    
    return ''.join(md_lines)


# ============================================================
# 4. 主程序
# ============================================================

def main():
    print("=" * 60)
    print("规范化表格生成工具")
    print("=" * 60)
    
    # 加载主数据
    print("\n加载综合分析主表...")
    master_file = '综合数据分析/02_面板数据/综合分析主表.xlsx'
    df_master = pd.read_excel(master_file)
    print(f"  数据形状: {df_master.shape}")
    print(f"  年份: {sorted(df_master['年份'].unique())}")
    print(f"  省份: {df_master['省份'].nunique()}个")
    print(f"  区域: {sorted(df_master['区域'].unique())}")
    
    # 运行熵权TOPSIS分析
    results = run_full_analysis(df_master)
    weights = results['weights']
    entropy_vals = results['entropy']
    diff_coeffs = results['diff_coeffs']
    composite_scores = results['composite_scores']
    econ_scores = results['econ_scores']
    env_scores = results['env_scores']
    
    # 保存结果到综合数据分析目录
    output_dir = '综合数据分析'
    
    # 生成表1
    md_table1 = generate_table1(
        weights,
        f'{output_dir}/04_权重计算/表1_指标体系及权重.xlsx'
    )
    
    # 生成表3
    md_table3 = generate_table3(
        weights, entropy_vals, diff_coeffs,
        f'{output_dir}/04_权重计算/表3_各指标熵权结果.xlsx'
    )
    
    # 生成表4
    md_table4 = generate_table4(
        df_master, composite_scores, econ_scores, env_scores,
        f'{output_dir}/05_TOPSIS得分/表4_全国及四大区域年均得分.xlsx'
    )
    
    # 生成表5
    md_table5 = generate_table5(
        df_master, composite_scores,
        f'{output_dir}/05_TOPSIS得分/表5_省份分年综合得分_节选.xlsx'
    )
    
    # 生成表6
    md_table6 = generate_table6(
        df_master, econ_scores, env_scores,
        f'{output_dir}/06_耦合协调度/表6_四大子系统省域结果.xlsx'
    )
    
    # 汇总报告
    print("\n" + "=" * 60)
    print("生成汇总Markdown报告...")
    print("=" * 60)
    
    report_path = f'{output_dir}/08_分析报告/熵权TOPSIS分析结果汇总.md'
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("# 熵权TOPSIS分析结果汇总\n\n")
        f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"**数据来源**: 综合分析主表 (31省份 × 5年 × 10指标)\n\n")
        f.write(f"**计算方法**: 熵权法-TOPSIS综合评价模型\n\n")
        f.write("---\n\n")
        
        f.write(md_table1)
        f.write("\n\n---\n\n")
        f.write(md_table3)
        f.write("\n\n---\n\n")
        f.write(md_table4)
        f.write("\n\n---\n\n")
        f.write(md_table5)
        f.write("\n\n---\n\n")
        f.write(md_table6)
        f.write("\n\n---\n\n")
        
        f.write("*注：所有表格数据均基于熵权TOPSIS方法计算得出，权重通过信息熵法客观赋权*\n")
    
    print(f"  ✓ 汇总报告已保存: {report_path}")
    
    # 输出结果摘要
    print("\n" + "=" * 60)
    print("表格生成完成！")
    print("=" * 60)
    print("\n输出文件列表:")
    print(f"  表1: {output_dir}/04_权重计算/表1_指标体系及权重.xlsx")
    print(f"  表3: {output_dir}/04_权重计算/表3_各指标熵权结果.xlsx")
    print(f"  表4: {output_dir}/05_TOPSIS得分/表4_全国及四大区域年均得分.xlsx")
    print(f"  表5: {output_dir}/05_TOPSIS得分/表5_省份分年综合得分_节选.xlsx")
    print(f"  表6: {output_dir}/06_耦合协调度/表6_四大子系统省域结果.xlsx")
    print(f"  汇总报告: {report_path}")
    
    # 打印权重排名
    print("\n" + "=" * 60)
    print("权重排名速览:")
    print("=" * 60)
    weight_df = pd.DataFrame({
        '指标': ALL_INDICATORS,
        '权重': weights
    }).sort_values('权重', ascending=False).reset_index(drop=True)
    
    for i, row in weight_df.iterrows():
        print(f"  第{i+1}名: {row['指标']} (权重: {row['权重']:.4f})")
    
    # 打印TOP5省份
    print("\n" + "=" * 60)
    print("TOP5省份速览 (按年均综合得分):")
    print("=" * 60)
    df_temp = df_master.copy()
    df_temp['综合得分'] = composite_scores
    top5 = df_temp.groupby('省份')['综合得分'].mean().sort_values(ascending=False).head(5)
    for rank, (province, score) in enumerate(top5.items(), 1):
        print(f"  第{rank}名: {province} (得分: {score:.4f})")


if __name__ == '__main__':
    main()
