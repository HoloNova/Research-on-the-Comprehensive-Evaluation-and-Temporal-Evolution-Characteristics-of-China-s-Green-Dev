"""
统一分析主数据构建脚本
自动合并可用指标面板数据、识别正负向指标、添加区域分类、标准化处理
生成完整可用于后续所有建模的主表
"""

import os
import pandas as pd
import numpy as np
import json
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# ============================================================
# 1. 省份-区域分类标准 (国家统计局标准)
# ============================================================
PROVINCE_REGION_MAP = {
    # 东部地区
    '北京': '东部', '天津': '东部', '河北': '东部', '上海': '东部',
    '江苏': '东部', '浙江': '东部', '福建': '东部', '山东': '东部',
    '广东': '东部', '海南': '东部',
    # 中部地区
    '山西': '中部', '安徽': '中部', '江西': '中部', '河南': '中部',
    '湖北': '中部', '湖南': '中部',
    # 西部地区
    '内蒙古': '西部', '广西': '西部', '重庆': '西部', '四川': '西部',
    '贵州': '西部', '云南': '西部', '西藏': '西部', '陕西': '西部',
    '甘肃': '西部', '青海': '西部', '宁夏': '西部', '新疆': '西部',
    # 东北地区
    '辽宁': '东北', '吉林': '东北', '黑龙江': '东北'
}

# ============================================================
# 2. 指标元数据定义
# ============================================================
INDICATOR_METADATA = {
    # 经济子系统
    '人均GDP': {
        'direction': 'positive',
        'subsystem': '经济',
        'unit': '元',
        'description': '人均国内生产总值'
    },
    '第三产业占比': {
        'direction': 'positive',
        'subsystem': '经济',
        'unit': '%',
        'description': '第三产业增加值占GDP比重'
    },
    'R&D占比': {
        'direction': 'positive',
        'subsystem': '经济',
        'unit': '%',
        'description': '研发经费投入占GDP比重'
    },
    # 环境子系统
    '单位GDP能耗': {
        'direction': 'negative',
        'subsystem': '环境',
        'unit': '吨标准煤/万元',
        'description': '单位GDP能源消耗'
    },
    '单位GDP水耗': {
        'direction': 'negative',
        'subsystem': '环境',
        'unit': '立方米/万元',
        'description': '单位GDP水资源消耗'
    },
    '污水处理率': {
        'direction': 'positive',
        'subsystem': '环境',
        'unit': '%',
        'description': '城市污水处理率'
    },
    '垃圾处理率': {
        'direction': 'positive',
        'subsystem': '环境',
        'unit': '%',
        'description': '城市生活垃圾无害化处理率'
    },
    '森林覆盖率': {
        'direction': 'positive',
        'subsystem': '环境',
        'unit': '%',
        'description': '森林面积占国土面积比重'
    },
    '优良天数比例': {
        'direction': 'positive',
        'subsystem': '环境',
        'unit': '%',
        'description': '空气质量优良天数比例'
    },
    '环保投资占比': {
        'direction': 'positive',
        'subsystem': '环境',
        'unit': '%',
        'description': '环保投资占GDP比重'
    }
}

# ============================================================
# 3. 核心函数
# ============================================================

def load_province_data(filepath='sample_data.xlsx'):
    """加载省份面板数据"""
    print(f"\n[步骤1] 加载省份面板数据: {filepath}")
    df = pd.read_excel(filepath)
    print(f"  - 数据形状: {df.shape}")
    print(f"  - 年份范围: {df['年份'].min()}-{df['年份'].max()}")
    print(f"  - 省份数量: {df['省份'].nunique()}")
    print(f"  - 指标列: {[c for c in df.columns if c not in ['年份', '省份']]}")
    return df


def add_region_info(df):
    """添加区域分类信息"""
    print("\n[步骤2] 添加区域分类信息")
    df['区域'] = df['省份'].map(PROVINCE_REGION_MAP)
    
    # 统计区域分布
    region_counts = df['区域'].value_counts()
    for region, count in region_counts.items():
        provinces = df[df['区域'] == region]['省份'].unique()
        print(f"  - {region}地区: {count}条记录, 省份: {', '.join(provinces)}")
    
    # 检查缺失
    missing = df[df['区域'].isna()]
    if len(missing) > 0:
        print(f"  ⚠ 警告: {len(missing)}条记录缺少区域分类")
        print(f"    缺失省份: {missing['省份'].unique()}")
    else:
        print("  ✓ 所有省份区域分类完整")
    
    return df


def identify_indicator_direction(df):
    """识别并验证指标方向"""
    print("\n[步骤3] 识别正负向指标")
    indicator_cols = [c for c in df.columns if c not in ['年份', '省份', '区域']]
    
    positive_indicators = []
    negative_indicators = []
    unknown_indicators = []
    
    for col in indicator_cols:
        if col in INDICATOR_METADATA:
            direction = INDICATOR_METADATA[col]['direction']
            if direction == 'positive':
                positive_indicators.append(col)
                print(f"  [+] 正向指标: {col} ({INDICATOR_METADATA[col]['description']})")
            else:
                negative_indicators.append(col)
                print(f"  [-] 负向指标: {col} ({INDICATOR_METADATA[col]['description']})")
        else:
            unknown_indicators.append(col)
            print(f"  [?] 未知指标: {col}")
    
    print(f"\n  汇总: 正向指标 {len(positive_indicators)}个, "
          f"负向指标 {len(negative_indicators)}个, "
          f"未知指标 {len(unknown_indicators)}个")
    
    return positive_indicators, negative_indicators, unknown_indicators


def check_data_quality(df, indicator_cols):
    """数据质量检查"""
    print("\n[步骤4] 数据质量检查")
    
    # 缺失值检查
    missing = df[indicator_cols].isnull().sum()
    total_missing = missing.sum()
    if total_missing > 0:
        print(f"  ⚠ 发现 {total_missing} 个缺失值:")
        for col, count in missing[missing > 0].items():
            print(f"    - {col}: {count}个缺失")
    else:
        print("  ✓ 无缺失值")
    
    # 异常值检查 (3σ原则)
    print("\n  异常值检测 (3σ原则):")
    outlier_count = 0
    for col in indicator_cols:
        mean = df[col].mean()
        std = df[col].std()
        lower = mean - 3 * std
        upper = mean + 3 * std
        outliers = df[(df[col] < lower) | (df[col] > upper)]
        if len(outliers) > 0:
            print(f"    - {col}: {len(outliers)}个异常值")
            outlier_count += len(outliers)
    
    if outlier_count == 0:
        print("    ✓ 无异常值")
    
    # 统计描述
    print("\n  数据概览:")
    print(f"    - 总记录数: {len(df)}")
    print(f"    - 时间跨度: {df['年份'].min()}-{df['年份'].max()} ({df['年份'].nunique()}年)")
    print(f"    - 省份数量: {df['省份'].nunique()}")
    print(f"    - 区域数量: {df['区域'].nunique()}")
    print(f"    - 指标数量: {len(indicator_cols)}")
    
    return total_missing, outlier_count


def normalize_data(df, positive_cols, negative_cols, method='minmax'):
    """标准化处理"""
    print(f"\n[步骤5] 标准化处理 (方法: {method})")
    
    df_norm = df.copy()
    all_indicators = positive_cols + negative_cols
    
    for col in all_indicators:
        min_val = df[col].min()
        max_val = df[col].max()
        range_val = max_val - min_val
        
        if range_val == 0:
            print(f"  ⚠ {col}: 取值范围为0, 跳过标准化")
            df_norm[f'{col}_标准化'] = 0
            continue
        
        if col in positive_cols:
            # 正向指标: 值越大越好
            df_norm[f'{col}_标准化'] = (df[col] - min_val) / range_val
        else:
            # 负向指标: 值越小越好
            df_norm[f'{col}_标准化'] = (max_val - df[col]) / range_val
        
        print(f"  ✓ {col}: 范围[{min_val:.2f}, {max_val:.2f}] -> [0, 1]")
    
    return df_norm


def build_master_table(df, positive_cols, negative_cols):
    """构建主表"""
    print("\n[步骤6] 构建分析主表")
    
    # 标准化数据
    df_master = normalize_data(df, positive_cols, negative_cols)
    
    # 整理列顺序
    base_cols = ['年份', '省份', '区域']
    original_cols = positive_cols + negative_cols
    normalized_cols = [f'{col}_标准化' for col in original_cols]
    
    master_cols = base_cols + original_cols + normalized_cols
    df_master = df_master[master_cols]
    
    print(f"\n  主表结构:")
    print(f"    - 基础信息列: {base_cols}")
    print(f"    - 原始指标列 ({len(original_cols)}个): {original_cols}")
    print(f"    - 标准化列 ({len(normalized_cols)}个): {normalized_cols}")
    print(f"    - 总列数: {len(df_master.columns)}")
    print(f"    - 总行数: {len(df_master)}")
    
    return df_master


def save_master_data(df_master, output_dir='综合数据分析'):
    """保存主数据到指定目录"""
    print(f"\n[步骤7] 保存主数据到 {output_dir}")
    
    # 1. 保存主表 Excel
    master_file = os.path.join(output_dir, '02_面板数据', '综合分析主表.xlsx')
    df_master.to_excel(master_file, index=False)
    print(f"  ✓ 主表已保存: {master_file}")
    
    # 2. 保存原始数据副本
    base_cols = ['年份', '省份', '区域']
    indicator_cols = [c for c in df_master.columns if c not in base_cols and not c.endswith('_标准化')]
    df_original = df_master[base_cols + indicator_cols]
    original_file = os.path.join(output_dir, '01_原始数据', '省份面板原始数据.xlsx')
    df_original.to_excel(original_file, index=False)
    print(f"  ✓ 原始数据已保存: {original_file}")
    
    # 3. 保存标准化数据
    norm_cols = [c for c in df_master.columns if c.endswith('_标准化')]
    df_normalized = df_master[base_cols + norm_cols]
    norm_file = os.path.join(output_dir, '03_标准化结果', '省份面板标准化数据.xlsx')
    df_normalized.to_excel(norm_file, index=False)
    print(f"  ✓ 标准化数据已保存: {norm_file}")
    
    # 4. 生成指标元数据 JSON
    metadata = {
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '数据概况': {
            '总记录数': len(df_master),
            '年份范围': f"{df_master['年份'].min()}-{df_master['年份'].max()}",
            '省份数量': df_master['省份'].nunique(),
            '区域数量': df_master['区域'].nunique(),
            '指标数量': len(indicator_cols)
        },
        '指标信息': {},
        '区域分类': {}
    }
    
    for col in indicator_cols:
        if col in INDICATOR_METADATA:
            meta = INDICATOR_METADATA[col]
            metadata['指标信息'][col] = {
                '方向': '正向' if meta['direction'] == 'positive' else '负向',
                '子系统': meta['subsystem'],
                '单位': meta['unit'],
                '描述': meta['description']
            }
    
    # 区域分类
    region_provinces = {}
    for province, region in PROVINCE_REGION_MAP.items():
        if region not in region_provinces:
            region_provinces[region] = []
        region_provinces[region].append(province)
    metadata['区域分类'] = region_provinces
    
    metadata_file = os.path.join(output_dir, '01_原始数据', '指标元数据.json')
    with open(metadata_file, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)
    print(f"  ✓ 指标元数据已保存: {metadata_file}")
    
    # 5. 生成数据字典
    dict_file = os.path.join(output_dir, '01_原始数据', '数据字典.xlsx')
    dict_data = []
    for col in df_master.columns:
        dict_data.append({
            '字段名': col,
            '数据类型': str(df_master[col].dtype),
            '描述': col.replace('_标准化', '的标准化值') if col.endswith('_标准化') else 
                   INDICATOR_METADATA.get(col, {}).get('description', col)
        })
    pd.DataFrame(dict_data).to_excel(dict_file, index=False)
    print(f"  ✓ 数据字典已保存: {dict_file}")
    
    return {
        'master_file': master_file,
        'original_file': original_file,
        'normalized_file': norm_file,
        'metadata_file': metadata_file,
        'dict_file': dict_file
    }


def generate_summary_report(df_master, positive_cols, negative_cols, output_dir='综合数据分析'):
    """生成数据概览报告"""
    print("\n[步骤8] 生成数据概览报告")
    
    report_file = os.path.join(output_dir, '08_分析报告', '主数据概览报告.md')
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("# 综合分析主数据概览报告\n\n")
        f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # 1. 数据概况
        f.write("## 1. 数据概况\n\n")
        f.write(f"| 项目 | 数值 |\n")
        f.write(f"|------|------|\n")
        f.write(f"| 总记录数 | {len(df_master)} |\n")
        f.write(f"| 时间范围 | {df_master['年份'].min()}-{df_master['年份'].max()}年 |\n")
        f.write(f"| 年份数量 | {df_master['年份'].nunique()} |\n")
        f.write(f"| 省份数量 | {df_master['省份'].nunique()} |\n")
        f.write(f"| 区域数量 | {df_master['区域'].nunique()} |\n")
        f.write(f"| 指标数量 | {len(positive_cols) + len(negative_cols)} |\n\n")
        
        # 2. 区域分布
        f.write("## 2. 区域分布\n\n")
        region_stats = df_master.groupby('区域')['省份'].nunique()
        f.write("| 区域 | 省份数量 | 省份列表 |\n")
        f.write("|------|---------|---------|\n")
        for region, count in region_stats.items():
            provinces = df_master[df_master['区域'] == region]['省份'].unique()
            f.write(f"| {region} | {count} | {', '.join(provinces)} |\n")
        f.write("\n")
        
        # 3. 指标体系
        f.write("## 3. 指标体系\n\n")
        f.write("### 3.1 正向指标\n\n")
        f.write("| 指标名称 | 单位 | 描述 | 最小值 | 最大值 | 均值 | 标准差 |\n")
        f.write("|---------|------|------|--------|--------|------|-------|\n")
        for col in positive_cols:
            if col in INDICATOR_METADATA:
                meta = INDICATOR_METADATA[col]
                f.write(f"| {col} | {meta['unit']} | {meta['description']} | ")
                f.write(f"{df_master[col].min():.2f} | {df_master[col].max():.2f} | ")
                f.write(f"{df_master[col].mean():.2f} | {df_master[col].std():.2f} |\n")
        
        f.write("\n### 3.2 负向指标\n\n")
        f.write("| 指标名称 | 单位 | 描述 | 最小值 | 最大值 | 均值 | 标准差 |\n")
        f.write("|---------|------|------|--------|--------|------|-------|\n")
        for col in negative_cols:
            if col in INDICATOR_METADATA:
                meta = INDICATOR_METADATA[col]
                f.write(f"| {col} | {meta['unit']} | {meta['description']} | ")
                f.write(f"{df_master[col].min():.2f} | {df_master[col].max():.2f} | ")
                f.write(f"{df_master[col].mean():.2f} | {df_master[col].std():.2f} |\n")
        
        # 4. 标准化结果概览
        f.write("\n## 4. 标准化结果概览\n\n")
        f.write("### 4.1 正向指标标准化统计\n\n")
        f.write("| 指标 | 最小值 | 最大值 | 均值 | 标准差 |\n")
        f.write("|------|--------|--------|------|-------|\n")
        for col in positive_cols:
            norm_col = f'{col}_标准化'
            f.write(f"| {col} | {df_master[norm_col].min():.4f} | {df_master[norm_col].max():.4f} | ")
            f.write(f"{df_master[norm_col].mean():.4f} | {df_master[norm_col].std():.4f} |\n")
        
        f.write("\n### 4.2 负向指标标准化统计\n\n")
        f.write("| 指标 | 最小值 | 最大值 | 均值 | 标准差 |\n")
        f.write("|------|--------|--------|------|-------|\n")
        for col in negative_cols:
            norm_col = f'{col}_标准化'
            f.write(f"| {col} | {df_master[norm_col].min():.4f} | {df_master[norm_col].max():.4f} | ")
            f.write(f"{df_master[norm_col].mean():.4f} | {df_master[norm_col].std():.4f} |\n")
        
        # 5. 时间趋势概览
        f.write("\n## 5. 时间趋势概览\n\n")
        numeric_cols = df_master.select_dtypes(include=[np.number]).columns.tolist()
        yearly_avg = df_master.groupby('年份')[numeric_cols].mean()
        f.write("### 5.1 各年度指标均值\n\n")
        f.write("| 年份 | 人均GDP | 第三产业占比 | R&D占比 | 单位GDP能耗 | 污水处理率 | 森林覆盖率 |\n")
        f.write("|------|---------|-------------|---------|-------------|-----------|-----------|\n")
        for year, row in yearly_avg.iterrows():
            f.write(f"| {year} | {row['人均GDP']:.0f} | {row['第三产业占比']:.1f}% | "
                    f"{row['R&D占比']:.2f}% | {row['单位GDP能耗']:.2f} | "
                    f"{row['污水处理率']:.1f}% | {row['森林覆盖率']:.1f}% |\n")
        
        # 6. 区域差异概览
        f.write("\n## 6. 区域差异概览\n\n")
        region_avg = df_master.groupby('区域')[numeric_cols].mean()
        f.write("### 6.1 各区域指标均值\n\n")
        f.write("| 区域 | 人均GDP | 第三产业占比 | R&D占比 | 单位GDP能耗 | 污水处理率 | 森林覆盖率 |\n")
        f.write("|------|---------|-------------|---------|-------------|-----------|-----------|\n")
        for region, row in region_avg.iterrows():
            f.write(f"| {region} | {row['人均GDP']:.0f} | {row['第三产业占比']:.1f}% | "
                    f"{row['R&D占比']:.2f}% | {row['单位GDP能耗']:.2f} | "
                    f"{row['污水处理率']:.1f}% | {row['森林覆盖率']:.1f}% |\n")
        
        f.write("\n---\n\n")
        f.write("*本报告由自动化工具生成，数据来源于统一分析主数据*\n")
    
    print(f"  ✓ 数据概览报告已保存: {report_file}")
    return report_file


def main():
    print("="*80)
    print("统一分析主数据构建工具")
    print("="*80)
    
    # 检查源数据
    source_file = 'sample_data.xlsx'
    if not os.path.exists(source_file):
        print(f"\n❌ 错误: 源数据文件 '{source_file}' 不存在!")
        return
    
    # 步骤1: 加载数据
    df = load_province_data(source_file)
    
    # 步骤2: 添加区域信息
    df = add_region_info(df)
    
    # 步骤3: 识别指标方向
    positive_cols, negative_cols, unknown_cols = identify_indicator_direction(df)
    
    if unknown_cols:
        print(f"\n⚠ 警告: 发现 {len(unknown_cols)} 个未定义方向的指标: {unknown_cols}")
        print("  请在 INDICATOR_METADATA 中补充定义")
    
    all_indicator_cols = positive_cols + negative_cols
    
    # 步骤4: 数据质量检查
    missing_count, outlier_count = check_data_quality(df, all_indicator_cols)
    
    # 步骤5-6: 构建主表
    df_master = build_master_table(df, positive_cols, negative_cols)
    
    # 步骤7: 保存主数据
    output_dir = '综合数据分析'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    saved_files = save_master_data(df_master, output_dir)
    
    # 步骤8: 生成概览报告
    report_file = generate_summary_report(df_master, positive_cols, negative_cols, output_dir)
    
    # 完成总结
    print("\n" + "="*80)
    print("主数据构建完成!")
    print("="*80)
    print(f"\n输出文件列表:")
    for key, path in saved_files.items():
        print(f"  - {key}: {path}")
    print(f"  - 报告: {report_file}")
    
    print(f"\n主数据摘要:")
    print(f"  - 数据形状: {df_master.shape}")
    print(f"  - 正向指标: {len(positive_cols)}个")
    print(f"  - 负向指标: {len(negative_cols)}个")
    print(f"  - 区域分类: 东/中/西/东北 四大区域")
    print(f"  - 标准化方法: 极差标准化 (Min-Max)")
    
    print("\n" + "="*80)
    print("主数据已就绪，可用于后续建模分析!")
    print("="*80)
    
    return df_master


if __name__ == '__main__':
    df_master = main()