"""
修复图表字体乱码问题
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

def get_chinese_font():
    """获取中文字体"""
    # 查找系统中的中文字体
    chinese_fonts = []
    for font in fm.findSystemFonts():
        try:
            font_name = fm.FontProperties(fname=font).get_name()
            if any(name in font.lower() for name in ['simhei', 'yahei', '黑体', '微软雅黑']):
                chinese_fonts.append(font)
        except:
            pass
    
    if chinese_fonts:
        print(f"找到中文字体: {chinese_fonts[0]}")
        return fm.FontProperties(fname=chinese_fonts[0])
    else:
        print("未找到中文字体，使用默认字体")
        return None

def regenerate_charts():
    """重新生成所有图表"""
    print("开始修复图表字体...")
    
    # 获取中文字体
    chinese_font = get_chinese_font()
    
    # 读取数据
    results = pd.read_excel('analysis_results/分析结果.xlsx')
    
    # 创建输出目录
    output_dir = '绿色发展研究结果'
    vis_dir = os.path.join(output_dir, '可视化图表')
    if not os.path.exists(vis_dir):
        os.makedirs(vis_dir)
    
    print(f"数据形状: {results.shape}")
    print(f"输出目录: {vis_dir}")
    
    # 图表1: 时间演化特征总图
    print("\n生成时间演化特征总图...")
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    
    numeric_cols = results.select_dtypes(include=[np.number]).columns
    yearly_avg = results.groupby('年份')[numeric_cols].mean()
    
    axes[0, 0].plot(yearly_avg.index, yearly_avg['综合得分'], 
                     marker='o', linewidth=2, markersize=8)
    axes[0, 0].set_title('绿色发展综合得分时间趋势', fontsize=14, fontweight='bold', fontproperties=chinese_font)
    axes[0, 0].set_xlabel('年份', fontproperties=chinese_font)
    axes[0, 0].set_ylabel('得分', fontproperties=chinese_font)
    axes[0, 0].grid(True, alpha=0.3)
    
    axes[0, 1].plot(yearly_avg.index, yearly_avg['经济子系统得分'], 
                     marker='s', linewidth=2, markersize=8, label='经济子系统', color='blue')
    axes[0, 1].plot(yearly_avg.index, yearly_avg['环境子系统得分'], 
                     marker='^', linewidth=2, markersize=8, label='环境子系统', color='green')
    axes[0, 1].set_title('子系统得分时间趋势', fontsize=14, fontweight='bold', fontproperties=chinese_font)
    axes[0, 1].set_xlabel('年份', fontproperties=chinese_font)
    axes[0, 1].set_ylabel('得分', fontproperties=chinese_font)
    if chinese_font:
        axes[0, 1].legend(prop=chinese_font)
    else:
        axes[0, 1].legend()
    axes[0, 1].grid(True, alpha=0.3)
    
    axes[1, 0].plot(yearly_avg.index, yearly_avg['耦合度'], 
                     marker='D', linewidth=2, markersize=8, label='耦合度', color='orange')
    axes[1, 0].plot(yearly_avg.index, yearly_avg['耦合协调度'], 
                     marker='*', linewidth=2, markersize=10, label='耦合协调度', color='red')
    axes[1, 0].set_title('耦合协调度时间趋势', fontsize=14, fontweight='bold', fontproperties=chinese_font)
    axes[1, 0].set_xlabel('年份', fontproperties=chinese_font)
    axes[1, 0].set_ylabel('协调度', fontproperties=chinese_font)
    if chinese_font:
        axes[1, 0].legend(prop=chinese_font)
    else:
        axes[1, 0].legend()
    axes[1, 0].grid(True, alpha=0.3)
    
    # 协调等级分布饼图
    level_counts = results['协调等级'].value_counts()
    wedges, texts, autotexts = axes[1, 1].pie(level_counts.values, labels=level_counts.index, autopct='%1.1f%%', startangle=90)
    axes[1, 1].set_title('协调等级分布', fontsize=14, fontweight='bold', fontproperties=chinese_font)
    if chinese_font:
        for text in texts + autotexts:
            text.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '时间演化特征总图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("✓ 时间演化特征总图生成完成")
    
    # 图表2: 省份得分热力图
    print("\n生成省份得分热力图...")
    pivot_df = results.pivot(index='省份', columns='年份', values='综合得分')
    fig, ax = plt.subplots(figsize=(16, 12))
    sns.heatmap(pivot_df, annot=True, fmt='.4f', cmap='RdYlGn', center=0.5, ax=ax)
    ax.set_title('各省份绿色发展综合得分热力图', fontsize=14, fontweight='bold', fontproperties=chinese_font)
    ax.set_xlabel('年份', fontproperties=chinese_font)
    ax.set_ylabel('省份', fontproperties=chinese_font)
    if chinese_font:
        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontproperties(chinese_font)
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '省份得分热力图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("✓ 省份得分热力图生成完成")
    
    # 图表3: 年度分布箱线图
    print("\n生成年度分布箱线图...")
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    sns.boxplot(x='年份', y='综合得分', data=results, ax=axes[0, 0])
    axes[0, 0].set_title('综合得分年度分布', fontsize=12, fontweight='bold', fontproperties=chinese_font)
    axes[0, 0].set_xlabel('年份', fontproperties=chinese_font)
    axes[0, 0].set_ylabel('综合得分', fontproperties=chinese_font)
    
    sns.boxplot(x='年份', y='经济子系统得分', data=results, ax=axes[0, 1])
    axes[0, 1].set_title('经济子系统得分年度分布', fontsize=12, fontweight='bold', fontproperties=chinese_font)
    axes[0, 1].set_xlabel('年份', fontproperties=chinese_font)
    axes[0, 1].set_ylabel('经济子系统得分', fontproperties=chinese_font)
    
    sns.boxplot(x='年份', y='环境子系统得分', data=results, ax=axes[1, 0])
    axes[1, 0].set_title('环境子系统得分年度分布', fontsize=12, fontweight='bold', fontproperties=chinese_font)
    axes[1, 0].set_xlabel('年份', fontproperties=chinese_font)
    axes[1, 0].set_ylabel('环境子系统得分', fontproperties=chinese_font)
    
    sns.boxplot(x='年份', y='耦合协调度', data=results, ax=axes[1, 1])
    axes[1, 1].set_title('耦合协调度年度分布', fontsize=12, fontweight='bold', fontproperties=chinese_font)
    axes[1, 1].set_xlabel('年份', fontproperties=chinese_font)
    axes[1, 1].set_ylabel('耦合协调度', fontproperties=chinese_font)
    
    # 设置所有子图的字体
    if chinese_font:
        for ax_row in axes:
            for ax in ax_row:
                for label in ax.get_xticklabels() + ax.get_yticklabels():
                    label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '年度分布箱线图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("✓ 年度分布箱线图生成完成")
    
    # 图表4: 经济环境关系散点图
    print("\n生成经济环境关系散点图...")
    fig, ax = plt.subplots(figsize=(10, 8))
    scatter = ax.scatter(results['经济子系统得分'], results['环境子系统得分'], 
                         c=results['年份'], cmap='viridis', alpha=0.6, s=100)
    cbar = plt.colorbar(scatter, label='年份')
    ax.set_xlabel('经济子系统得分', fontsize=12, fontproperties=chinese_font)
    ax.set_ylabel('环境子系统得分', fontsize=12, fontproperties=chinese_font)
    ax.set_title('经济与环境子系统关系散点图', fontsize=14, fontweight='bold', fontproperties=chinese_font)
    ax.grid(True, alpha=0.3)
    
    # 设置字体
    if chinese_font:
        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontproperties(chinese_font)
        cbar.ax.set_ylabel('年份', fontproperties=chinese_font)
        for label in cbar.ax.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    plt.savefig(os.path.join(vis_dir, '经济环境关系散点图.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("✓ 经济环境关系散点图生成完成")
    
    print("\n所有图表已重新生成，字体问题已修复！")

if __name__ == '__main__':
    regenerate_charts()
