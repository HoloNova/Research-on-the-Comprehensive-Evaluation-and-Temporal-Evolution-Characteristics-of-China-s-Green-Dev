"""
测试Matplotlib中文字体设置
"""

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 测试字体设置
print("=== Matplotlib字体测试 ===")
print("当前font.family:", plt.rcParams['font.family'])
print("当前font.sans-serif:", plt.rcParams['font.sans-serif'])
print("axes.unicode_minus:", plt.rcParams['axes.unicode_minus'])

# 查找系统中的中文字体
print("\n=== 系统中文字体 ===")
found = False
for font in fm.findSystemFonts():
    font_name = fm.FontProperties(fname=font).get_name()
    if any(chinese_char in font_name for chinese_char in ['SimHei', '黑体', 'YaHei', '微软雅黑', 'Song', '宋体']):
        print(f"找到字体: {font_name} - {font}")
        found = True

if not found:
    print("未找到中文字体")

# 测试绘制中文
print("\n=== 测试中文绘制 ===")
try:
    plt.figure(figsize=(8, 4))
    plt.plot([1, 2, 3], [4, 5, 6])
    plt.title('测试中文标题')
    plt.xlabel('横坐标')
    plt.ylabel('纵坐标')
    plt.savefig('test_chinese_font.png', dpi=100)
    print("✓ 中文绘制测试成功")
except Exception as e:
    print(f"✗ 中文绘制测试失败: {e}")
finally:
    plt.close()

print("\n字体测试完成！")
