import matplotlib.pyplot as plt
import numpy as np
import os

# 确保目录存在
os.makedirs('01-industry/images', exist_ok=True)

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'PingFang SC', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

# 煤炭需求来源数据 (基于EIA数据和行业分析)
labels = ['电力生产', '钢铁工业', '化学工业', '其他']
sizes = [62, 18, 12, 8]  # 百分比
colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']

# 创建饼图 - 调整布局
fig, ax = plt.subplots(figsize=(10, 8))

# 绘制饼图，调整标签位置
wedges, texts, autotexts = ax.pie(sizes, 
                                  labels=labels, 
                                  colors=colors,
                                  autopct='%1.1f%%', 
                                  startangle=90,
                                  textprops={'fontsize': 12},
                                  pctdistance=0.8,
                                  labeldistance=1.1)

# 设置标题
ax.set_title('中国煤炭需求分布 (2023年)', 
             fontsize=16, fontweight='bold', pad=20)

# 美化百分比文字
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontsize(11)
    autotext.set_weight('bold')

# 美化标签文字
for text in texts:
    text.set_fontsize(12)
    text.set_weight('normal')

# 确保图表居中
ax.axis('equal')

# 添加数据来源注释
plt.figtext(0.5, 0.02, '数据来源：美国能源信息署(EIA)，行业分析', 
            ha='center', fontsize=9, style='italic', color='gray')

# 保存图表
plt.tight_layout()
plt.savefig('01-industry/images/coal_demand_sources_cn.png', dpi=300, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.show()

print("中文图表已保存到: 01-industry/images/coal_demand_sources_cn.png")