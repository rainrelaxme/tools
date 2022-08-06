# -*- coding:utf-8 -*-
# authored by shawn
import json

from map.distance import distance_two, total

"""
# 计算两点之间的距离和高度差
pot1 = [31.25781982421875, 120.73540554470486, 4.3862183718010783]
pot2 = [31.257807617187499, 120.73540934244792, 4.9125050101429224]
new_distance = distance_two(pot1, pot2)
new_height = pot2[2] - pot1[2]
print(f"两点之间的距离是{new_distance},海拔差是{new_height}")
"""

# 计算一组点之间的距离和高度差
# import json
# filename = r'D:\project\python\little-tools\map\gps_pot.json'
filename = r'D:\project\python\little-tools\map\altitude.json'

# 读取json文件
with open(filename) as f_obj:
    pots = json.load(f_obj)
# print(pots)

i = 0
step = 100   # 每step取一个点
dist_sum = []  # 运动距离
hgt_sum_pos = []  # 爬升
hgt_sum_neg = []  # 下降
pre_pot = []  # 前一个点
current_pot = []  # 后一个点
"""
for item in pots:
    # 前后两个点相减
    pre_pot = current_pot
    current_pot = [item['latitude'], item['longitude'], item['alitiude']]
    if not pre_pot:
        pass
    else:
        dist_sum.append(distance_two(pre_pot, current_pot))
        height = current_pot[2] - pre_pot[2]
        if height <= 0:
            hgt_sum_neg.append(height)
        else:
            hgt_sum_pos.append(height)

distance_sum = total(dist_sum)
height_sum_pos = total(hgt_sum_pos)
height_sum_neg = total(hgt_sum_neg)
print(f"共{len(dist_sum)}个数据,距离总和是{distance_sum},共爬升{height_sum_pos},下降{height_sum_neg}")
print(f"\n距离差是：{dist_sum}")
print(f"\n爬升：{hgt_sum_pos}\n下降：{hgt_sum_neg}")
"""

# step决定取多少个点
for item in pots:
    # step前后减
    if i == 0 or i % step == 0:
        pre_pot = current_pot
        current_pot = [item['latitude'], item['longitude'], item['alitiude']]
        if not pre_pot:
            pass
        else:
            dist_sum.append(distance_two(pre_pot, current_pot))
            height = current_pot[2] - pre_pot[2]
            if height <= 0:
                hgt_sum_neg.append(height)
            else:
                hgt_sum_pos.append(height)
    else:
        pass
    i = i + 1


distance_sum = total(dist_sum)
height_sum_pos = total(hgt_sum_pos)
height_sum_neg = total(hgt_sum_neg)
print(f"共{len(dist_sum)}个数据,距离总和是{distance_sum},共爬升{height_sum_pos},下降{height_sum_neg}")
print(f"\n距离差是：{dist_sum}")
print(f"\n爬升：{hgt_sum_pos}\n下降：{hgt_sum_neg}")