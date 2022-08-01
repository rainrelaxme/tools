# 通过gps坐标计算两点之间距离
# -*- coding:utf-8 -*-
import json
import math


def distance_two(first_pot, second_pot):
    # 纬差1度的距离是 111712.69150641055729984301412873米，经差1度的距离是102834.74258026089786013677476285米
    distance = math.sqrt((first_pot[0] - second_pot[0]) * (first_pot[0] - second_pot[0]) * 111712.7 * 111712.7 +
                         (first_pot[1] - second_pot[1]) * (first_pot[1] - second_pot[1]) * 102834.7 * 102834.7)
    # 保留3位小数
    distance = format(distance, '.3f')
    return distance


# import json

filename = r'D:\project\python\little-tools\map\gps_pot.json'

# 读取json文件
with open(filename) as f_obj:
    pots = json.load(f_obj)

dist = []   # 运动距离
pre_pot = []    # 前一个点
current_pot = []    # 后一个点

for item in pots:
    # print(f"纬度：{item['latitude']}, 经度：{item['longitude']}")
    # 前后两个点相减
    pre_pot = current_pot
    current_pot = [item['latitude'], item['longitude']]
    # if item == 0:
    #     current_pot = [item['latitude'], item['longitude']]
    #     pre_pot = []
    #     break
    # else:
    #     pre_pot = current_pot
    #     current_pot = [item['latitude'], item['longitude']]
    if not pre_pot:
        pass
    else:
        dist.append(distance_two(pre_pot, current_pot))


def total(my_list):
    # 计算总和
    sum = 0
    for i_list in my_list:
        sum = sum + float(i_list)
    return sum


distance_sum = total(dist)
print(f"共{len(dist)}个数据：\n{dist}\n总和是{distance_sum}")

