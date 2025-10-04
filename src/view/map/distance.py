# 通过gps坐标计算两点之间距离
# -*- coding:utf-8 -*-
import math


def distance_two(first_pot, second_pot):
    # 纬差1度的距离是 111712.69150641055729984301412873米，经差1度的距离是102834.74258026089786013677476285米
    distance = math.sqrt((first_pot[0] - second_pot[0]) * (first_pot[0] - second_pot[0]) * 111712.7 * 111712.7 +
                         (first_pot[1] - second_pot[1]) * (first_pot[1] - second_pot[1]) * 102834.7 * 102834.7)
    # 保留3位小数
    distance = format(distance, '.3f')
    return distance


def total(my_list):
    # 计算总和
    sum = 0
    for i_list in my_list:
        sum = sum + float(i_list)
    return sum




