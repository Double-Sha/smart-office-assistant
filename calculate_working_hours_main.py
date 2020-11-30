#!/usr/bin/env python
# -*- coding: utf-8 -*-

# @Team : PRA-BD
# @Author: shasha.mao
# @Date : 11/30/2020 1:20 PM
# @File : calculate_working_hours_main.py
# @Tool : PyCharm

from office_utils.calculate_working_hours import CalculateWorkinggHours

temp = CalculateWorkinggHours(r"./resources/&nbsp&nbsp&nbsp&nbsp打卡查看.xls")
print(['%s:%s' % item for item in temp.__dict__.items()])
temp.calculation(excle_save_path=r"./resources/")  # 计算结果保存至excel
temp.plot(figure_save_path=r"./resources/")  # 计算结果可视化