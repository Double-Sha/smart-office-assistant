# -*- coding: utf-8 -*-
"""
@author: shasha.mao
version 1.0 by shasha.mao 2019.06.15 - first version
version 2.0 by shasha.mao 2019.06.25 - add the feature of odd clocking in
version 3.0 by shasha.mao 2020.11.30 - 代码重构
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import xlrd
from datetime import datetime


class CalculateWorkinggHours:
    """读取单个月份的员工打卡工时记录，得出该月的每日工作时长，保存至excel并进行可视化"""

    def __init__(self, raw_excel_path):
        """读取原始工时记录表格，确定统计时间信息，员工信息"""

        # 确定数据体
        self.data = xlrd.open_workbook(raw_excel_path).sheet_by_index(0)  # 读取第一个sheet
        # 确定年数，月份，工号，姓名，科室等信息
        time_info = datetime.strptime(self.data.col_values(3, start_rowx=3, end_rowx=4)[0], "%Y-%m-%d %H:%M")
        self.year = time_info.year
        self.month = time_info.month
        self.id = self.data.col_values(0, start_rowx=3, end_rowx=4)[0]
        self.name = self.data.col_values(1, start_rowx=3, end_rowx=4)[0]
        self.dept = self.data.col_values(2, start_rowx=3, end_rowx=4)[0]


    def calculation(self, excle_save_path=r"../resources/"):
        """计算每日工作时长，统计奇数次打卡行为"""

        MAX_WROK_DAYS = 31
        workingHours = [0]
        working_date_list = ['']
        timeCols = self.data.col_values(3, start_rowx=1)  # 取出时间这一列

        # ------------split the time to days and time in days------------
        dayWithYears = ['']  # 2019-06-01
        days = ['']  # 06-01
        timeInDays = ['']  # 09:30:00
        oddTimes = ['odd times of clocking in']  # odd times of clocking in
        dayWithYears[0], timeInDays[0] = timeCols[0].split()
        i = 1
        while i < len(timeCols):
            dayWithYears.append((timeCols[i].split())[0])
            timeInDays.append((timeCols[i].split())[1])
            i += 1

        days[0] = (dayWithYears[0].split("-", 1))[1]
        i = 1
        while i < len(timeInDays):
            days.append((dayWithYears[i].split("-", 1))[1])
            i += 1

        # ------------Calculate the working hours------------
        timePoints = [datetime.now()]
        timePoints[0] = datetime.strptime(timeCols[0], "%Y-%m-%d %H:%M")
        for i in np.arange(1, len(days), 1):
            timePoints.append(datetime.strptime(timeCols[i], "%Y-%m-%d %H:%M"))

        j = 0
        k = 0
        while j < len(days) - 1:
            if j != 0:
                if days[j] == days[j + 1]:
                    workingHours.append((timePoints[j] - timePoints[j + 1]).seconds / 3600)
                    working_date_list.append(days[j])
                    j += 2
                else:
                    oddTimes.append(days[j])
                    j += 1
            else:
                if days[j] == days[j + 1]:
                    workingHours[k] = (timePoints[j] - timePoints[j + 1]).seconds / 3600
                    working_date_list[k] = days[j]
                    j += 2
                else:
                    oddTimes.append(days[j])
                    j += 1
        if working_date_list[0] == '' and workingHours[0] == 0:
            del working_date_list[0]
            del workingHours[0]

        i = 0
        while i < len(working_date_list) - 1:
            if working_date_list[i] == working_date_list[i + 1]:
                del working_date_list[i]
                workingHours[i + 1] += (workingHours[i])
                del workingHours[i]
            i += 1

        for x in oddTimes:
            # remove the working with the odd times
            if x in working_date_list:
                a = working_date_list.index(x)
                del working_date_list[a]
                del workingHours[a]
        self.odd_times = oddTimes

        # 将每日工作时长计算结果以excel格式进行保存
        self.working_hours_df = pd.DataFrame({
            "working_date": np.array(working_date_list),
            "working_hours_per_day": np.array(workingHours)
        })
        path = excle_save_path + "%s_%s_%s-%s_每日工作时长.xlsx" % (self.id, self.name, self.year, self.month)  # 拼接保存文件名
        self.working_hours_df.to_excel(path, index=0)
        print("%s_%s_%s-%s_每日工作时长计算结果已保存至excel，请在resources文件夹下查看" % (self.id, self.name, self.year, self.month))


    def plot(self, figure_save_path=r"../resources/"):
        """对统计结果进行可视化"""

        # 统计指标计算
        totalWorkDays = len(self.working_hours_df.working_date)
        maxHours = self.working_hours_df.working_hours_per_day.max()
        maxIndex = self.working_hours_df.working_hours_per_day.idxmax()
        minHours = self.working_hours_df.working_hours_per_day.min()
        minIndex = self.working_hours_df.working_hours_per_day.idxmin()
        avgHours = self.working_hours_df.working_hours_per_day.mean()
        total_working_hours = self.working_hours_df.working_hours_per_day.sum()

        i = 0
        normalWorkCounter = 0
        mediumWorkCounter = 0
        highWorkCounter = 0
        while i < len(self.working_hours_df.working_hours_per_day):
            if self.working_hours_df.working_hours_per_day[i] <= 9:
                normalWorkCounter += 1
            elif (self.working_hours_df.working_hours_per_day[i] > 9) & (self.working_hours_df.working_hours_per_day[i] <= 11):
                mediumWorkCounter += 1
            else:
                highWorkCounter += 1
            i += 1
        overtimeTimes = [normalWorkCounter, mediumWorkCounter, highWorkCounter]
        overtimeDegree = ['normal', 'medium', 'high']

        plt.style.use('ggplot')
        # # ggplot fivethirtyeight
        fig = plt.figure()
        fig.set_size_inches(18.5, 10.5)

        # fig1 工作强度饼状图统计
        ax1 = fig.add_subplot(2, 2, 1)
        colors = ['yellowgreen', 'lightskyblue', 'lightcoral']
        ax1.pie(x=overtimeTimes, labels=overtimeDegree, autopct='%0.1f%%', colors=colors,
                explode=(0.1, 0, 0.1), pctdistance=0.6, labeldistance=1.1)
        plt.axis('equal')
        plt.legend(["normal:≤9h", "medium:9h-11h", "high:>11h"], loc='best')
        plt.title("Fig.1 Overtime Degree", size=18)

        # fig2  统计信息描述
        ax2 = fig.add_subplot(2, 2, 2)
        ax2.plot([0, 1], [0, 1], alpha=0)
        plt.grid()
        ax2.set_xticks([])
        ax2.set_yticks([])
        plt.title("Fig.2 Overview", size=18)
        plt.text(0, 0.9, 'Dept : %s      ID : %s      %s-%s' % (self.dept, self.id, self.year, self.month), size=12, color='black',
                 alpha=0.9)
        plt.text(0, 0.75, 'Total %d working days, %.1f hours without the days with odd clocking in' % (
        totalWorkDays, total_working_hours),
                 size=12, color='black', alpha=0.9)
        plt.text(0, 0.6, 'Average working hours per day : %.1f' % avgHours, size=12, color='black', alpha=1)
        plt.text(0, 0.45, 'Maximum working hours is %.1f hours in %s ' % (maxHours, self.working_hours_df.working_date.iloc[maxIndex]),
                 size=12, color='black', alpha=0.9)
        plt.text(0, 0.3, 'Minimum working hours is %.1f hours in %s  ' % (minHours, self.working_hours_df.working_date.iloc[minIndex]),
                 size=12, color='black', alpha=0.9)

        plt.text(0.0, 0.15, 'There are odd times of clocking in in below days:', size=10, color='black', alpha=1.0)

        j = 0
        del self.odd_times[0]
        for i in self.odd_times:
            plt.text(0.6 + j, 0.15, "(%s)" % i, color='r')
            j += 0.1
        plt.text(0, 0, 'Please feel free to contact @shasha.mao with any questions you may have', size=10,
                 color='coral', alpha=1.0)

        # fig3 该月每日工作时长可视化
        ax3 = fig.add_subplot(2, 1, 2)
        ax3.plot_date(self.working_hours_df.working_date, self.working_hours_df.working_hours_per_day, 'r', linewidth=2, marker='x', markersize=9, markeredgewidth=3)
        ax3.xaxis.set_tick_params(rotation=45)
        ax3.set_yticks([minHours, maxHours + 2], 'b')
        plt.xlabel('Date', color="k", size=16)
        plt.ylabel('Working Hours Per Day', color="k", size=16)
        plt.title("Fig.3 Working Hours Record", size=18)

        plt.tick_params(labelsize=14, colors="k")
        for a, b in zip(self.working_hours_df.working_date, self.working_hours_df.working_hours_per_day):
            plt.text(a, b + 1, round(b, 1), ha='center', va='center', fontsize=14)

        path = figure_save_path + "%s_%s_%s-%s_工作时长统计.png" % (self.id, self.name, self.year, self.month)  # 拼接保存文件名
        plt.savefig(path, dpi=400)
        print("%s_%s_%s-%s_本月工作时长统计可视化完成，请在resources文件夹下查看" % (self.id, self.name, self.year, self.month))



def main():
    """简述类的用法"""
    temp = CalculateWorkinggHours(r"../resources/&nbsp&nbsp&nbsp&nbsp打卡查看.xls")

    print(['%s:%s' % item for item in temp.__dict__.items()])
    temp.calculation()
    temp.plot()



if __name__ == "__main__":

    main()
