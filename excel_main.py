# 作者：han1549302096
#简介：生成一个周日和夜班的轮值表

import calendar
import tkinter as tk
import datetime
import xlwings as xw
import os
import numpy as np
import ttkbootstrap as ttk








class Duty(object):

    def __init__(self):
        self.list_night_staff = None
        self.today = datetime.datetime.today()  # 今天
        self.year = self.today.year  # 今年
        self.month = self.today.month  # 本月
        self.calendar_current_month = calendar.monthcalendar(self.year, self.month)  # 本月日历列表
        if self.month == 1:
            self.previous_month = 12  # 上个月
            self.calendar_previous_month = calendar.monthcalendar(self.year - 1, 12)  # 上个月日历列表
        else:
            self.previous_month = self.month - 1  # 上个月
            self.calendar_previous_month = calendar.monthcalendar(self.year, self.month - 1)  # 上个月日历列表
        self.list_ordinary = []  # 创建假期列表
        self.global_count = 0
        self.loop = 0
        self.list_line2 = []  # 第二行列表
        self.list_line3 = []  # 第三行列表
        self.list_line4 = []
        self.first_time = 0
        self.list_ordinary_staff = []
        self.list_special_staff = []
        self.list_special = []
        self.list_holiday_remove = []
        self.phone = ''
        self.index = []
        self.windows = None
        self.global_count1 = 0
        self.text = ['']
        self.var_111 = None

    def method_custom(self, year, month):
        """自定义初始化"""
        self.today = datetime.datetime.today()  # 今天
        if bool(len(year)) is False:
            self.year = datetime.datetime.today().year  # 今年
        else:
            self.year = int(year)
        if bool(len(month)) is False:
            self.month = datetime.datetime.today().month
        else:
            self.month = int(month)
        self.calendar_current_month = calendar.monthcalendar(self.year, self.month)  # 本月日历列表
        if self.month == 1:
            self.previous_month = 12  # 上个月
            self.calendar_previous_month = calendar.monthcalendar(self.year - 1, self.previous_month)  # 上个月日历列表
        else:
            self.previous_month = self.month - 1  # 上个月
            self.calendar_previous_month = calendar.monthcalendar(self.year, self.previous_month)  # 上个月日历列表

    def column_to_name(self,column):
        """转换A1C1"""
        self.show('正在转换A1C1格式')
        str1 = ''
        column1 = column
        while column != 0:
            if column % 26 == 0:  # 说明是整数
                str1 += 'Z'
            else:
                if column1 % 26 == 0:
                    str1 += chr((column - 1) % 26 - 1 + 65)
                else:
                    str1 += chr(column % 26 - 1 + 65)
            column //= 26
            if column1 % 26 == 0 and column == 1:
                return str1[::-1]
        return str1[::-1]

    def method_list(self, str1):
        """值班数组"""
        try:
            list1 = str1.split(sep=' ')
            if list1[0] == '':
                list1.remove('')
            if bool(int(list1[0])) is True:
                list1 = list(map(int, list1))
            return list1
        except:
            list1 = str1.split(sep=' ')
            if list1[0] == '':
                list1.remove('')
            return list1

    def save(self,year, month, str_staff_daytime, loop, time,
             str_night_staff, var_jjr, var_remove, var_jjr_staff, var_phone):

        self.text.clear()
        self.show('正在保存')
        train_acc = [year, month, str_staff_daytime, loop, time,
                     str_night_staff, var_jjr, var_remove, var_jjr_staff, var_phone]
        if not os.path.exists("result_variable"):
            os.mkdir("result_variable")
        np.savez('result_variable/train_acc', train_acc)
        self.show('已保存')

    def load(self):

        train_acc = np.load('result_variable/train_acc.npz')
        print(428, 'train_acc', train_acc)
        list_load = train_acc['arr_0']
        print(430, "load.npy done")
        print(431, 'list_load', list_load)
        year = list_load[0]
        month = list_load[1]
        str_staff_daytime = list_load[2]
        loop = list_load[3]
        time = list_load[4]
        str_night_staff = list_load[5]
        var_jjr = list_load[6]
        var_remove = list_load[7]
        var_jjr_staff = list_load[8]
        var_phone = list_load[9]
        return year, month, str_staff_daytime, loop, time, str_night_staff, var_jjr, var_remove, var_jjr_staff, var_phone

    def method_count_duty_staff(self, aa, cmd):
        """人员计数，cmd=0为下个人,1为当前人"""
        self.show('正在获取人员')
        if cmd == 0:
            if len(aa) == self.global_count:
                self.global_count = 0
            bb = aa[self.global_count]
            self.global_count += 1
        else:
            bb = aa[self.global_count]
        return bb

    def method_list_ordinary(self):  # 周日列表
        """获取上月末和本月的周日"""
        self.show('正在获取周日')
        self.list_ordinary.clear()
        for i in range(len(self.calendar_previous_month)):  # 遍历上个月
            if self.calendar_previous_month[i][6] != 0 and self.calendar_previous_month[i][6] > 28:  # 如果周日不等于0且大于28
                self.list_ordinary.append(self.calendar_previous_month[i][6])  # 符合条件的加入list_holiday
        for i in range(len(self.calendar_current_month)):  # 遍历本月
            if self.calendar_current_month[i][6] != 0 and self.calendar_current_month[i][6] < 29:
                self.list_ordinary.append(self.calendar_current_month[i][6])  # 符合条件的也加入list_holiday

    def show(self, text):
        progressbarOne = ttk.Progressbar(self.windows, bootstyle="success-striped", length=590)
        progressbarOne.place(x=10, y=400)
        # 进度值最大值
        progressbarOne['maximum'] = 106
        # 进度值初始值
        if self.global_count1 < 106:
            progressbarOne['value'] = self.global_count1
        else:
            self.global_count1 = 0

        self.text.append(text)
        self.global_count1 = len(self.text)
        self.var_111 = tk.StringVar()
        self.var_111.set(text)
        step = ttk.Label(self.windows, textvariable=self.var_111, bootstyle="inverse-success",font=('微软雅黑', 8))
        step.place(x=70, y=363)
        print(text)
        self.windows.update()

    def method_write_excel(self, days):
        """写excel"""
        self.text.clear()
        self.show('打开工作簿')
        wb = xw.Book()  # 打开一个工作簿
        self.show('关闭提示和警告消息')
        wb.display_alerts = False
        self.show('关闭屏幕刷新')
        wb.screen_updating = False
        self.show('打开工作表')
        sheet = wb.sheets[0]
        self.show('建立日期行数组')
        self.list_line2 = ['日期']  # 第二行列表
        self.show('建立白班行数组')
        self.list_line3 = ['白班(8:00-17:00)']  # 第三行列表
        self.show('建立夜班行数组')
        self.list_line4 = []

        self.show('开始写入日期行')
        if days != 29 and days != 28:
            self.list_line2.extend(list(range(29, days + 1)))  # 上个月的天数添加到list_line2
        if days == 29:
            self.list_line2.extend(list(range(29, 30)))
        self.list_line2.extend(list(range(1, 29)))  # 这个月的天数添加到list_line2
        sheet.range('A2').value = self.list_line2  # list_line2写入excel
        self.show('日期行写入完成')

        self.show('开始写入白班行')
        self.method_list_daytime()
        sheet.range('A3').value = self.list_line3
        self.show('白班行写入完成')
        self.show('开始写入夜班行')
        self.method_list_night()
        sheet.range('A4').value = self.list_line4
        self.show('夜班行写入完成')
        self.show('开始进行格式化操作')
        # self.method_list_staff(self.str_staff, 1)
        range_row_A4A_ = 'A4:A' + str(len(self.list_night_staff) + 3)
        # range_col_A1_1 = 'A1:' + str(self.method_self.column_to_name(int(len(self.list_line2)))) + '1'
        range_col_A1AA1 = 'A1:' + str(self.column_to_name(int(self.list_line2.index(24)))) + '1'
        range_col_A_1A_1 = str(self.column_to_name(int(self.list_line2.index(25)))) + '1:' + str(
            self.column_to_name(int(self.list_line2.index(28, 16) + 1))) + '1'
        range_col_B9_9 = 'B' + str(len(self.list_night_staff) + 4) + ':' \
                         + str(self.column_to_name(int(self.list_line2.index(16)))) + str(len(self.list_night_staff) + 4)
        range_col__9_9 = str(self.column_to_name(int(self.list_line2.index(17)))) + str(
            len(self.list_night_staff) + 4) + ':' + str(self.column_to_name(int(self.list_line2.index(28, 16) + 1))) + str(
            len(self.list_night_staff) + 4)
        range_cell_A_ = 'A' + str(len(self.list_night_staff) + 4)
        range_cell_note = 'B' + str(len(self.list_night_staff) + 4)
        range_cell_call = str(self.column_to_name(int(self.list_line2.index(17)))) + str(len(self.list_night_staff) + 4)
        range_all = 'A1:' + str(self.column_to_name(int(len(self.list_line2)))) + str(len(self.list_night_staff) + 4)
        range_all_1 = 'A2:' + str(self.column_to_name(int(len(self.list_line2)))) + str(len(self.list_night_staff) + 4)
        sheet.range(range_row_A4A_).api.Merge()
        sheet.range(range_col_A1AA1).api.Merge()
        sheet.range(range_col_A_1A_1).api.Merge()
        sheet.range(range_col_B9_9).api.Merge()
        sheet.range(range_col__9_9).api.Merge()
        sheet.range(range_cell_A_).value = [['备注']]
        sheet.range(range_cell_note).value = [['1.如果突发故障，由当班班长根据日常问题处理手册线性处理，车间班长无法解决并且紧急的故障，'
                                               '请拨打值班夜班人员电话寻求协助。\n'
                                               '2.如果遇到值班人员空缺，后面的自动补空位']]
        sheet.range(range_cell_call).value = [[str(self.phone)]]
        sheet.range(range_col_A1AA1).value = [['膜组装四车间' + str(self.year) + '年' + str(
            self.month) + '月设备维修人员值班安排表']]
        sheet.range('A4').value = [['夜班值班人员']]
        sheet.range(range_col_A_1A_1).value = [
            ['(' + str(self.previous_month) + '月' + str(self.list_line2[1]) + '日-' +
             str(self.month) + '月' + str(self.list_line2[-1]) + '日)']]

        sheet.cells.api.WrapText = True
        # sheet.cells.api.EntireColumn.AutoFit = True
        sheet.cells.api.HorizontalAlignment = -4108
        sheet.cells.column_width = 4
        sheet.cells.row_height = 40
        sheet.range(range_cell_note).api.HorizontalAlignment = -4131
        sheet.range(range_cell_call).api.HorizontalAlignment = -4131
        sheet.range(range_col_A_1A_1).api.HorizontalAlignment = -4152
        sheet.range(range_col_A_1A_1).api.VerticalAlignment = -4107
        sheet.range(range_cell_A_).row_height = 120
        sheet.range('A3').row_height = 80
        sheet.range(range_col_A1AA1).api.Font.Size = 24
        sheet.range(range_col_A1AA1).api.Font.Bold = True
        sheet.range(range_cell_note).api.Font.Size = 18
        sheet.range(range_cell_note).api.Font.Bold = True
        sheet.range(range_cell_call).api.Font.Size = 18
        sheet.range(range_cell_call).api.Font.Bold = True
        sheet.range(range_col_A_1A_1).api.Font.Size = 10
        sheet.range(range_col_A_1A_1).api.Font.Bold = False
        for i in self.index:
            sheet.range(str(self.column_to_name(i + 2)) + '2:' + str(
                self.column_to_name(i + 2)) + '3').color = 255, 255, 0
        # list_all = [7,8,9,10,11,12]
        for i in [7, 8, 9, 10, 11, 12]:
            sheet.range(range_all_1).api.Borders(i).LineStyle = 1
        for i in [7, 8, 9, 10]:
            sheet.range(range_all).api.Borders(i).Weight = 3
        self.show('已完成格式化操作')
        self.show('写入完成！')

    def method_list_daytime(self):
        """白班值班数组"""
        self.method_list_ordinary()
        self.show('正在创建引索空列表')
        self.index = []
        self.show('正在获取本月天数')
        list_local_day = self.list_line2[1:]  # 获取天数
        self.show('正在生成白班列表背景数据')
        list_local_daytime = ['全体'] * (len(self.list_line2[1:]))
        self.show('正在清零全局计数器')
        self.global_count = 0
        self.show('正在移除不休假的日期')
        self.show('正在判断是否存在不休假的日期')
        if bool(self.list_holiday_remove) is True:
            self.show('正在整理数据')
            self.list_holiday_remove = list(map(int, self.list_holiday_remove))
            self.show('将不休假的日期从周日和特殊假期中移除')
            for i in self.list_holiday_remove:
                if i in self.list_ordinary:
                    self.list_ordinary.remove(i)
                if i in self.list_special:
                    self.list_special.remove(i)
        self.show('将周日中的特殊假期移除')
        for i in self.list_special:
            if i in self.list_ordinary:
                self.list_ordinary.remove(i)
        self.show('判断周日存在')
        if bool(self.list_ordinary) is True:  # 如果特殊假日不为空
            self.show('清空全局计数器')
            self.global_count = 0
            self.show('判断是否存在值班人员')
            if bool(self.list_ordinary_staff) is True:  # 如果人员不为空
                self.show('创建索引列表')
                list_local_ordinary_holiday_index = []  # 创建索引列表
                self.show('遍历周日的日期')
                for day in self.list_ordinary:
                    list_local_ordinary_holiday_index.append(list_local_day.index(day))
                self.show('正在整理数据')
                list_local_ordinary_holiday_index.sort()
                self.show('正在索引假期对应单元格的位置')
                self.index = list_local_ordinary_holiday_index[:]
                self.show('正在将周日值班人员填入对应的日期中')
                for index in list_local_ordinary_holiday_index:
                    list_local_daytime[index] = self.method_count_duty_staff(self.list_ordinary_staff, 0)

        self.show('判断特殊假期存在')
        if bool(self.list_special) is True:  # 如果特殊假日不为空
            self.show('清空全局计数器')
            self.global_count = 0
            self.show('判断人员存在')
            if bool(self.list_special_staff) is True:  # 如果人员不为空
                self.show('创建索引列表')
                list_local_special_holiday_index = []  # 创建索引列表
                self.show('遍历特殊假期')
                for day in self.list_special:
                    list_local_special_holiday_index.append(list_local_day.index(day))
                self.show('正在整理数据')
                list_local_special_holiday_index.sort()
                self.show('正在索引假期对应单元格的位置')
                self.index.extend(list_local_special_holiday_index)
                self.show('正在将特殊假期值班人员填入对应的日期中')
                for index in list_local_special_holiday_index:
                    list_local_daytime[index] = self.method_count_duty_staff(self.list_special_staff, 0)
        self.show('正在写入白班数组！')
        self.list_line3.extend(list_local_daytime)

    def method_list_night(self):
        """夜班值班二维数组"""
        self.show('开始计算夜班数组')
        self.show('清零全局计数器')
        self.global_count = 0  # 每次进入人数加一个计数器清零        #建立二维数组
        self.show('正在创建夜班二维数组')
        list_night = [['' for _ in range(len(self.list_line2))] for _ in range(int(len(self.list_night_staff)))]
        # pp = len(self.list_line2)    # 无用
        self.show('正在计算夜班人员个数')
        len1 = int(len(self.list_night_staff))  # 计算出人员个数
        if len1 == 0:
            return 0
        self.show('正在计算临时变量')
        len2 = int(len(self.list_line2)) - 1 - self.first_time - self.loop * (len1 - 1)
        len3 = len2 // (len1 * self.loop)
        len4 = len2 - len3 * self.loop * len1
        len5 = len4 % self.loop
        len6 = len4 // self.loop
        len7 = len2 % (len1 * self.loop)
        self.show('正在写入第一个人的首次值班')
        list_night[0][1:1 + self.first_time] = [self.method_count_duty_staff(self.list_night_staff,
                                                                             0)] * self.first_time

        for n in range(1, len1):
            remove1 = self.loop * (n - 1)
            start1 = 1 + self.first_time
            end1 = start1 + self.loop
            self.show('正在写入头部非常规循环')
            list_night[n][start1 + remove1:end1 + remove1] = [self.method_count_duty_staff(self.list_night_staff,
                                                                                           0)] * self.loop

        if len3 >= 0:
            for m in range(len3):
                for n in range(len1):
                    remove1 = self.loop * n + len1 * self.loop * m
                    start1 = 1 + self.first_time + self.loop * (len1 - 1)
                    end1 = start1 + self.loop
                    self.show('正在写入中间完整循环')
                    list_night[n][start1 + remove1:end1 + remove1] = [self.method_count_duty_staff(
                        self.list_night_staff, 0)] * self.loop

        if len7 != 0 and len4 >= self.loop:
            for n in range(len6):
                remove1 = self.loop * n
                start1 = 1 + self.first_time + self.loop * (len1 - 1) + len3 * self.loop * len1
                end1 = start1 + self.loop
                self.show('正在写入尾部不完成循环')
                list_night[n][start1 + remove1:end1 + remove1] = [self.method_count_duty_staff(self.list_night_staff,
                                                                                               0)] * self.loop

        if len4 > 0 and len5 > 0:
            self.show('正在写入最后一人的最后值班(负值写入)')
            list_night[len6][-len5:] = [self.method_count_duty_staff(self.list_night_staff, 0)] * len5
            # 负值写入，剩余多少长度都填
        self.show('正在将二维数组传送给self变量')
        self.list_line4.extend(list_night)
        # 赋值给self.list_line4

    def method_select_previous_month_days(self):
        """通过判断上个月的天数，来决定写excel要写多少天"""
        self.show('正在判断要写excel多少天')
        if calendar.monthrange(self.year, self.previous_month)[1] == 31:
            days = 31
            self.method_write_excel(days)
        elif calendar.monthrange(self.year, self.previous_month)[1] == 30:
            days = 30
            self.method_write_excel(days)
        elif calendar.monthrange(self.year, self.previous_month)[1] == 29:
            days = 29
            self.method_write_excel(days)
        elif calendar.monthrange(self.year, self.previous_month)[1] == 28:
            days = 28
            self.method_write_excel(days)

    def method_gui(self):
        """基于tkinter的GUI"""
        # save(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)

        data_load = list(self.load())
        year11 = data_load[0]
        month11 = data_load[1]
        str_staff_daytime = data_load[2]
        loop = data_load[3]
        frequency = data_load[4]
        str_night_staff = data_load[5]
        jjr = data_load[6]
        holiday_remove = data_load[7]
        jjr_staff = data_load[8]
        phone = data_load[9]

        self.windows = tk.Tk()
        self.windows.title('值班表生成工具')
        self.windows.geometry('600x420')
        self.windows.resizable(0, 0)
        note = ttk.Notebook()
        note.place(relx=0.02, rely=0.2, relwidth=0.95, relheight=0.6)
        # 第一页
        frame1 = tk.Frame()
        note.add(frame1, text='白班')
        ttk.Label(frame1, text='白班值班顺序:',
                  font=('微软雅黑', 12)).place(x=10, y=20)
        var_staff_daytime = tk.StringVar()
        var_staff_daytime.set(str_staff_daytime)
        ttk.Entry(frame1, textvariable=var_staff_daytime, width=30,
                  font=('微软雅黑', 12)).place(x=120, y=15)

        # 第二页
        frame2 = tk.Frame()
        note.add(frame2, text='夜班')
        tk.Label(frame2, text='夜班每人连续值班几天:',
                 font=('微软雅黑', 12)).place(x=10, y=20)
        tk.Label(frame2, text='夜班值班顺序:',
                 font=('微软雅黑', 12)).place(x=10, y=60)
        tk.Label(frame2, text='月初夜班排头需值班几天:',
                 font=('微软雅黑', 12)).place(x=10, y=100)

        var_loop = tk.StringVar()
        var_loop.set(loop)
        ttk.Entry(frame2, textvariable=var_loop, width=30,
                  font=('微软雅黑', 14)).place(x=200, y=20)

        var_staff_night = tk.StringVar()
        var_staff_night.set(str_night_staff)
        ttk.Entry(frame2, textvariable=var_staff_night, width=30,
                  font=('微软雅黑', 14)).place(x=200, y=60)

        var_time = tk.StringVar()
        var_time.set(frequency)
        ttk.Entry(frame2, textvariable=var_time, width=30,
                  font=('微软雅黑', 14)).place(x=200, y=100)

        # 第三页
        frame3 = tk.Frame()
        note.add(frame3, text='节假日')
        tk.Label(frame3, text='特殊节假日是哪几天:',
                 font=('微软雅黑', 12)).place(x=10, y=20)
        tk.Label(frame3, text='不休假是哪几天:',
                 font=('微软雅黑', 12)).place(x=10, y=60)
        tk.Label(frame3, text='特殊节假日值班顺序:',
                 font=('微软雅黑', 12)).place(x=10, y=100)
        var_jjr = tk.StringVar()
        var_jjr.set(jjr)
        ttk.Entry(frame3, textvariable=var_jjr, width=30,
                  font=('微软雅黑', 14)).place(x=200, y=20)

        var_remove = tk.StringVar()
        var_remove.set(holiday_remove)
        ttk.Entry(frame3, textvariable=var_remove, width=30,
                  font=('微软雅黑', 14)).place(x=200, y=60)
        var_jjr_staff = tk.StringVar()
        var_jjr_staff.set(jjr_staff)
        ttk.Entry(frame3, textvariable=var_jjr_staff, width=30,
                  font=('微软雅黑', 14)).place(x=200, y=100)

        # 第四页
        frame4 = tk.Frame()
        note.add(frame4, text='电话')
        tk.Label(frame4, text='值班电话:',
                 font=('微软雅黑', 12)).place(x=10, y=20)

        tk.Label(self.windows, text='年:',
                 font=('微软雅黑', 12)).place(x=420, y=30)
        tk.Label(self.windows, text='月:',
                 font=('微软雅黑', 12)).place(x=420, y=70)



        var_phone = tk.StringVar()
        var_phone.set(phone)
        ppp = ttk.Text(frame4, font=('微软雅黑', 14), height=7, width=40)
        ppp.insert('0.0', phone)
        ppp.place(x=100, y=20)

        var_nian = tk.StringVar()
        var_nian.set(year11)
        select1 = ttk.Combobox(self.windows, width=12, textvariable=var_nian, state="readonly")
        lis = []
        lis.extend(range(2018, 2099))
        select1['values'] = lis
        select1.place(x=450, y=30)

        var_yue = tk.StringVar()
        var_yue.set(month11)
        select2 = ttk.Combobox(self.windows, width=12, textvariable=var_yue, state="readonly")
        lis = []
        lis.extend(range(1, 13))
        select2['values'] = lis
        select2.place(x=450, y=70)

        d = ttk.LabelFrame(self.windows, height=100, width=50)
        d.place(relx=0.02, rely=0.81, relwidth=0.95, relheight=0.135)

        tk.Label(self.windows, text='运行状态:',
                 font=('微软雅黑', 8)).place(x=15, y=363)

        c = ttk.Button(self.windows, text='保存', bootstyle="info", width=10,
                       command=lambda: self.save(var_nian.get(), var_yue.get(), var_staff_daytime.get(), var_loop.get(),
                                            var_time.get(), var_staff_night.get(), var_jjr.get(), var_remove.get(),
                                            var_jjr_staff.get(), ppp.get("1.0", "end")))
        c.place(x=380, y=357)

        b = ttk.Button(self.windows, text='生成', bootstyle="success", width=10,
                       command=lambda: self.method_hit(
                           var_nian.get(), var_yue.get(), var_staff_daytime.get(),
                           var_loop.get(), var_time.get(), var_staff_night.get(),
                           var_jjr.get(), var_remove.get(), var_jjr_staff.get(), ppp.get("1.0", "end")))
        b.place(x=480, y=357)

        self.windows.mainloop()

    def method_hit(self, year, month, str_staff_daytime, loop, time1,
                   str_night_staff, var_jjr, var_remove, var_jjr_staff, var_phone):
        if loop == '':
            self.loop = 0
        else:
            self.loop = int(loop)
        if time1 == '':
            self.first_time = 0
        else:
            self.first_time = int(time1)
        self.list_special = self.method_list(var_jjr)[:]
        self.list_ordinary_staff = self.method_list(str_staff_daytime)[:]
        self.list_night_staff = self.method_list(str_night_staff)[:]
        self.list_special_staff = self.method_list(var_jjr_staff)[:]
        self.phone = var_phone
        print(408, 'year,month,str_staff_daytime,loop,time1,str_night_staff,var_jjr,var_remove',
              var_jjr_staff, year, month, str_staff_daytime, loop, time1, str_night_staff, var_jjr, var_remove,
              var_jjr_staff, var_phone)
        self.list_holiday_remove = self.method_list(var_remove)
        if bool(year) is not False or bool(month) is not False:
            self.method_custom(year, month)
        self.global_count1 = 0
        self.method_select_previous_month_days()

        self.var_111.set(self.text[-1])
        step = ttk.Label(self.windows, textvariable=self.var_111, bootstyle="inverse-success",font=('微软雅黑', 8))
        step.place(x=70, y=363)


start = Duty()
start.method_gui()
