# encoding:utf-8
import glob
import xlrd
from xlutils.copy import copy
import xlsxwriter


class yf_time_struct(object):
    year = 0
    month = 0
    day = 0
    hour = 0
    minute = 0
    sec = 0

    def trup2str(self, trup):
        s = ''
        for tmp in trup:
            s += '%02d:' % tmp
        return s

    def show(self, view=False, rt_str=True):
        out = (2000 + self.year, self.month, self.day)
        out1 = (self.hour + 8, self.minute, self.sec)
        if view:
            pass
        if rt_str:
            day = self.trup2str(out).strip(":").replace(":", "-")
            time = self.trup2str(out1).strip(":")
            dat_time = day + " " + time
            return dat_time
        else:
            return out


class utc_time(object):
    monthtable = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    monthtable_leap = (31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

    def __init__(self):
        self.time = yf_time_struct()
        self.MIN_YDAR = 14
        self.MAX_YEAR = 29
        self.year = 2000

    def modify_utc_start_time(self, min_year=14, jan1week=3, start_year=2000):
        self.MIN_YDAR = min_year
        self.JAN1WEEK = jan1week
        self.year = start_year

    def isleap(self, year):
        year += self.year
        return (((year % 400) == 0) or (((year % 4) == 0) and not ((year % 100) == 0)))

    def seconds_to_utc(self, seconds):
        sec = seconds % 60
        minute = seconds / 60
        hour = minute / 60
        day = hour / 24

        self.time.sec = sec
        self.time.minute = minute % 60
        self.time.hour = hour % 24

        year = self.MIN_YDAR
        while (1):
            leap = self.isleap(year)
            if day < (365 + leap):
                break
            day -= 365 + leap
            year += 1

        self.time.year = year % 100

        mtbl = self.monthtable_leap if leap > 0 else self.monthtable

        for month in range(12):
            if day < mtbl[month]:
                break
            day -= mtbl[month]

        self.time.day = day + 1
        self.time.month = month + 1

        return self.time


def rever_bytes(str_buf):
    assert (len(str_buf) % 2 == 0)
    out = ''
    for i in range(0, len(str_buf), 2):
        out = str_buf[i:i + 2] + out
    return out


def get_real_ord(num):
    num = int(num, base=16)
    num_len = len(hex(num).replace('0x', '').replace('L', ''))
    num_bytes = num_len / 2 if num_len % 2 == 0 else num_len / 2 + 1
    num = num if num & (0x1 << (num_bytes * 8 - 1)) == 0 else -((0x1 << num_bytes * 8) - 1 - num + 1)
    return num


def get_time(time, file_name):
    file = open(file_name, "r")
    lines = file.readlines()
    for line in lines:
        j = 0
        if "FE7F" in line:
            line2 = line.split("remove")[0]
            file1 = open(u"原始数据%s.txt" % time, "w")
            while j < len(line2):
                file1.write(line2[j:j + 12])
                file1.write("\n")
                j += 12
            file1.close()
            break


def get_data(time):
    t = utc_time()
    file = open(u"原始数据%s.txt" % time, "r")
    lines = file.readlines()
    file2 = open(u"解析数据%s.txt" % time, "w")
    file2.write("ax   ay   az\n")
    for line in lines:
        if line[0:4] == "FE7F":
            tt = rever_bytes(line[4:12])
            time1 = t.seconds_to_utc(eval("0x" + tt)).show(True)
            file2.write("time %s\n" % time1)
        elif line[0:8] == "CDABBADC":
            spm = int(rever_bytes(line[8:12]), base=16)
            file2.write("spm: *** %s ***\n" % spm)
        elif line[0:8] == "34122143":
            heart = int(rever_bytes(line[8:12]), base=16)
            file2.write("heart: *** %s ***\n" % heart)
        else:
            x1 = rever_bytes(line[0:4])
            x = get_real_ord(x1)

            y1 = rever_bytes(line[4:8])
            y = get_real_ord(y1)

            z1 = rever_bytes(line[8:12])
            z = get_real_ord(z1)

            file2.write(str(x) + "  " + str(y) + "  " + str(z) + "\n")


def handle_data(time):
    file = glob.glob(u"解析数据*.txt")
    k = 0
    list1 = []
    time_list = 0
    for i in range(len(file)):
        file1 = open(file[i], "r")
        list = file1.readlines()

        for i in range(len(list)):
            if "time" in list[i]:
                k += 1
                list1.append(i)
        list1.append(len(list) - 1)
        s = 0
        while k > time_list:
            workbook = xlsxwriter.Workbook('data/%s_%s.xls' % (time, time_list))
            workbook.close()

            o_fd = open('data/%s_%s.txt' % (time, time_list), 'w')
            for i in range(len(list) - (list1[s]) - 1):
                o_fd.write(list[i + (list1[s])])
                if i + (list1[s]) + 1 == list1[s + 1]:
                    break

            o_fd.close()
            s += 1
            time_list += 1

    file1 = glob.glob("data/%s*.txt" % time)
    for i in range(len(file1)):
        m = 1
        file2 = open(file1[i], "r")
        lines = file2.readlines()
        excel = xlrd.open_workbook("data/%s_%s.xls" % (time, i), "r")
        wb = copy(excel)
        ws = wb.get_sheet(0)
        # ws.write(0, 0, u"心率")
        ws.write(0, 0, u"步频")
        for line in lines:
            if "spm" in line:
                spm = int(line.split("*** ")[1].split(" ***")[0])
                ws.write(m, 0, spm)
                m += 1
        wb.save("data/%s_%s.xls" % (time, i))


def chart(time):
    file1 = glob.glob("data/*.xls")

    for i in range(len(file1)):
        j = 0
        workbook = xlsxwriter.Workbook('%s_chart_%s.xls' % (time, i))  # 创建一个excel文件
        data = xlrd.open_workbook('data/%s_%s.xls' % (time, i), 'rb')  # 打开fname文件
        name = "sheet1"
        table = data.sheet_by_index(0)  # 通过索引获取xls文件第0个sheet
        nrows = table.nrows  # 获取table工作表总行数
        ncols = table.ncols  # 获取table工作表总列数
        worksheet = workbook.add_worksheet("%s" % name)  # 创建一个工作表对象
        worksheet.set_column(0, ncols, 10)  # 设定列的宽度为22像素
        worksheet.set_column(0, 0, 20)


        chart = workbook.add_chart({'type': 'line'})
        chart.set_y_axis({'name': u'步频'})
        chart.set_x_axis({'name': u'时间'})
        chart.set_title({'name': u'步频数据'})
        chart.set_size({'width': 600, 'height': 400})
        for i in range(nrows):

            worksheet.set_row(i, 15)  # 设定第i行单元格属性，高度为22像素，行索引从0开始
            worksheet.write(j, 0, table.cell_value(i, 0))
            j += 1
        chart.add_series({'name': u'步频',
                          #'categories': '=%s!$A$2:$A$%s' % (name, j),
                          'values': '=%s!$A$2:$A$%s' % (name, j)})

        worksheet.insert_chart('D1', chart, {'x_offset': 200, 'y_offset': 100})
        workbook.close()


def run():
    try:
        file = glob.glob("*.log")
        for i in range(len(file)):
            time = file[i].split("RTT_Terminal_")[1].split(".")[0]
            get_time(time, file[i])
            get_data(time)
            handle_data(time)
            chart(time)
    except Exception as e:
        print(e)


if __name__ == '__main__':
    run()

