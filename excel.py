import xlrd
import numpy as np
import matplotlib.mlab as mlab
import matplotlib.pyplot as plt

data = None


class DataSet:

    def __init__(self, filename):
        try:
            self.origin = xlrd.open_workbook(filename)
            print("data load success!")
        except Exception as e:
            print(str(e))

        self.obj_list = None

        # 攻击
        self.weapon_labels = ['Biological', 'Chemical', 'Radiological', 'Nuclear',
                         'Firearms', 'Explosives', 'Fake Weapons', 'Incendiary',
                         'Melee', 'Vehicle', 'Sabotage Equipment', 'Other', 'Unknown']

    def get_obj_list(self):
        if self.obj_list is not None:
            return self.obj_list

        table = self.origin.sheets()[0]
        nrows = table.nrows
        ncols = table.ncols
        colnames = table.row_values(0)
        self.obj_list = []
        for ri in range(1, nrows):
            row = table.row_values(ri)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                self.obj_list.append(app)
        return self.obj_list


# 打开EXCEL文件
def open_excel(file):
    global data
    if data is not None:
        return data
    try:
        data = xlrd.open_workbook(file)
        print("data load success!")
        return data
    except Exception as e:
        print(str(e))


obj_list = None


# 数据预处理
# 将excel转化为对象list
def preprocessing(file):
    global obj_list
    if obj_list is not None:
        return obj_list

    data = open_excel(file)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    colnames = table.row_values(0)
    obj_list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            obj_list.append(app)
    return obj_list


def test(file):
    obj_list = preprocessing(file)
    i = 1
    for row in obj_list:
        if i == 100:
            break
        print(row)
        i += 1


# 分析中国十年间发生的受害目标为运输时间的数量，死亡人数，受伤人数
# 中国，country 44
# 十年间，iyear 2007-2017
# 受害目标 运输，targtype1 19
# 事件数量
# 死亡人数 nkill
# 受伤人数 nwound
def china_ten_tran(file):
    obj_list = preprocessing(file)
    event_list = []
    wound_count = 0
    kill_count = 0
    for row in obj_list:
        if row["country"] == 44 and 2007 <= row["iyear"] < 2017 and row['targtype1'] == 19:
            event_list.append(row)
            if row["nwound"] is not None and row["nwound"] != 0:
                wound_count += row["nwound"]

            if row["nkill"] is not None and row["nkill"] != 0:
                kill_count += row["nkill"]

    print("受伤人数： " + str(wound_count))
    print("直接死亡人数： " + str(kill_count))
    event_count = len(event_list)
    print("事件总数： " + str(event_count))
    return event_count, wound_count, kill_count


# 攻击方式
# 饼状图
def china_ten_tran_weapon_pie(file):
    obj_list = preprocessing(file)
    table = data.sheets()[0]

    # 攻击
    weapon_labels = ['Biological', 'Chemical', 'Radiological', 'Nuclear',
                     'Firearms', 'Explosives', 'Fake Weapons', 'Incendiary',
                     'Melee', 'Vehicle', 'Sabotage Equipment', 'Other', 'Unknown']
    weapon_counts = [0 for i in range(0, len(weapon_labels))]

    # 攻击方式
    event_list = []
    for row in obj_list:
        if row["country"] == 44 and 2007 <= row["iyear"] < 2017 and row['targtype1'] == 19:
            event_list.append(row)
            if row['weaptype1'] is not None and row['weaptype1'] > 0:
                weapon_counts[int(row['weaptype1']) - 1] += 1

    labels = []
    counts = []
    for i in range(len(weapon_labels)):
        if weapon_counts[i] != 0:
            labels.append(weapon_labels[i])
            counts.append(weapon_counts[i])

    plt.figure()
    plt.pie(counts, labels=labels, autopct='%1.2f%%')
    plt.title("Pie chart")

    plt.show()
    plt.savefig("PieChart.jpg")


# 条件过滤器
class DataFilter:
    def __init__(self):
        self._obj_filter_list = []

    def filter(self, row):
        for f in self._obj_filter_list:
            if f(row) is not True:
                return False
            return True

    def add_filter(self, f):
        self._obj_filter_list.append(f)

    def empty_filter(self):
        self._obj_filter_list = []


# 中国
# 受害目标：巴士 targsubtype1 99
def china_bus(ds):
    obj_list = ds.get_obj_list()
    df = DataFilter()
    df.add_filter(lambda x: x["country"] == 44)
    df.add_filter(lambda x: x['targtype1'] == 19)
    df.add_filter(lambda x: x["targsubtype1"] == 99)
    bus_event_list = []
    for row in obj_list:
        if df.filter(row):
            bus_event_list.append(row)
    print("中国巴士事件数量: " + str(len(bus_event_list)))

    df.empty_filter()
    df.add_filter(lambda x: x["country"] == 44)
    df.add_filter(lambda x: x['targtype1'] == 19)
    tran_event_list = []
    for row in obj_list:
        if df.filter(row):
            tran_event_list.append(row)
    print("中国运输事件数量: " + str(len(tran_event_list)))

    # 柱状图
    weapon_labels = ds.weapon_labels
    weapon_counts = [0 for i in range(len(weapon_labels))]
    for obj in tran_event_list:
        if obj['weaptype1'] is not None and obj['weaptype1'] > 0:
            weapon_counts[int(obj['weaptype1']) - 1] += 1
    labels = []
    counts = []
    for i in range(len(weapon_labels)):
        if weapon_counts[i] != 0:
            labels.append(weapon_labels[i])
            counts.append(weapon_counts[i])

    plt.bar(range(len(counts)), counts, tick_label=labels)
    plt.title("Bar chart")

    plt.show()
    plt.savefig("BarChart.jpg")



# 中国
# 受害目标：巴士站


# 受害目标：中国巴士、中国巴士站


if __name__ == "__main__":
    # test('excelData.xlsx')
    file_name = 'excelData.xlsx'
    #china_ten_tran('excelData.xlsx')
    #china_ten_tran_weapon_pie('excelData.xlsx')
    ds = DataSet(file_name)
    china_bus(ds)
