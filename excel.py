import matplotlib.pyplot as plt
import xlrd


# 将数据封装为对象列表
class DataSet:

    def load_data(self, file_name):
        try:
            self.origin = xlrd.open_workbook(file_name)
            print("data load success!")
        except Exception as e:
            print(str(e))

    def transfor_origin(self):
        table = self.origin.sheets()[0]
        nrows = table.nrows
        ncols = table.ncols
        colnames = table.row_values(0)
        for ri in range(1, nrows):
            row = table.row_values(ri)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                self.obj_list.append(app)

    def __init__(self, file_name=None, obj_list=None):
        self.obj_list = []

        if file_name is not None:
            self.load_data(file_name)
            self.transfor_origin()

        if obj_list is not None:
            self.obj_list = obj_list

        # 攻击
        self.weapon_labels = ['Biological', 'Chemical', 'Radiological', 'Nuclear',
                              'Firearms', 'Explosives', 'Fake Weapons', 'Incendiary',
                              'Melee', 'Vehicle', 'Sabotage Equipment', 'Other', 'Unknown']

    def get_obj_list(self):
        return self.obj_list


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


class ResultSet:

    def __init__(self, ds, df):
        self.event_list = []
        self._ds = ds
        self._df = df

    def filter(self):
        for row in self._ds.get_obj_list():
            if self._df.filter(row):
                self.event_list.append(row)

    def count_kill_and_wound(self):
        wound_count = 0
        kill_count = 0
        for row in self.event_list:
            if row["nwound"] is not None and row["nwound"] != 0 and row["nwound"] != '':
                wound_count += int(row["nwound"])
            if row["nkill"] is not None and row["nkill"] != 0 and row["nkill"] != '':
                kill_count += int(row["nkill"])
        print("受伤人数： " + str(wound_count))
        print("直接死亡人数： " + str(kill_count))
        return wound_count, kill_count

    def count_event(self):
        print("事件总数： " + str(len(self.event_list)))
        return len(self.event_list)

    def count_weapen_info(self):
        weapon_labels = ds.weapon_labels
        weapon_counts = [0 for i in range(len(weapon_labels))]
        for obj in self.event_list:
            if obj['weaptype1'] is not None and obj['weaptype1'] > 0:
                weapon_counts[int(obj['weaptype1']) - 1] += 1
        return weapon_labels, weapon_counts

    def weapon_bar(self, topic):
        """柱状图"""
        weapon_labels, weapon_counts = self.count_weapen_info()
        labels = []
        counts = []
        for i in range(len(weapon_labels)):
            if weapon_counts[i] != 0:
                labels.append(weapon_labels[i])
                counts.append(weapon_counts[i])

        plt.bar(range(len(counts)), counts, tick_label=labels)
        plt.title(topic)
        #plt.show()
        plt.savefig(topic + "_bar_" +  ".png")
        plt.show()

    def weapon_pie(self, topic):
        """扇形图"""
        weapon_labels, weapon_counts = self.count_weapen_info()
        labels = []
        counts = []
        for i in range(len(weapon_labels)):
            if weapon_counts[i] != 0:
                labels.append(weapon_labels[i])
                counts.append(weapon_counts[i])

        plt.figure()
        plt.pie(counts, labels=labels, autopct='%1.2f%%')
        plt.title(topic)
        plt.savefig(topic + "_pie_" + ".png")
        plt.show()

    def further_filter(self, df):
        ds = DataSet(obj_list=self.event_list)
        return ResultSet(ds, df)


# 受害目标 运输，targtype1 19
# 事件数量
# 死亡人数 nkill
# 受伤人数 nwound
def count_event_wound_kill(ds, country=None, targtype1=None, targsubtype1=None):
    df = DataFilter()
    if country is not None:
        df.add_filter(lambda x: x["country"] == country)
    if targtype1 is not None:
        df.add_filter(lambda x: x['targtype1'] == targtype1)
    if targsubtype1 is not None:
        df.add_filter(lambda x: x["targsubtype1"] == targsubtype1)
    rs = ResultSet(ds, df)
    rs.filter()
    event_count = rs.count_event()
    wound_count, kill_count = rs.count_kill_and_wound()
    return event_count, wound_count, kill_count


# 袭击方式分析
# 扇形图、直方图
def draft_pie_bar(ds, country=None, targtype1=None, targsubtype1=None, topic=None):
    df = DataFilter()
    if country is not None:
        df.add_filter(lambda x: x["country"] == country)
    if targtype1 is not None:
        df.add_filter(lambda x: x['targtype1'] == targtype1)
    if targsubtype1 is not None:
        df.add_filter(lambda x: x["targsubtype1"] == targsubtype1)
    rs = ResultSet(ds, df)
    rs.filter()
    rs.weapon_bar(topic)
    rs.weapon_pie(topic)


#
#
# 运输
# 攻击方式
def chain_tran(ds, country):
    df = DataFilter()
    df.add_filter(lambda x: x["country"] == country)
    df.add_filter(lambda x: x['targtype1'] == 19)
    rs = ResultSet(ds, df)
    rs.filter()
    rs.count_event()


# 受害目标：中国巴士、中国巴士站
def chain_bus(ds, country):
    df = DataFilter()
    df.add_filter(lambda x: x["country"] == country)
    df.add_filter(lambda x: x["targsubtype1"] == 99)
    rs = ResultSet(ds, df)
    rs.filter()
    event_count = rs.count_event()
    wound_count, kill_count = rs.count_kill_and_wound()
    return event_count, wound_count, kill_count


# 受害目标：中国巴士、中国巴士站
def chain_bus_station(ds, country):
    df = DataFilter()
    df.add_filter(lambda x: x["country"] == country)
    df.add_filter(lambda x: x["targsubtype1"] == 101)
    rs = ResultSet(ds, df)
    rs.filter()
    event_count = rs.count_event()
    wound_count, kill_count = rs.count_kill_and_wound()
    return event_count, wound_count, kill_count


# 巴士站
# 受伤人数
# 死亡人数
# 攻击手段
# 中国
def compute_china(ds):
    tran_event_count, tran_wound_count, tran_kill_count = count_event_wound_kill(ds, country=44, targtype1=19)
    draft_pie_bar(ds, country=44, targtype1=19, topic="chinaTran")
    print("")

    bus_event_count, bus_wound_count, bus_kill_count = count_event_wound_kill(ds, country=44, targsubtype1=99)
    draft_pie_bar(ds, country=44, targsubtype1=99, topic="chinaBus")


    print("")

    bus_station_event_count, bus_station_wound_count, bus_station_kill_count \
        = count_event_wound_kill(ds, country=44, targsubtype1=101)
    draft_pie_bar(ds, country=44, targsubtype1=101, topic="chinaBusStation")

    print("")


# 美国
def compute_america(ds):
    tran_event_count, tran_wound_count, tran_kill_count = count_event_wound_kill(ds, country=217, targtype1=19)
    draft_pie_bar(ds, country=217, targtype1=19, topic="AmericaTran")


    print("")

    bus_event_count, bus_wound_count, bus_kill_count = count_event_wound_kill(ds, country=217, targsubtype1=99)
    draft_pie_bar(ds, country=217, targsubtype1=99, topic="AmericaBus")

    print("")

    bus_station_event_count, bus_station_wound_count, bus_station_kill_count \
        = count_event_wound_kill(ds, country=217, targsubtype1=101)
    draft_pie_bar(ds, country=217, targsubtype1=101, topic="AmericaBusStation")
    print("")


# 法国
def compute_france(ds):
    tran_event_count, tran_wound_count, tran_kill_count = count_event_wound_kill(ds, country=69, targtype1=19)
    draft_pie_bar(ds, country=69, targtype1=19, topic="FranceTran")
    print("")

    bus_event_count, bus_wound_count, bus_kill_count = count_event_wound_kill(ds, country=69, targsubtype1=99)
    draft_pie_bar(ds, country=69, targsubtype1=99, topic="FranceBus")
    print("")

    bus_station_event_count, bus_station_wound_count, bus_station_kill_count \
        = count_event_wound_kill(ds, country=69, targsubtype1=101)
    draft_pie_bar(ds, country=69, targsubtype1=101, topic="FranceBusStation")
    print("")


# 西班牙
def compute_spain(ds):
    tran_event_count, tran_wound_count, tran_kill_count = count_event_wound_kill(ds, country=185, targtype1=19)
    draft_pie_bar(ds, country=185, targtype1=19, topic="SpainTran")
    print("")

    bus_event_count, bus_wound_count, bus_kill_count = count_event_wound_kill(ds, country=185, targsubtype1=99)
    draft_pie_bar(ds, country=185, targsubtype1=99, topic="SpainBus")
    print("")

    bus_station_event_count, bus_station_wound_count, bus_station_kill_count \
        = count_event_wound_kill(ds, country=185, targsubtype1=101)
    draft_pie_bar(ds, country=185, targsubtype1=101, topic="SpainBusStation")
    print("")


def compute_world(ds):
    tran_event_count, tran_wound_count, tran_kill_count = count_event_wound_kill(ds, targtype1=19)
    draft_pie_bar(ds, targtype1=19, topic="worldTran")
    print("")

    bus_event_count, bus_wound_count, bus_kill_count = count_event_wound_kill(ds, targsubtype1=99)
    draft_pie_bar(ds, targsubtype1=99, topic="worldBus")
    print("")

    bus_station_event_count, bus_station_wound_count, bus_station_kill_count \
        = count_event_wound_kill(ds, targsubtype1=101)
    draft_pie_bar(ds, targsubtype1=101, topic="worldBusStation")
    print("")


if __name__ == "__main__":
    # test('excelData.xlsx')
    file_name = 'excelData.xlsx'
    # china_ten_tran('excelData.xlsx')
    # china_ten_tran_weapon_pie('excelData.xlsx')
    ds = DataSet(file_name)
    compute_china(ds)
    compute_america(ds)
    compute_france(ds)
    compute_spain(ds)
    compute_world(ds)

