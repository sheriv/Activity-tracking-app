import os
import xlrd
from mx import DateTime
import xlsxwriter


class Util(object):

    def __init__(self):

        self.save_path = '\\tops'

        self.time_threshold = 20
        self.periods_list = [7, 30, 3, 40000]
        self.make_save_dir(self.save_path)
        self.start()

    def start(self):

        for period in self.periods_list:

            list = []
            for file in os.listdir(os.getcwd()):

                if file.endswith(".xlsx") and '2' in file and '-' in file:

                    if self.period_check(file, period):

                        list.append(file)

            self.write(self.xlsx_to_dict(list), period)

    def period_check(self, filename, period):

        file_dates = filename.split('.')[0].split('-')
        filetime = DateTime.DateTime(int(file_dates[0]), int(file_dates[1]), int(file_dates[2]))
        periodtime = DateTime.now() - DateTime.DateTimeDelta(period)

        if filetime >= periodtime:

            return True

        else:

            return False

    def xlsx_to_dict(self, names_list):

        dict = {}

        for item in names_list:
            
            wb = xlrd.open_workbook(item)
            sh = wb.sheet_by_index(0)


            for i in range(sh.nrows):

                name = sh.cell(i, 0).value
                time = sh.cell(i, 1).value

                if (name in dict) != True:

                    j = self.time_format_timedelta(time)
                    dict[name] = DateTime.DateTimeDelta(0,j[0],j[1],j[2])

                else:

                    k = self.time_format_timedelta(time)
                    dict[name] += DateTime.DateTimeDelta(0,k[0],k[1],k[2])

        return dict

    def time_format_datetime(self, str):

        return str.split('.')[0]

    def time_format_timedelta(self, str):

        l = []
        for i in str.split('.')[0].split(':'):

            l.append(int(i))
        return l

    def make_save_dir(self, save_path):
        self.dir = os.getcwd() + save_path
        
        if not os.path.exists(self.dir):
            
            os.makedirs(self.dir)

    def write(self, wdict, period):

        dict = wdict
        sortedk = sorted(dict, key=lambda k: dict[k], reverse=True)
        path = self.dir + '\\' + str(period) + '.xlsx'
        workbook = xlsxwriter.Workbook(path)

        worksheet = workbook.add_worksheet()
        row = 0
        col = 0
        i = 0

        for key in sortedk:

            if dict[key] > DateTime.DateTimeDelta(0, 0, 0, self.time_threshold):
                worksheet.write(row, col, key)
                worksheet.write(row, col + 1, str(dict[key]))
                row += 1

            else:
                pass

        workbook.close()


A = Util()


