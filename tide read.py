
import pandas
import numpy as np
import scipy
import matplotlib.pyplot as plt
import datetime
import dateutil

def say_out(xxx):
    #错误提示
    print(xxx)

def process(y):
    #信号去噪
    global switch_of_process
    switch_of_process = True
    rft=np.fft.rfft(y)
    rft[int(len(rft)/2):] = 0
    return np.fft.irfft(rft)

class Tide (object):
    def __init__(self,filename,time1=datetime.datetime(2000,1,1,00,0),time2 = datetime.datetime.now()):
        self.tide_sheet = pandas.read_excel(filename, sheetname='tide')
        if 'tide' in self.tide_sheet.columns.values:
            if 'time' in self.tide_sheet.columns.values:
                pass
                #############################################################
        else:
            say_out('Please Check Your Format of the File')

        temp_data=self.tide_sheet

        t = []
        for x in temp_data.time:
            t.append(dateutil.parser.parse(x))
        temp_data['format_time'] = t
        temp_data['tide_init'] = temp_data.tide
        temp_data.tide = process(temp_data.tide)

        temp_data = temp_data.ix[temp_data.format_time <= time2]
        temp_data = temp_data.ix[temp_data.format_time >= time1]


    def plot_tide_compare(self):
        self.fig0, self.ax0 = plt.subplots()
        fig = self.fig0
        ax = self.ax0
        temp_data = self.tide_sheet
        if True:
            line1, = ax.plot(temp_data.format_time, temp_data.tide_init, 'o', label='Init')
            line2, = ax.plot(temp_data.format_time, temp_data.tide, linewidth=0.5, color='b', label='Processed')
            plt.xlim(temp_data.format_time.min(), temp_data.format_time.max())

            ax.legend(handles=[line1, line2])
            plt.savefig('x.png')

if __name__ == '__main__':
    name = "test.xlsx"
    x =Tide(name)
    x.plot_tide_compare()
