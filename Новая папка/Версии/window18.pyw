import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, QCoreApplication
from PyQt5 import QtCore
import os
import time
import shutil
import threading
import pandas as pd
import itertools
import numpy as np

def thread(my_func):
        def wrapper(*args, **kwargs):
                my_thread = threading.Thread(target=my_func, args=args, kwargs=kwargs)
                my_thread.start()
        return wrapper

def get_work_dir():
    DirConst='X:\Рабочая папка\ГСС\МЮС\Новая папка'
    return DirConst

def read_xl_file(r_file):
    xl = pd.ExcelFile(r_file)
    sh=xl.sheet_names[0]
    r_df=pd.read_excel(xl, sheet_name=sh, header=None, dtype=str)
    return r_df

def get_files():    
    work_directory=get_work_dir()
    files=os.listdir(work_directory)
    new_files=[]
    for one_file in files:
        if os.path.splitext(one_file)[-1]=='.xlsx'\
           and os.path.splitext(one_file)[0].find('Взвешивание')!=-1:\
           new_files.append(one_file)        
    return(new_files)

@thread
def collect_files(signal):
        cur_directory=get_work_dir()
        new_files=get_files()
        if os.path.exists(os.path.join(cur_directory, 'All.xlsx')):
                df_collect_files = read_xl_file('All.xlsx')
        else:
                df_collect_files=None
        dest_directory=cur_directory+'\\Архив\\'
        os.chdir(cur_directory)
        if not os.path.isdir(dest_directory):    os.makedirs(dest_directory)
        count_passed=0
        for ff in new_files:
                df_temp=read_xl_file(ff)
                if df_temp.iloc[0,0]!='{Document.ШтрихкодЕмкости}':
                        df_collect_files=pd.concat([df_collect_files,df_temp],ignore_index=True)
                        shutil.move(os.path.join(cur_directory, ff),os.path.join(dest_directory, ff))
                else:
                        os.remove(os.path.join(cur_directory, ff))
                count_passed+=1        
                signal.emit(["Обработка",True,count_passed, len(new_files),("Обработано файлов: "+ str(count_passed)+" из "+str(len(new_files))), True,True,True])        
        signal.emit(["Записываю файл архива...",False,0,0," ",True,True,True])
        df_collect_files[6] = df_collect_files[6].apply(pd.to_numeric)
        df_collect_files.drop_duplicates(inplace=True)
        df_collect_files.to_excel('All.xlsx',index=False,header=False)
        signal.emit(["Архив записан...",False,0,0," ",True,False,False])




@thread
def group_batches(signal):
        signal.emit(["Читаю файлы...",False,0,0,"", True,True,True])
        df_all=read_xl_file('All.xlsx')
        df8=[]
        signal.emit(["Считаю",False,0,0,"", True,True,True])        
        df_all.columns=['Ключ','Варка','Код','Наименование','Квазипартия','Взвесил', 'Количество']        
        df_all['Количество'] = df_all['Количество'].apply(pd.to_numeric) ## Преобразуем значения из строкового в числовое
        df8=df_all.groupby(['Варка','Код','Квазипартия'])['Количество'].sum().reset_index()
        df15=df8.groupby(['Варка','Код'])['Квазипартия','Количество'].agg(lambda x: list(x)).reset_index()
     ##df15=df15.apply(lambda row: [row['Варка'],row['Код'],list(np.ravel(list(zip(row['Квазипартия'],row['Количество']))))], axis=1)

        ##df15['v7']=df15.apply(lambda row: [row['Варка'],row['Код']]+list(np.ravel(list(zip(row['Квазипартия'],row['Количество'])))), axis=1)
        df15=df15.apply(lambda row: [row['Варка'],row['Код']]+list(np.ravel(list(zip(row['Квазипартия'],row['Количество'])))), axis=1)



        ##df15['v6'] = df15[['Варка', 'Код']].values.tolist()
        
        ##df15=df15.apply(lambda row: list(np.ravel(list(zip(row['Квазипартия'],row['Количество'])))), axis=1)
        df16 = pd.DataFrame(df15.values.tolist())

     
        
        
        signal.emit(["Пишу",False,0,0,"", True,True,True]) 
        df15.to_excel('All7.xlsx',index=False,header=True)
        df16.to_excel('All8.xlsx',index=False,header=False)
        signal.emit(["Архив записан...",False,0,0,"", True,True,False])


class MyWindow(QWidget):
    signal = pyqtSignal(list, name="signal")
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):        
        self.FirstBtn = QPushButton("Обработать файлы взвешиваний", self)        
        self.FirstBtn.clicked.connect(lambda: collect_files(self.signal))
        self.SecondBtn = QPushButton("Прикрепить варки к архиву", self)
        self.SecondBtn.clicked.connect(lambda: group_batches(self.signal))
        self.ExitBtn = QPushButton("Выход", self)
        self.ExitBtn.clicked.connect(sys.exit)
        self.StatusLabel=QLabel(self)
        self.InfoLabel=QLabel(self)
        self.pb=QProgressBar(self)
        hbox = QHBoxLayout()
        hbox.addWidget(self.StatusLabel)
        vbox = QVBoxLayout()
        vbox.addLayout(hbox)
        vbox.addWidget(self.pb)   
        vbox.addWidget(self.InfoLabel)        
        vbox.addWidget(self.FirstBtn)
        vbox.addWidget(self.SecondBtn)
        vbox.addWidget(self.ExitBtn)        
        self.setLayout(vbox)
        self.setWindowTitle('Прослеживаемость')
        self.signal.connect(self.FormRefresh,QtCore.Qt.QueuedConnection)
        self.FormInit()
        self.show()        
        
    def FormRefresh(self,SignalList):
            ## Перерисовка формы
            ## 
            self.StatusLabel.setText(SignalList[0])
            self.pb.setVisible(SignalList[1])
            if SignalList[1]:  self.pb.setValue(100*(SignalList[2]/SignalList[3]))
            self.InfoLabel.setText(SignalList[4])
            self.FirstBtn.setDisabled(SignalList[5])
            self.SecondBtn.setDisabled(SignalList[6])
            self.ExitBtn.setDisabled(SignalList[7])

    def FormInit(self):
        if len(get_files())==0:
                self.FirstBtn.setDisabled(True)
                self.StatusLabel.setText("В директории нет файлов для обработки... \nВы можете выполнить синхронизацию файлов и запустить программу еще раз \
для внесения данных взвешиваний \nили прикрепить данные взвешивания к архиву варок")
                self.pb.setVisible(False)
        else:
                self.StatusLabel.setText("В директории есть   "+str(len(get_files()))+ " файлов для обработки...                                             ")               
                self.pb.setValue(0)
     
            
        

    
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWindow()    
    sys.exit(app.exec_())
