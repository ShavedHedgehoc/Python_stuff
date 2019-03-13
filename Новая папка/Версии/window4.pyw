import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, QCoreApplication
from PyQt5 import QtCore
import os
import time
import shutil
import threading
import pandas as pd

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
        signal.emit("Обработка",count_passed, len(new_files))
    signal.emit("Запись",count_passed, len(new_files))
    df_collect_files[6] = df_collect_files[6].apply(pd.to_numeric)
    df_collect_files.drop_duplicates(inplace=True)
    df_collect_files.to_excel('All.xlsx',index=False,header=False)




@thread
def group_batches(signal):
        cur_directory=get_work_dir()
        new_files=get_files()
        signal.emit("Начало группировки",0, len(new_files))
        df_collect_files = read_xl_file('All.xlsx')
        df_varka=read_xl_file('Варки.xlsx')
        df_collect_files[6] = df_collect_files[6].apply(pd.to_numeric)
        df_first_group=df_collect_files[[1,2,3,4,5,6]].groupby([1,2])
        df_data=pd.DataFrame(dtype=str)
        df_data2=pd.DataFrame(dtype=str)
        name_passed=0
        for name, group in df_first_group:
                t_list=[]
                df_second_group=group.groupby([4])[6].sum()
                t_list.append(name[0])
                t_list.append(name[1])
                for i in df_second_group.index:
                        ##print(str(i)) КОЛИЧЕСТВО СТРОК!!!!!
                        name_passed+=1
                        t_list.append(i)
                        t_list.append(df_second_group[i])
                        signal.emit("Группировка",name_passed, len(df_first_group)+1)                        
                t_series=pd.Series(t_list)
                df_data=df_data.append(t_series,ignore_index=True)
                        
        ##df_data.to_excel('All2.xlsx',index=False,header=False)
        
        
        varka_rows_passed=0
        signal.emit("Начало обработки варок",0, 0)

        for row in df_varka.itertuples():
                ll=list(row)
                newrow=df_data[(df_data[0]==row[4])&(df_data[1].map(lambda x: int(x))==int(row[7]))]
                if len(newrow)!=0:
                        sd=newrow.values.tolist()
                        sd1=sum(sd,[])                        
                        sd2=sd1[2:]                        
                        ll.extend(sd2)                                
                r_series=pd.Series(ll)
                df_data2=df_data2.append(r_series,ignore_index=True)
                varka_rows_passed+=1
                signal.emit("Обработка варок",varka_rows_passed, len(df_varka.index))
        df_data2.to_excel('All3.xlsx',index=False,header=False)
        signal.emit("Конец",0, len(new_files))


class MyWindow(QWidget):
    signal = pyqtSignal(str,int,int, name="signal")
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):        
        self.btn = QPushButton("Обработать файлы взвешиваний", self)        
        self.btn.clicked.connect(lambda: collect_files(self.signal))
        self.btn0 = QPushButton("Прикрепить варки к архиву", self)
        self.btn0.clicked.connect(lambda: group_batches(self.signal))
        self.btn1 = QPushButton("Выход", self)
        self.btn1.clicked.connect(sys.exit)
        self.tl0=QLabel(self)
        self.tl=QLabel(self)
        self.pb=QProgressBar(self)
        hbox = QHBoxLayout()
        hbox.addWidget(self.tl0)
        vbox = QVBoxLayout()
        vbox.addLayout(hbox)
        vbox.addWidget(self.pb)   
        vbox.addWidget(self.tl)        
        vbox.addWidget(self.btn)
        vbox.addWidget(self.btn0)
        vbox.addWidget(self.btn1)        
        self.setLayout(vbox)
        self.setWindowTitle('Прослеживаемость')
        self.signal.connect(self.FormRefresh,QtCore.Qt.QueuedConnection)
        self.FormInit()
        self.show()        
        
    def FormRefresh(self,lll,data,data2):
        if lll=="Обработка":
                self.tl.setText("Обработано файлов: "+ str(data)+" из "+str(data2))
                self.pb.setValue(100*(data/data2))        
                self.btn.setDisabled(True)
        if lll=="Запись":
                self.tl.setText("Записываю файл архива...")
                self.pb.setVisible(False)
                self.btn.setDisabled(True)
        if lll=="Конец":
                self.tl.setText("Обработка закончена")
                self.pb.setVisible(False)
                self.btn.setDisabled(True)
                self1.btn.setDisabled(False)
        
        if lll=="Начало группировки":
                self.tl.setText("Группирую данные взвешиваний....")
                self.pb.setVisible(True)
                self.pb.setValue(0)
        if lll=="Группировка":
                self.tl.setText("Группирую данные взвешиваний.... \nОбработано строк: "+ str(data)+" из "+str(data2))
                self.pb.setValue(100*(data/data2))        
                self.btn.setDisabled(True)

        if lll=="Начало обработки варок":
                self.tl.setText("Обрабатываю файл варок...")
                self.pb.setValue(0)        
                self.btn.setDisabled(True)
                self.btn1.setDisabled(True)


        if lll=="Обработка варок":
                self.tl.setText("Обрабатываю файл варок... \nОбработано строк: "+ str(data)+" из "+str(data2))
                self.pb.setValue(100*(data/data2))        
                self.btn.setDisabled(True)

    def FormInit(self):
        if len(get_files())==0:
                self.btn.setDisabled(True)
                self.tl0.setText("В директории нет файлов для обработки... \nВы можете выполнить синхронизацию файлов и запустить программу еще раз \
для внесения данных взвешиваний \nили прикрепить данные взвешивания к архиву варок")
                self.pb.setVisible(False)
        else:
                self.tl0.setText("В директории есть   "+str(len(get_files()))+ " файлов для обработки... \n")               
                self.pb.setValue(0)
     
            
        

    
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWindow()    
    sys.exit(app.exec_())
