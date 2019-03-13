import xlwings as xw
import pandas as pd
import os
import sys
import shutil
from PyQt5.QtWidgets import QWidget, QApplication, QLabel, QPushButton

def get_work_dir():
    DirConst='X:\Рабочая папка\ГСС\МЮС\Новая папка'
    return DirConst

def read_xl_file(r_file):
    xl = pd.ExcelFile(r_file)
    sh=xl.sheet_names[0]
    r_df=pd.read_excel(xl, sheet_name=sh, header=None, dtype=str)
    return r_df

def read_data(SheetName, CellName, HeaderValue):
    sh=xw.Book.caller().sheets[SheetName]
    rng=sh.range(CellName).expand()
    df=rng.options(pd.DataFrame, header=HeaderValue, index=0).value ## string!!!!
    return df

def write_data(SheetName, CellName, DFrame, IndValue,HValue):    
    sh=xw.Book.caller().sheets[SheetName]
    sh.clear_contents()    
    sh.range(CellName).options(index=IndValue, header=HValue).value=DFrame
    
def get_files():    
    work_directory=get_work_dir()
    files=os.listdir(work_directory)
    new_files=[]
    for one_file in files:
        if os.path.splitext(one_file)[-1]=='.xlsx'\
           and os.path.splitext(one_file)[0].find('Взвешивание')!=-1:\
           new_files.append(one_file)        
    return(new_files)

def collect_files():
    files=get_files()
    cur_directory=get_work_dir()
    dest_directory=cur_directory+'\\Архив\\'
    os.chdir(cur_directory)
    if not os.path.isdir(dest_directory):    os.makedirs(dest_directory)
    files_count=len(files)
    df_collect_files = read_data("Лист2","A1",0)    
    if files_count==0:
        pass
    else:
        count_passed=0
        count_removed=0
        for ff in files:
            df_temp=read_xl_file(ff)
            if df_temp.iloc[0,0]!='{Document.ШтрихкодЕмкости}': 
                df_collect_files=pd.concat([df_collect_files,df_temp],ignore_index=True)
                shutil.move(os.path.join(cur_directory, ff),os.path.join(dest_directory, ff))
            else: 
                os.remove(os.path.join(cur_directory, ff))
                count_removed+=1
                count_passed+=1
        if df_collect_files.shape[-1]<7: df_collect_files[6]=None        
        df_collect_files[6]=df_collect_files[6].fillna(0)
        df_collect_files[6] = df_collect_files[6].apply(pd.to_numeric)
        df_collect_files.drop_duplicates(inplace=True)
        ## Здесь нужна сортировка файлов
    write_data("Лист3","A1",df_collect_files,False,False)
    return df_collect_files




class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    def initUI(self):
        btn = QPushButton('Запуск', self)
        btn.move(50, 50)
        btn.clicked.connect(self.buttonClicked)        
        self.setGeometry(300, 300, 250, 150)
        self.setWindowTitle('Statusbar')        
        self.show()
    def buttonClicked(self):
        sender=self.sender()
        if sender.text()=="Запуск":
            df=collect_files()##print("Запуск")
            sender.setText("Выпуск")
        elif sender.text()=="Выпуск":
            print("Выпуск")
            sender.setText("Выйти")
        else:
            self.close()

def main():
    app = QApplication(sys.argv)
    ex = Example()    
    sys.exit(app.exec_())

    
