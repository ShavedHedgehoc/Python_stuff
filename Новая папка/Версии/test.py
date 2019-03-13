import xlwings as xw
import pandas as pd
import os
import sys
import shutil
from PyQt5.QtWidgets import QMainWindow, QApplication

class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initWindow()
        b=self.QLAbel("fsfds")
    def initWindow(self):
        self.setGeometry(300, 300, 250, 150)
        self.setWindowTitle('Statusbar')
        self.show()
    def UpdateWindow(self, arg1):
        self.setWindowTitle(arg1)
    def CloseWindow(self):
        self.destroy()




def read_xl_file(r_file):
    xl = pd.ExcelFile(r_file)
    sh=xl.sheet_names[0]
    r_df=pd.read_excel(xl, sheet_name=sh, header=None, dtype=str)
    return r_df

def read_data(SheetName, CellName, HeaderValue):    
    sh=xw.Book.caller().sheets[SheetName]
    rng=sh.range(CellName).expand()
    df=rng.options(pd.DataFrame, header=HeaderValue, index=0).value
    return df

def write_data(SheetName, CellName, DFrame, IndValue,HValue):                                                   ## Процедура записи датафрейма на лист текущей книги
                                                                                                                ## Параметры:
    sh=xw.Book.caller().sheets[SheetName]                                                                       ## SheetName - Имя страницы, CellName - имя первой ячейки
    sh.range(CellName).options(index=IndValue, header=HValue).value=DFrame                                      ## DFrame - датафрейм
                                                                                                                ## IndValue,HValue - булевы, вывод индексов и заголовков
##################################################################################################################################################################################

#####################################################################################################################################################################################
def get_files():                                                                                                ## Процедура сбора файлов взвешивания                               #
    work_directory=os.getcwd()
    
    files=os.listdir(work_directory)                                                                            ## Собирает в текущей директории и                                  #
    new_files=[]                                                                                                ## возвращает массив файлов,                                        #
    for one_file in files:                                                                                      ## имя которых начинается на "Взвешивание",                         #
        if os.path.splitext(one_file)[-1]=='.xlsx' and os.path.splitext(one_file)[0].find('Взвешивание')!=-1: new_files.append(one_file)                                                                          ##                                                                  #
    return(new_files)                                                                                           ##                                                                  #
#####################################################################################################################################################################################


def collect_files():
    df_collect_files=None
    files=get_files()
    cur_directory=os.getcwd()
    dest_directory=cur_directory+'\\Архив\\'    
    if not os.path.isdir(dest_directory):    os.makedirs(dest_directory)    
    files_count=len(files)
    
    return df_collect_files






if __name__ == "__main__":
##def main():  
    df_collect=collect_files()
    
   ## write_data("Лист3","A1",df_collect,False,True)

    
##    df1['код 1С'] = df1['код 1С'].apply(pd.to_numeric)
##    df2=read_data("Лист2","A1",1)
##    df2.columns=['код 1С','Наименование','К-во','Ед. изм.','Партия','Дата','Ед. хр.','Ячейка']
##    df2['код 1С'] = df2['код 1С'].apply(pd.to_numeric)
##    df12=pd.merge(df1,df2,on='код 1С', how='inner')
##    df2=df2[~df2['код 1С'].isin(df12['код 1С'])]
##    df3=df2.sort_values(by=['Наименование'])
##    write_data("Лист3","A1",df3)
    
