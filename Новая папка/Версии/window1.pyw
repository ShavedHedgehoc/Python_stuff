import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot
from PyQt5 import QtCore
import os
import time
import shutil



class SignalSender(QObject):
    signal = pyqtSignal(int,int)

class Example(QWidget):
    
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):        
        self.c=SignalSender()
        self.c.signal.connect(self.blb, )
        self.btn = QPushButton("", self)        
        self.btn.clicked.connect(self.buttonClicked)
        self.lw = QListWidget(self)
        self.tl=QLabel(self)
        self.pb=QProgressBar(self)
        hbox = QHBoxLayout()
        hbox.addWidget(self.lw)
        vbox = QVBoxLayout()
        vbox.addLayout(hbox)
        vbox.addWidget(self.pb)   
        vbox.addWidget(self.tl)        
        vbox.addWidget(self.btn)        
        self.setLayout(vbox)
        self.listapp()
        self.setWindowTitle('Statusbar')        
        self.show()
        
    def buttonClicked(self):
        self.collect_files()
        
    def blb(self,arg1,arg2):
##        self.tl.setText(str(arg)+"Поймал!")
##        self.tl.update()
        self.pb.setValue(100*(arg2/arg1))
        print(str(arg1))
    
    def listapp(self):
        for ff in get_files():
            self.lw.addItem(ff)
        self.tl.setText("Файлов к обработке: "+str(len(get_files())))
        if len(get_files())==0:
            self.btn.setText("Запустить рассчет")
        else:
            self.btn.setText("Запустить обработку файлов")

    def collect_files(self):    
        files=get_files()
        cur_directory=get_work_dir()
        dest_directory=cur_directory+'\\Архив\\'
        os.chdir(cur_directory)
        if not os.path.isdir(dest_directory):    os.makedirs(dest_directory)
        count_passed=0
        for ff in files:
            shutil.move(os.path.join(cur_directory, ff),os.path.join(dest_directory, ff))
            count_passed+=1
            self.c.signal.emit(count_passed, len(files))
            time.sleep(0.2)
            self.update
    
def get_work_dir():
        DirConst='X:\Рабочая папка\ГСС\МЮС\Новая папка'
        return DirConst

def get_files():    
    work_directory=get_work_dir()
    files=os.listdir(work_directory)
    new_files=[]
    for one_file in files:
        if os.path.splitext(one_file)[-1]=='.xlsx'\
           and os.path.splitext(one_file)[0].find('Взвешивание')!=-1:\
           new_files.append(one_file)
    return(new_files)    
##        sender=self.sender()
##        
##        if sender.text()=="Запустить обработку файлов":
##            sender.setText("Запустить рассчет")
##            self.MyObject.send(1)
##
##            self.tl.setText("Обрабатываю файлы...")
##            self.collect_files()
##        elif sender.text()=="Запустить рассчет":
##            self.tl.setText("Рассчитываю...")
##            self.MyObject.send(2)
##            
##            sender.setText("Выйти")
##        else:
##            self.close()
    
            

##def read_xl_file(r_file):
##    xl = pd.ExcelFile(r_file)
##    sh=xl.sheet_names[0]
##    r_df=pd.read_excel(xl, sheet_name=sh, header=None, dtype=str)
##    return r_df
##def get_work_dir():
##        DirConst='X:\Рабочая папка\ГСС\МЮС\Новая папка'
##        return DirConst
##def get_files():    
##    work_directory=get_work_dir()
##    files=os.listdir(work_directory)
##    new_files=[]
##    for one_file in files:
##        if os.path.splitext(one_file)[-1]=='.xlsx'\
##           and os.path.splitext(one_file)[0].find('Взвешивание')!=-1:\
##           new_files.append(one_file)
##    return(new_files)
##
##class Example(QWidget):
##    my_signal=pyqtSignal(int)
##    
##
##    def __init__(self):
##        super().__init__()
##        self.initUI()
##
##    def initUI(self):
##        self.btn = QPushButton("", self)        
##        self.btn.clicked.connect(self.buttonClicked)        
##        self.lw = QListWidget(self)
##        self.tl=QLabel(self)
##        hbox = QHBoxLayout()
##        hbox.addWidget(self.lw)
##        vbox = QVBoxLayout()
##        vbox.addLayout(hbox)
##        vbox.addWidget(self.tl)        
##        vbox.addWidget(self.btn)        
##        self.setLayout(vbox)
##        self.listapp()
##        self.setWindowTitle('Statusbar')        
##        self.show()
##
##    @pyqtSlot()
##    def signal_handler(self,*args):    
##        for arg in args:
##            self.tl.setText(str(arg))    
##
##    
##        
##    def listapp(self):
##        for ff in get_files():
##            self.lw.addItem(ff)
##        self.tl.setText("Файлов к обработке: "+str(len(get_files())))
##        if len(get_files())==0:
##            self.btn.setText("Запустить рассчет")
##        else:
##            self.btn.setText("Запустить обработку файлов")            
##        
##    
##    def buttonClicked(self):
##        sender=self.sender()
##        
##        if sender.text()=="Запустить обработку файлов":
##            sender.setText("Запустить рассчет")
##            
##
##            self.tl.setText("Обрабатываю файлы...")
##            self.collect_files()
##        elif sender.text()=="Запустить рассчет":
##            self.tl.setText("Рассчитываю...")
##            
##            sender.setText("Выйти")
##        else:
##            self.close()
##    
##    def collect_files(self):
##        MyObject = SignalSender()
##        MyObject.signal.connect(signal_handler,)
##    
##        files=get_files()
##        cur_directory=get_work_dir()
##        dest_directory=cur_directory+'\\Архив\\'
##        os.chdir(cur_directory)
##        if not os.path.isdir(dest_directory):    os.makedirs(dest_directory)
##        count_passed=0
##        for ff in files:
##            shutil.move(os.path.join(cur_directory, ff),os.path.join(dest_directory, ff))
##            count_passed+=1
##            MyObject.send(count_passed)
##            
##            
##            
##    
    

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()    
    sys.exit(app.exec_())
