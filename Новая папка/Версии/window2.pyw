import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot
from PyQt5 import QtCore
import os
import time
import shutil
import threading


def thread(my_func):
        def wrapper(*args, **kwargs):
            my_thread = threading.Thread(target=my_func, args=args, kwargs=kwargs)
            my_thread.start()
        return wrapper

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
@thread
def collect_files(signal):    
    cur_directory=get_work_dir()
    new_files=get_files()
    ##signal.emit(new_files,0, len(new_files))
    dest_directory=cur_directory+'\\Архив\\'
    os.chdir(cur_directory)
    if not os.path.isdir(dest_directory):    os.makedirs(dest_directory)
    count_passed=0
    for ff in new_files:
        shutil.move(os.path.join(cur_directory, ff),os.path.join(dest_directory, ff))
        count_passed+=1
        ##signal.emit(new_files,count_passed, len(new_files))
        signal.emit(ff,count_passed, len(new_files))
        

class Example(QWidget):
    signal = pyqtSignal(str,int,int, name="signal")
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):        
        
        self.btn = QPushButton("", self)        
        self.btn.clicked.connect(lambda: collect_files(self.signal))
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
        
        self.setWindowTitle('Statusbar')
        self.signal.connect(self.blb,QtCore.Qt.QueuedConnection)
        self.pp()
        self.show()
        
        
    def blb(self,lll,data,data2):
        ##print(self.lw.QListWidgetItem(1))
        self.tl.setText("Обработано файлов: "+ str(data)+" из "+str(data2))
        self.pb.setValue(100*(data/data2))

    def pp(self):
        self.pb.setValue(0)
        self.btn.setText("Запуск")
        for ff in get_files():
            self.lw.addItem(ff)
        
            
        

    
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()    
    sys.exit(app.exec_())
