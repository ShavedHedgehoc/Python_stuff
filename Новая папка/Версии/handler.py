from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot


class SignalSender(QObject):
    
    signal = pyqtSignal(int)
    
    def __init__(self):        
        QObject.__init__(self)
        
    def send(self,*args):        
        for arg in args:
            self.signal.emit(arg)

@pyqtSlot()
def signal_handler(*args):    
    for arg in args:
        print(arg)    

if __name__ == "__main__":
    
    MyObject = SignalSender()
    MyObject.signal.connect(signal_handler,)
    MyObject.send(5,88,77)
    ##for i in range(10):
      ##  MyObject.send(i,88,77)

