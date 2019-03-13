from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot
 
class PunchingBag(QObject):
    ''' Represents a punching bag; when you punch it, it
        emits a signal that indicates that it was punched. '''
    punched = pyqtSignal(int)
 
    def __init__(self):
        # Initialize the PunchingBag as a QObject
        QObject.__init__(self)
 
    def punch(self,arg):
        ''' Punch the bag '''
        self.punched.emit(arg)



@pyqtSlot()
def say_punched(arg):
    ''' Give evidence that a bag was punched. '''
    print('Bag was punched.')
    print(arg)
 
bag = PunchingBag()
# Connect the bag's punched signal to the say_punched slot
bag.punched.connect(say_punched, )
for i in range(10):
    bag.punch(i)
