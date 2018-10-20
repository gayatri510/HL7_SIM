#!/usr/bin/python2.7
# encoding=utf8
import sys
reload(sys)
sys.setdefaultencoding('utf8')
from gui import Gui
from PyQt4 import QtGui

def main():
    app = QtGui.QApplication(sys.argv)
    ex = Gui(app)
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
