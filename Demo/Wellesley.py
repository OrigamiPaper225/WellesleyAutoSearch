# Remaining Tasks:
# Enable user to enter input to change where search counter starts from
# This can be in a separate tab
# Enable user to possibly update options, howerbver, that is hard

import pandas
import webbrowser
import itertools
import pyperclip
# from selenium import webdriver
# from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.common.by import By
from PyQt5 import QtCore

from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QFormLayout
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QDialog
from PyQt5.QtWidgets import QDialogButtonBox
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QObject
import sys


def nameStatus():
    try:
        nameMsg.setText(name)
    except:
        nameMsg.setText('Empty cell, skip to next')
        print('Empty cell, skip to next')


def companyNameStatus():
    try:
        companyMsg.setText(copyItem)
    except:
        companyMsg.setText('Empty cell, skip to next')
        print('Empty cell, skip to next')


def job():
    try:
        jobMsg.setText(title)
        companyMsg.setText(copyItem)
    except:
        jobMsg.setText('Empty cell, skip to next')
        print('Empty cell, skip to next')


def counter():
    try:
        counter = str(count)
        countMsg.setText(counter)
        print('this is count' + counter)
    except:
        countMsg.setText('Empty cell, skip to next')
        print('Empty cell, skip to next')


count = 0



def startSearch():
    global count

    global name
    global copyItem
    global title

    #data = pandas.read_csv('/Users/jamesloh/PycharmProjects/WellesleyAutoSearch/datasets/data1.csv')
    data = pandas.read_csv('~/datasets/data1.csv')
    names = data.loc[:, "Name "]

    def search_item(search_query):
        webbrowser.open("https://www.linkedin.com/search/results/all/?keywords=" +
                        str(search_query) + "&origin=GLOBAL_SEARCH_HEADER&sid=(s5")

    name = names[count]

    def iterate():
        print('Name:')
        print(name)

        print('Count out of 2100:')
        print(count)
        print(title)
        search_item(name)

    # Trying to scroll down to experience page
    # def scroll():
    #     # element = driver.find_element_by_id("experience")
    #     # actions = ActionChains(driver)
    #     # actions.move_to_element(element).perform()
    #     driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
    # for name in names:
    # starts here
    companyName = data.loc[count, 'Company']
    title = data.loc[count, 'Title']
    # if pandas.isnull(copyItem) and pandas.isnull(data.loc[count, 'Name ']):
    # if pandas.isnull(data.loc[:, 'Name ']): # Tried making it skip over nan
    #     # print('Count out of 2100:')
    #     # count += 1
    #     print('Empty row, proceeding to next...')
    #     iterate()
    if pandas.isnull(companyName):
        print(companyName)
        count += 1
        iterate()
        # input("Press Enter to continue...")
        # scroll()
        # input("Press Enter to continue...")
    else:
        print('Company name has been copied to your clipboard')
        copyItem = companyName
        pyperclip.copy(' ' + copyItem)
        count += 1
        print(copyItem)
        iterate()
        # input("Press Enter to continue...")
        # scroll()
        # input("Press Enter to continue...")

    # for name in itertools.islice(names, start, 2139):
    #     copyItem = data.loc[count, 'Company']
    #     # if pandas.isnull(copyItem) and pandas.isnull(data.loc[count, 'Name ']):
    #     # if pandas.isnull(data.loc[:, 'Name ']): # Tried making it skip over nan
    #     #     # print('Count out of 2100:')
    #     #     # count += 1
    #     #     print('Empty row, proceeding to next...')
    #     #     iterate()
    #     if pandas.isnull(copyItem):
    #         # nameMsg = name
    #         #
    #         # companyMsg = copyItem
    #         # countMsg = count
    #         print(copyItem)
    #         count += 1
    #         iterate()
    #         # input("Press Enter to continue...")
    #         # scroll()
    #         # input("Press Enter to continue...")
    #     else:
    #         print('Company name has been copied to your clipboard')
    #         pyperclip.copy(' ' + copyItem)
    #         count += 1
    #         print(copyItem)
    #         # companyMsg = copyItem
    #         iterate()
    #         # input("Press Enter to continue...")
    #         # scroll()
    #         # input("Press Enter to continue...")


app = QApplication(sys.argv)
window = QWidget()
window.setStyleSheet("background-color:lightgrey;")
window.setWindowTitle('Wellesley LinkedIn Searcher')
window.setFixedSize(300, 350)
# layout = QVBoxLayout()

btns = QDialogButtonBox()

btn = QPushButton('Search')
btn.clicked.connect(startSearch)
btn.clicked.connect(nameStatus)
btn.clicked.connect(companyNameStatus)
btn.clicked.connect(job)
btn.clicked.connect(counter)

dlgLayout = QVBoxLayout()

formLayout = QFormLayout()
formLayout.addWidget(btn)
btn.setFixedSize(150, 25)

startAt = QLineEdit()
nameMsg = QLabel('')
companyMsg = QLabel('')
jobMsg = QLabel('')
countMsg = QLabel('')

formLayout.addWidget(startAt)
formLayout.addWidget(nameMsg)
formLayout.addWidget(companyMsg)
formLayout.addWidget(jobMsg)
formLayout.addWidget(countMsg)

formLayout.addRow('Start Search:', btn)
formLayout.addRow('Start At:', startAt)
formLayout.addRow('Name:', nameMsg)
formLayout.addRow('Company Name:', companyMsg)
formLayout.addRow('Job:', jobMsg)
formLayout.addRow('Count:', countMsg)


class eventFilter(QtCore.QObject):
    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.KeyPress:
            print(startAt.text())
            if event.key() == 16777220:
                global count
                count = int(startAt.text())
        return obj.eventFilter(obj, event)


# startAt.returnPressed()
#

filter = eventFilter(startAt)
startAt.installEventFilter(filter)


dlgLayout.addLayout(formLayout)
dlgLayout.setAlignment(Qt.AlignCenter)

window.setLayout(dlgLayout)
window.show()
sys.exit(app.exec_())
