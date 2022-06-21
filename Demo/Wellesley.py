# Possibly try and implement scroll, but not that needed
import sys
import pandas
import webbrowser
# import itertools
import pyperclip
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
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


def nameStatus():
    try:
        nameMsg.setText(name)
    except:
        nameMsg.setText('Empty cell, skip to next')
        print('Empty cell, skip to next')


def companyNameStatus():
    try:
        global companyBefore
        companyBefore = copyItem
        companyMsg.setText(copyItem)

    except:
        companyMsg.setText('Empty cell, skip to next')
        print('Empty cell, skip to next')


def job():
    try:
        global titleBefore
        titleBefore = title
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
    global companyBefore
    global titleBefore
    global filepath
    filepath = '/Users/jamesloh/PycharmProjects/WellesleyAutoSearch/datasets/data3.xlsx'
    # filepath = r'C:\WellesleyAutoSearch-Wellesley\datasets\data3.csv'
    # filepath = '/Users/jamesloh/PycharmProjects/SeleniumTest/datasets/data2.csv'
    # filepath = '~/datasets/data1.csv'
    # data = pandas.read_csv('/Users/jamesloh/PycharmProjects/WellesleyAutoSearch/datasets/data1.csv')
    # data = pandas.read_csv('~/datasets/data1.csv')
    global data
    # data = pandas.read_csv(filepath)
    data = pandas.read_excel(filepath)
    global wb
    wb = openpyxl.load_workbook(filepath)
    global sheet_obj
    sheet_obj = wb.active


    names = data.loc[:, "Name "]

    def search_item(search_query):
        webbrowser.open("https://www.linkedin.com/search/results/all/?keywords=" +
                        str(search_query) + "&origin=GLOBAL_SEARCH_HEADER&sid=(s5")

    name = names[count]



    # Trying to scroll down to experience page
    # def scroll():
    #     # element = driver.find_element_by_id("experience")
    #     # actions = ActionChains(driver)
    #     # actions.move_to_element(element).perform()
    #     driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
    # for name in names:
    # starts here
    global companyName
    companyName = data.loc[count, 'Company']
    title = data.loc[count, 'Title']


    def iterate():
        print('Name:')
        print(name)

        print('Count out of 2100:')
        print(count)
        print(title)
        search_item(name)
        global companyDetector

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
# window.setStyleSheet("background-color:lightgrey;")
window.setWindowTitle('Wellesley LinkedIn Searcher')
window.setFixedSize(300, 450)
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

# startAt = QLineEdit()
# nameMsg = QLabel('')
# companyMsg = QLabel('')
# jobMsg = QLabel('')
# countMsg = QLabel('')

companyColumn = QLineEdit()
titleColumn = QLineEdit()

startAt = QLineEdit()
nameMsg = QLabel('')
companyMsg = QLineEdit('')
jobMsg = QLineEdit('')
countMsg = QLabel('')

formLayout.addWidget(companyColumn)
formLayout.addWidget(titleColumn)

formLayout.addWidget(startAt)
formLayout.addWidget(nameMsg)
formLayout.addWidget(companyMsg)
formLayout.addWidget(jobMsg)
formLayout.addWidget(countMsg)

lineEditHeight = 20
lineEditWidth = 125

companyColumn.setFixedSize(lineEditWidth/3, lineEditHeight)
titleColumn.setFixedSize(lineEditWidth/3, lineEditHeight)
startAt.setFixedSize(lineEditWidth, lineEditHeight)
nameMsg.setFixedSize(lineEditWidth, lineEditHeight)
companyMsg.setFixedSize(lineEditWidth, lineEditHeight)
jobMsg.setFixedSize(lineEditWidth, lineEditHeight)
countMsg.setFixedSize(lineEditWidth, lineEditHeight)

formLayout.addRow('Company Column:', companyColumn)
formLayout.addRow('Title Column:', titleColumn)
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
            if event.key() == 16777220 and companyColumn.text() != '': # and titleBefore != title:
                global companyCol
                companyCol = companyColumn.text()
            if event.key() == 16777220 and titleColumn.text() != '': # and titleBefore != title:
                global titleCol
                titleCol = titleColumn.text()
            if event.key() == 16777220 and startAt.text() != '':
                global count
                try:
                    count = (int(startAt.text()))
                except:
                    print('Enter a number into count')
            if event.key() == 16777220 and companyMsg.text() != '': # and companyBefore != copyItem:
                global copyItem
                #copyItem = companyMsg.text()
                newCompanyName = companyMsg.text()
                # data.at[count, 'New Company'] = copyItem
                # print(copyItem)
                # print(data.at[count, 'New Company'])
                #companyVar = 'G' + str(count + 1)
                companyVar = companyCol + str(count + 1)
                companyVar2 = 'A' + str(count+2)
                companyTopVar = companyCol + '1'
                #sheet_obj[companyVar].value = copyItem
                font = Font(color="FF0000")
                sheet_obj[companyVar].value = newCompanyName
                sheet_obj[companyVar].font = font
                sheet_obj[companyTopVar].value = 'New Company'
                # Tried and failed to change next value of first column's Company name to old company name
                # companyDetector = data.loc[count + 1, 'Company']
                # if companyDetector == '':
                #     sheet_obj[companyVar2].value = companyName
                #     print('Old company name')
                print(companyVar)
                wb.save(filepath)
                print('New Company Saved')
            if event.key() == 16777220 and jobMsg.text() != '': # and titleBefore != title:
                global title
                title = jobMsg.text()
                #titleVar = 'H' + str(count + 1)
                titleVar = titleCol + str(count + 1)
                titleTopVar = titleCol + '1'
                sheet_obj[titleVar].value = title
                sheet_obj[titleVar].font = font
                sheet_obj[titleTopVar].value = 'New Title'
                print(titleVar)
                wb.save(filepath)
                print('New title saved')

        return obj.eventFilter(obj, event)


# startAt.returnPressed()
#

companyNameFilter = eventFilter(companyColumn)
companyColumn.installEventFilter(companyNameFilter)

titleNameFilter = eventFilter(titleColumn)
titleColumn.installEventFilter(titleNameFilter)

countFilter = eventFilter(startAt)
startAt.installEventFilter(countFilter)

companyNameFilter = eventFilter(companyMsg)
companyMsg.installEventFilter(companyNameFilter)

jobFilter = eventFilter(jobMsg)
jobMsg.installEventFilter(jobFilter)

dlgLayout.addLayout(formLayout)
dlgLayout.setAlignment(Qt.AlignCenter)

window.setLayout(dlgLayout)
window.show()
sys.exit(app.exec_())
