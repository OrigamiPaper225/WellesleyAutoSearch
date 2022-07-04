# Possibly try and implement scroll, but not that needed
# Add 2 inputs, one for name of top of Company column, and the other for name of top of Name and title.
# If possible, make this a separate tab
# Use Tkinter to select config file
import re
import sys
import os
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
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QToolBar
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QAction
from PyQt5.QtWidgets import QDialogButtonBox
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QObject

import configparser
import tkinter as tk
from tkinter import filedialog as fd
from tkinter.filedialog import askopenfilename
from tkinter import Tk
Tk().withdraw()

config_obj = configparser.ConfigParser()
# Read config.ini file
configpath = './config.ini'
#configpath = '/Users/jamesloh/PycharmProjects/WellesleyAutoSearch/Demo/config.ini'
config_obj.read(configpath)
# Get the postgresql section
#global user_info
user_info = config_obj["user_info"]

filepath = user_info["filepath"]
targetnamecolid = user_info["targetnamecolid"]
targetfirmcolid = user_info["targetfirmcolid"]
targettitlecolid = user_info["targettitlecolid"]
newfirmcol = user_info["newfirmcol"]
newtitlecol = user_info["newtitlecol"]
count = int(user_info["countsaved"])



# def getDataFromConfig():

print(targetnamecolid)
print(targetfirmcolid)
print(newfirmcol)
print(newtitlecol)
print(count)

def getFileName():
    global filepath
    try:
        filepath = askopenfilename()
    except:
        filepath = './Example Excel (completely random profiles).xlsx'


def _getData():
    """Getting data from Excel"""
    global data
    data = pandas.read_excel(filepath)
    #self.names = self.data.loc[:, "Name"]
    print('Data received')
    """Organizing into Variables"""
    print(data)
    print(filepath)



def _prepExcel():
    global sheet_obj
    global wb
    wb = openpyxl.load_workbook(filepath)
    sheet_obj = wb.active

def _prepData():
    #try:
    global names
    global name
    try:
        names = data.loc[:, targetnamecolid]
        print(targetnamecolid)

        #if count-1<0:
        name = names[count]
    except:
        try:
            names = data.loc[:, 'Name']
            name = names[count]
        except:
            try:
                names = data.loc[:,'Candidate']
                name = names[count]
            except:
                try:
                    names = data.loc[:,'Name ']
                    name = names[count]
                except:
                    print('Name has a problem')
    #     statusLabel.setText('Status: Please check Candidate Name')
        # else:
        #     name = names[count - 1]
    # except:# print(count)
    #     print('Names has a problem')
    #try:
    global companyName
    try:
        companyName = data.loc[count, targetfirmcolid]
    except:
        try:
            companyName = data.loc[:, 'Firm']
        except:
            try:
                companyName = data.loc[:, 'Company']
            except:
                print('Company has a problem')
    #     statusLabel.setText('Status: Please check Company Name')
    # except:
    #     print('Company name has a problem')
    #try:
    global title
    try:
        title = data.loc[count, targettitlecolid]
        print(title)
    except:
        print('Title has a problem')
        try:
            title = data.loc[:, 'Title']
        except:
            title = data.loc[:, 'Position']

    #     statusLabel.setText('Status: Please check Title')
    # except:
    #     print('Please re-enter the Name, Company, or Title column')


def setNewConfigData():
    user_info["filepath"] = filepath
    user_info["targetnamecolid"] = targetNameColumn.text()
    print(user_info["targetnamecolid"])
    user_info["targetfirmcolid"] = targetCompanyColumn.text()
    print(user_info["targetfirmcolid"])
    user_info["targettitlecolid"] = targetTitleColumn.text()
    user_info["newfirmcol"] = newCompanyNameColumn.text()
    user_info["newtitlecol"] = newTitleColumn.text()
    user_info["countsaved"] = startAt.text()

    # Write changes back to file

    with open(configpath, 'w') as configfile:
        config_obj.write(configfile)

class AnotherWindow(QWidget):

    lineEditHeight = 15
    lineEditWidth = 125

    # targetnamecol = getConfigData.targetnamecol
    # targetfirmcol = getConfigData.targetfirmcol
    # newfirmcol = getConfigData.newfirmcol
    # newtitlecol = getConfigData.newtitlecol
    # countsaved = getConfigData.countsaved

    """
    This "window" is a QWidget. If it has no parent, it
    will appear as a free-floating window as we want.
    """
    def __init__(self):
        # QMainWindow.__init__(self, parent=parent)
        super().__init__()
        self.setStyleSheet("background-color: #FFFFFF;font-family: Arial;")
        # self.layout = QFormLayout()
        # self.label = QLabel("hi")
        # self.layout.addWidget(self.label)
        # self.setLayout(self.layout)
        # layout = QVBoxLayout()
        # self.label = QLabel("Another Window")
        # layout.addWidget(self.label)
        # self.setLayout(layout)
        self.dlgLayout = QVBoxLayout()
        self.formLayout = QFormLayout()
        self.generalLayout = self.dlgLayout
        self.dlgLayout.addLayout(self.formLayout)
        # self._centralWidget = QWidget(self)
        # self.setCentralWidget(self._centralWidget)
        # self._centralWidget.setLayout(self.generalLayout)
        #self.generalLayout = self.dlgLayout
        self.setLayout(self.dlgLayout)

        self._createInputs2()
        self.organizeText()
        self.setFixedSize(300, 330)
        self.setWindowTitle('Config')
        self._connectSignals()
        _prepData()
        self._prepNewCols()



        #self.title = data.loc[count, 'Title']

    # def _updateData(self):
    #     global names
    #     names = data.loc[:, targetnamecolid]
    #     global companyName
    #     companyName = data.loc[count, targetfirmcolid]
    #     global title
    #     title = data.loc[count, targettitlecolid]

    # targetNameColumn.text =
    # targetCompanyColumn.text =
    # targetTitleColumn.text =
    # newCompanyNameColumn.text =
    # targetTitleColumn.text =
    # startAt.text =
    def _createInputs2(self):
        self.lineEditHeight = 30
        self.lineEditWidth = 125

        global selectFile
        selectFile = QPushButton('Select Excel File')
        selectFile.setStyleSheet("background-color: #EEF3F8;"
                                       "border-radius: 6px;")
        selectFile.setFont(QFont('Arial', 10))
        global targetNameColumn
        targetNameColumn = QLineEdit(user_info["targetnamecolid"])
        targetNameColumn.setFont(QFont('Arial', 10))
        targetNameColumn.setStyleSheet("background-color: #EEF3F8;"
                                       "border-radius: 6px;")
        global targetCompanyColumn
        targetCompanyColumn = QLineEdit(user_info["targetfirmcolid"])
        targetCompanyColumn.setFont(QFont('Arial', 10))
        targetCompanyColumn.setStyleSheet("background-color: #EEF3F8;"
                                       "border-radius: 6px;")
        global targetTitleColumn
        targetTitleColumn = QLineEdit(user_info["targettitlecolid"])
        targetTitleColumn.setFont(QFont('Arial', 10))
        targetTitleColumn.setStyleSheet("background-color: #EEF3F8;"
                                       "border-radius: 6px;")
        global newCompanyNameColumn
        newCompanyNameColumn = QLineEdit(user_info["newfirmcol"])
        newCompanyNameColumn.setFont(QFont('Arial', 10))
        newCompanyNameColumn.setStyleSheet("background-color: #EEF3F8;"
                                       "border-radius: 6px;")
        global newTitleColumn
        newTitleColumn = QLineEdit(user_info["newtitlecol"])
        newTitleColumn.setFont(QFont('Arial', 10))
        newTitleColumn.setStyleSheet("background-color: #EEF3F8;"
                                           "border-radius: 6px;")
        global startAt
        startAt = QLineEdit(user_info["countsaved"])
        startAt.setFont(QFont('Arial', 10))
        startAt.setStyleSheet("background-color: #EEF3F8;"
                                     "border-radius: 6px;")
        startAt.setFocus()

        selectFile.setFixedSize(int(self.lineEditWidth), self.lineEditHeight)
        targetNameColumn.setFixedSize(int(self.lineEditWidth / 2), self.lineEditHeight)
        targetCompanyColumn.setFixedSize(int(self.lineEditWidth / 2), self.lineEditHeight)
        targetTitleColumn.setFixedSize(int(self.lineEditWidth / 2), self.lineEditHeight)
        newCompanyNameColumn.setFixedSize(int(self.lineEditWidth / 2), self.lineEditHeight)
        newTitleColumn.setFixedSize(int(self.lineEditWidth / 2), self.lineEditHeight)
        startAt.setFixedSize(int(self.lineEditWidth/2), self.lineEditHeight)
        global fileNameLabel
        fileNameLabel = QLabel('Current file:' + os.path.basename(filepath))
        fileNameLabel.setFont(QFont('Arial', 10))
        fileNameLabel.setStyleSheet("color:black;font-weight: 600;")

        self.generalLayout.addWidget(selectFile)
        self.generalLayout.addWidget(fileNameLabel)
        self.generalLayout.addWidget(targetNameColumn)
        self.generalLayout.addWidget(targetCompanyColumn)
        self.generalLayout.addWidget(targetTitleColumn)
        self.generalLayout.addWidget(newCompanyNameColumn)
        self.generalLayout.addWidget(newTitleColumn)
        self.generalLayout.addWidget(startAt)
        selectFile.setFont(QFont('Arial', 10))
        selectFile.setStyleSheet(
            "color: white;"
            "background-color: #0A66C2;"
            "border-style: solid;"
            "border-width: 2px;"
            "font-weight: 600;"
            "border-color: #0A66C2;"
            "border-radius: 3px")
        self.savebtn = QPushButton('Save Selection')
        self.savebtn.setFont(QFont('Arial', 10))
        self.savebtn.setStyleSheet(
            "color: white;"
            "background-color: #0A66C2;"
            "border-style: solid;"
            "border-width: 2px;"
            "font-weight: 600;"
            "border-color: #0A66C2;"
            "border-radius: 3px")
        self.savebtn.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        self.generalLayout.addWidget(self.savebtn)
        #self.layout.addLayout(self.layout)

    def organizeText(self):

        self.cl1 = QLabel('Target Name Column:')
        self.cl2 = QLabel('Target Company Column:')
        self.cl3 = QLabel('Target Title Column')
        self.cl4 = QLabel('New Company Column:')
        self.cl5 = QLabel('New Title Column:')
        self.cl6 = QLabel('Start At:')


        self.cl1.setStyleSheet("color:black;font-weight: 600;")
        self.cl1.setFont(QFont('Arial', 10))
        self.cl2.setStyleSheet("color:black;font-weight: 600;")
        self.cl2.setFont(QFont('Arial', 10))
        self.cl3.setStyleSheet("color:black;font-weight: 600;")
        self.cl3.setFont(QFont('Arial', 10))
        self.cl4.setStyleSheet("color:black;font-weight: 600;")
        self.cl4.setFont(QFont('Arial', 10))
        self.cl5.setStyleSheet("color:black;font-weight: 600;")
        self.cl5.setFont(QFont('Arial', 10))
        self.cl6.setStyleSheet("color:black;font-weight: 600;")
        self.cl6.setFont(QFont('Arial', 10))

        self.formLayout.addRow(self.cl1, targetNameColumn)
        self.formLayout.addRow(self.cl2, targetCompanyColumn)
        self.formLayout.addRow(self.cl3, targetTitleColumn)
        self.formLayout.addRow(self.cl4, newCompanyNameColumn)
        self.formLayout.addRow(self.cl5, newTitleColumn)
        self.formLayout.addRow(self.cl6, startAt)

    def targetNameColumnConfig(self):
        #try:
        if targetNameColumn.text() != '':  # and titleBefore != title:
            #self.targetNameID = self.targetNameColumn.text()
            #print('target name column empty')
            global targetnamecolid
            targetnamecolid = targetNameColumn.text()
            print(targetnamecolid)
        #self.targetNameCol = self.targetNameColumn.text()
        self.nameTargetCell = targetnamecolid # + str(MainWindow.count + 1)
        print(self.nameTargetCell)
            #self.font = Font(color="FF0000")
        # except:
        #     print('Enter a job title')

    def targetCompanyColumnConfig(self):
        #try:
        if targetCompanyColumn.text() != '':  # and titleBefore != title:
            global targetcompanycolid
            targetcompanycolid = targetCompanyColumn.text()
            print(targetcompanycolid)
        #self.targetCompanyCol = self.targetCompanyColumn.text()
        self.companyTargetCell = targetcompanycolid + str(count + 1)
        print(self.companyTargetCell)
        #self.font = Font(color="FF0000")
        # except:
        #     print('Enter a job title')

    def updateFilePath(self):
        #try:
        #self.targetNameCol = self.targetNameColumn.text()
        global fileNameLabel
        fileNameLabel.setText('Current file:' + os.path.basename(filepath))
        print(fileNameLabel)


    def targetNameColumnConfig(self):
        #try:
        if targetNameColumn.text() != '':  # and titleBefore != title:
            #self.targetNameID = self.targetNameColumn.text()
            #print('target name column empty')
            global targetnamecolid
            targetnamecolid = targetNameColumn.text()
            print(targetnamecolid)
        #self.targetNameCol = self.targetNameColumn.text()
        self.nameTargetCell = targetnamecolid # + str(MainWindow.count + 1)
        print(self.nameTargetCell)



    def targetTitleColumnConfig(self):
        #try:
        if targetTitleColumn.text() != '':  # and titleBefore != title:
            global targettitlecolid
            targettitlecolid = targetTitleColumn.text()
            print(targettitlecolid)
        #self.targetCompanyCol = self.targetCompanyColumn.text()
        self.titleTargetCell = targettitlecolid + str(count + 1)
        print(self.titleTargetCell)

    def updateCompanyColumnConfig(self):
        if newCompanyNameColumn.text() != '':  # and titleBefore != title:
            global newCompanyNameCol
            newCompanyNameCol = newCompanyNameColumn.text()
            print(newCompanyNameCol)
            print(newCompanyNameColumn.text())
        global companyMsg
        global newCompanyName
        newCompanyName = companyMsg.text()
        print('this is' + newCompanyNameCol)
        self.newCompanyTargetCell = newCompanyNameCol + str(count + 1)
        sheet_obj[self.newCompanyTargetCell].value = newCompanyName
        companyTopVar = newCompanyNameCol + '1'
        sheet_obj[companyTopVar].value = 'New Company'
        print('Updating Company Column')

    def updateJobColumnConfig(self):
        if newTitleColumn.text() != '':  # and titleBefore != title:
            #global newTitleNameCol
            global newTitleNameCol
            newTitleNameCol = newTitleColumn.text()
            print(newTitleColumn.text())
            print('this is new title name column' + newTitleNameCol)
        #global title
        global jobMsg
        global newTitleName
        newTitleName = jobMsg.text()
        print(newTitleName)
        self.newTitleTargetCell = newTitleNameCol + str(count + 1)
        sheet_obj[self.newTitleTargetCell].value = newTitleName
        titleTopVar = newTitleNameCol + '1'
        sheet_obj[titleTopVar].value = 'New Title'
        #print(self.titleTopVar)
        #sheet_obj[self.titleTopVar].value = 'New Title'
        #self.titleTopVar = self.titleCol + '1'
        print('Enter a job title')

    def updateCount(self):
        try:
            global count
            count = (int(startAt.text()))
            #startAt.setText('')
            print('Count updated')
        except:
            print('Enter a number into count')

    def _prepNewCols(self):
        global newCompanyNameCol
        global newTitleNameCol
        companyTopVar = newfirmcol + '1'
        sheet_obj[companyTopVar].value = 'New Company'
        titleTopVar = newtitlecol + '1'
        sheet_obj[titleTopVar].value = 'New Title'

    def _askForNewFile(self):
        try:
            global filepath
            filepath = askopenfilename()
            print(filepath)
        except:
            filepath = './Example Excel (completely random profiles).xlsx'
            print('Bad filepath')

    def _connectSignals(self):
        """Connect signals and slots."""
        selectFile.clicked.connect(self._askForNewFile)

        startAt.returnPressed.connect(self.updateCount)

        targetNameColumn.returnPressed.connect(self.targetNameColumnConfig)
        targetCompanyColumn.returnPressed.connect(self.targetCompanyColumnConfig)
        targetTitleColumn.returnPressed.connect(self.targetTitleColumnConfig)

        newCompanyNameColumn.returnPressed.connect(self.updateCompanyColumnConfig)
        newTitleColumn.returnPressed.connect(self.updateJobColumnConfig)

        self.savebtn.clicked.connect(self.updateFilePath)
        self.savebtn.clicked.connect(_getData)
        self.savebtn.clicked.connect(_prepData)
        self.savebtn.clicked.connect(_prepExcel)
        self.savebtn.clicked.connect(setNewConfigData)


class MainWindow(QMainWindow):
    # global lineEditHeight
    # global lineEditWidth
    lineEditHeight = 30
    lineEditWidth = 125
    #count = int(user_info["countsaved"])
    run_once = 0

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background-color: #FFFFFF;")
        # self._view = view
        # self._connectSignals()
        self.w = AnotherWindow()
        self.dlgLayout = QVBoxLayout()
        self.formLayout = QFormLayout()
        self.generalLayout = self.dlgLayout
        self.dlgLayout.addLayout(self.formLayout)
        self._centralWidget = QWidget(self)
        self.setCentralWidget(self._centralWidget)
        self._centralWidget.setLayout(self.generalLayout)
        self._createMenu()
        self._createToolBar()
        self._createButtons()
        self._createLabels()
        self._createInputs()
        self.organizeText()
        # self._getData()
        # self._prepExcel()

        #self._view = view
        # Connect signals and slots
        self._connectSignals()



    def _createMenu(self):
        self.menu = self.menuBar().addMenu("&Menu")
        self.menu.addAction('&Exit', self.close)
        #self.menu.addAction('&Home', self.close)
        #self.menu.addAction('&Config', self.show_new_window)
        self.menu.addAction('&Config', self.toggle_window)


        #self.setCentralWidget(QPushButton)

    def _createToolBar(self):
        tools = QToolBar()
        self.addToolBar(tools)
        tools.addAction('Exit', self.close)
        #tools.addAction('&Home', self.close)
        #tools.addAction('&Config', self.show_new_window)
        tools.addAction('&Config', self.toggle_window)
        #self.setCentralWidget(QPushButton)
        tools.setStyleSheet("color: white;"
                                "background-color: #0A66C2;"
                                "border-style: solid;"
                                "border-width: 2px;"
                                "font-weight: 600;"
                                "font-size: 12px;"
                                "padding-bottom:2px;"
                                "border-color: #0A66C2;")

    def _createLabels(self):
        self.nameMsg = QLabel('')
        self.nameMsg.setStyleSheet("border-bottom: 2px solid black;")
        self.countMsg = QLabel('')
        self.formLayout.addWidget(self.nameMsg)
        self.formLayout.addWidget(self.countMsg)
        self.nameMsg.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        self.countMsg.setStyleSheet("border-bottom: 2px solid black;")
        self.countMsg.setFixedSize(self.lineEditWidth, self.lineEditHeight)






    def _createInputs(self):
        """Create the display."""
        # self.companyColumn = QLineEdit('')
        # self.titleColumn = QLineEdit('')
        # self.startAt = QLineEdit('')
        # self.startAt.setFocus()
        global companyMsg
        companyMsg = QLineEdit('')
        companyMsg.setStyleSheet(#"border: 2px solid  #007AFF;"
                                "background-color: #EEF3F8;"
                                 "border-radius: 6px;")
        global jobMsg
        jobMsg = QLineEdit('')
        jobMsg.setStyleSheet(#"border: 2px solid  #007AFF;"
                                "background-color: #EEF3F8;"
                                 "border-radius: 6px;")

        # self.companyColumn.setFixedSize(int(self.lineEditWidth / 3), self.lineEditHeight)
        # self.titleColumn.setFixedSize(int(self.lineEditWidth / 3), self.lineEditHeight)
        # self.startAt.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        companyMsg.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        jobMsg.setFixedSize(self.lineEditWidth, self.lineEditHeight)

        # self.generalLayout.addWidget(self.companyColumn)
        # self.generalLayout.addWidget(self.titleColumn)
        # self.generalLayout.addWidget(self.startAt)
        self.generalLayout.addWidget(companyMsg)
        self.generalLayout.addWidget(jobMsg)

        # self.display.setReadOnly(True)
        # Add the display to the general layout
        # self.generalLayout.addWidget(self.display)

    def _createButtons(self):
        """Create the buttons."""
        self.btn = QPushButton('LinkedIn Search')
        self.btn.setFont(QFont('Arial', 10))
        self.btn.setStyleSheet(
            "color: white;"
            "background-color: #0A66C2;"
            "border-style: solid;"
            "border-width: 2px;"
            "font-weight: 600;"
            "border-color: #0A66C2;"
            "border-radius: 3px")
        self.configbtn = QPushButton('Config')
        self.configbtn.setFont(QFont('Arial', 10))
        self.google = QPushButton('Search')
        self.google.setFont(QFont('Arial', 10))
        self.backbtn = QPushButton('Go Back')
        self.backbtn.setFont(QFont('Arial', 10))
        self.backbtn.setStyleSheet(
            "padding: 2px 5px 2px 5px;"
            "color: white;"
            "background-color: #0A66C2;"
            "border-style: solid;"
            "border-width: 2px;"
            "font-weight: 600;"
            "border-color: #0A66C2;"
            "border-radius: 3px")
        self.google.setStyleSheet(
            "color: white;"
            "background-color: #0A66C2;"
            # "color: #DB4437;"
            "border-style: solid;"
            "border-width: 2px;"
            "font-weight: 600;"
            "border-color: #0A66C2;"
            "border-radius: 3px")
        self.btn.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        self.backbtn.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        self.google.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        self.btn.setFixedSize(self.lineEditWidth, self.lineEditHeight)
        self.formLayout.addWidget(self.btn)
        self.formLayout.addWidget(self.google)
        self.formLayout.addWidget(self.backbtn)
        #self.backbtn.move(500,-50)
    # def setInputText(self, text):
    #     """Set display's text."""
    #     self.formLayout.addRow('Go Back:', self.backbtn)
    #     self.formLayout.addRow('Company Column:', self.ompanyColumn)
    #     self.formLayout.addRow('Title Column:', self.titleColumn)
    #     self.formLayout.addRow('Start Search:', self.btn)
    #     self.formLayout.addRow('Start At:', self.startAt)
    #     self.formLayout.addRow('Name:', self.nameMsg)
    #     self.formLayout.addRow('Company Name:', self.companyMsg)
    #     self.formLayout.addRow('Job:', self.jobMsg)
    #     self.formLayout.addRow('Count:', self.countMsg)
    #     self.formLayout.addRow('Google Search:', self.google)


    def organizeText(self):
        """Set display's text."""
        self.formLayout.addWidget(self.backbtn)
        self.backbtn.move(50, 0)
        #self.generalLayout.addWidget(self.backbtn)
        # self.formLayout.addRow('Company Column:', self.companyColumn)
        # self.formLayout.addRow('Title Column:', self.titleColumn)
        self.startLabel = QLabel('Start Search:')
        self.startLabel.setStyleSheet("color:black;font-weight: 600;")
        self.startLabel.setFont(QFont('Arial', 10))
        self.formLayout.addRow(self.startLabel, self.btn)
        # self.formLayout.addRow('Start At:', self.startAt)
        self.nameLabel = QLabel('Candidate Name:')
        self.nameLabel.setStyleSheet("color:black;font-weight: 600;")
        self.nameLabel.setFont(QFont('Arial', 10))
        self.formLayout.addRow(self.nameLabel, self.nameMsg)
        self.firmLabel = QLabel('Firm Name:')
        self.firmLabel.setStyleSheet("color:black;font-weight: 600;")
        self.firmLabel.setFont(QFont('Arial', 10))
        self.formLayout.addRow(self.firmLabel, companyMsg)
        self.jobLabel = QLabel('Candidate Title:')
        self.jobLabel.setStyleSheet("color:black;font-weight: 600;")
        self.jobLabel.setFont(QFont('Arial', 10))
        self.formLayout.addRow(self.jobLabel, jobMsg)
        self.countLabel = QLabel('Count:')
        self.countLabel.setStyleSheet("color:black;font-weight: 600;")
        self.countLabel.setFont(QFont('Arial', 10))
        self.formLayout.addRow(self.countLabel, self.countMsg)
        self.googleLabel = QLabel('Google Search:')
        self.googleLabel.setStyleSheet("color:black;font-weight: 600;")
        self.googleLabel.setFont(QFont('Arial', 10))
        self.formLayout.addRow(self.googleLabel, self.google)

        #global statusLabel
        #statusLabel = QLabel('Status:')
        self.TipsLabel = QLabel('Tip: Include firm name next to each')
        self.Tips2Label = QLabel('candidate in the Excel sheet')
        #statusLabel.setStyleSheet("color:black;font-weight: 600;")
        #statusLabel.setFont(QFont('Arial', 10))
        self.TipsLabel.setStyleSheet("color:black;font-weight: 600;")
        self.TipsLabel.setFont(QFont('Arial', 10))
        #self.generalLayout.addWidget(statusLabel)
        self.generalLayout.addWidget(self.TipsLabel)
        self.generalLayout.addWidget(self.Tips2Label)
        #self.formLayout.addRow(self.TipsLabel, self.Tips2Label)
        self.Tips2Label.setStyleSheet("color:black;font-weight: 600;")
        self.Tips2Label.setFont(QFont('Arial', 10))
        # self.formLayout.addRow(, '')
    def toggle_window(self):
        if self.w.isVisible():
            self.w.hide()
        else:
            self.w.show()

    def nameStatus(self):
        def splitName(cell):
            newCell = cell.split()
            print(newCell)
            separator = '+'
            joined = separator.join(newCell)
            print(joined)
            return cell
        # try:
        try:

            global name
            self.nameMsg.setText(str(name))
            print(name)
            self.nameGoogle = splitName(str(name))
            print(self.nameGoogle)
        except:
            print('No name')
        # except:
        #     self.nameMsg.setText('Empty cell, skip to next')
        #     print('Empty cell, skip to next')

    def companyNameStatus(self):
        # try:
        def splitName(cell):
            newCell = cell.split()
            print(newCell)
            separator = '+'
            joined = separator.join(newCell)
            print(joined)
            return cell
        #self.companyBefore = companyName
        companyMsg.setText(str(companyName))
        self.companyGoogle = splitName(str(companyName))

        # except:
        #     companyMsg.setText('Empty cell, skip to next')
        #     print('Empty cell, skip to next')

    def job(self):
        #try:
            #global titleBefore
            #self.titleBefore = title
        try:
            jobMsg.setText(title)
            print('this is title'+title)
        except:
            print('No job')
            #companyMsg.setText(self.copyItem)
        # except:
        #     jobMsg.setText('Empty cell, skip to next')
        #     print('Empty cell, skip to next')

    def counter(self):
        try:
            self.counter = str(count+1)
            self.countMsg.setText(self.counter)
            print('this is count' + self.counter)
            print(count)
        except:
            self.countMsg.setText('Empty cell, skip to next')
            print('Empty cell, skip to next')

    def goback(self):
        global count
        count -= 2
        print(count)

    def startSearch(self):
        #try:
        def iterate():
            # print('Name:')
            # print(self.name)
            # print('Count out of 2100:')
            # print(self.count)
            # print(self.title)


            #print(names[3])
            def linkedinsearch_item(search_query):
                webbrowser.open("https://www.linkedin.com/search/results/all/?keywords=" +
                                str(search_query) + "&origin=GLOBAL_SEARCH_HEADER&sid=(s5")
            linkedinsearch_item(name)
            # global companyDetector
            # self._view.setDisplayText()

        # filepath = r'C:\WellesleyAutoSearch-Wellesley\datasets\data3.csv'
        # filepath = '/Users/jamesloh/PycharmProjects/SeleniumTest/datasets/data2.csv'
        # filepath = '~/datasets/data1.csv'
        #
        # if pandas.isnull(newCompanyName):
        #     print(newCompanyName)
        #     self.count += 1
        #     iterate()
        # else:
        print('Company name has been copied to your clipboard')
        global companyName
        global count
        global targetfirmcolid
        if type(companyName) == str:
            try:
                self.copyItem = companyName
                pyperclip.copy(' ' + self.copyItem)
                print(self.copyItem)
            except:
                # try:
                #     pyperclip.copy(' ' + str(data.loc[count - 1, targetCompanyColumn.text()]))
                #     # print('this is copy item' + self.copyItem)
                #     print(data.loc[count, targetCompanyColumn.text()])
                # except:
                pyperclip.copy(' ' + str(data.loc[count, user_info["targetfirmcolid"]]))
                # print('this is copy item' + self.copyItem)
                print(data.loc[count, user_info["targetfirmcolid"]])
        else:
            print("No copy")

        # try:
        #     self.copyItem = companyName
        #     pyperclip.copy(' ' + self.copyItem)
        #     #print('this is copy item' + self.copyItem)

        count += 1
        print(count)
        iterate()
        # except:
        #     print('Incomplete, please redo')
        #     print(filepath)


    def googleSearch(self):
        try:
            def googlesearch_item(search_query):
                webbrowser.open("https://www.google.com/search?client=firefox-b-d&q=" +
                                str(search_query))
            searchItem = self.nameGoogle + ' ' + self.companyGoogle
            googlesearch_item(str(searchItem))
        except:
            print('Please LinkedIn search first')

    # def updateCompanyColumnConfig(self):
    #     if self.newCompanyNameColumn.text() != '':  # and titleBefore != title:
    #         global newCompanyNameCol
    #         newCompanyNameCol = self.newCompanyNameColumn.text()
    #         print(newCompanyNameCol)
    #     global companyMsg
    #     self.newCompanyName = companyMsg.text()
    #     self.newCompanyTargetCell = newCompanyNameCol + str(MainWindow.count + 1)
    #     sheet_obj[self.newCompanyTargetCell].value = self.newCompanyName
    #     self.companyTopVar = newCompanyNameCol + '1'
    #     sheet_obj[self.companyTopVar].value = 'New Company'
    #     print('Updating Company Column')

    def updateCompanyName(self):
        #self.companyVar2 = 'A' + str(self.count + 2)
        print('Updating Company Name')
        newCompanyName = companyMsg.text()
        #global newCompanyNameCol
        self.updateNewCompanyTargetCell = newfirmcol + str(count + 1)
        # sheet_obj[companyVar].value = copyItem
        #self.font = Font(color="FF0000")
        #sheet_obj[self.updateNewCompanyTargetCell].font = self.font
        sheet_obj[self.updateNewCompanyTargetCell].value = newCompanyName
        #sheet_obj[AnotherWindow.companyTopVar].value = 'New Company'
        global wb
        wb.save(filepath)

    def updateJobName(self):
        print('New title saved')
        #global title
        title = jobMsg.text()
        self.updateNewTitleTargetCell = newtitlecol + str(count + 1)
        print(self.updateNewTitleTargetCell)
        sheet_obj[self.updateNewTitleTargetCell].font = self.font
        sheet_obj[self.updateNewTitleTargetCell].value = title
        #sheet_obj[self.updateNewTitleTargetCell].value = 'New Title'
        wb.save(filepath)

    # def updateCompanyName(self):
    #     if self.companyColumn.text() != '':  # and titleBefore != title:
    #         companyCol = self.companyColumn.text()
    #     print('Updating Company Name')
    #     self.newCompanyName = self.companyMsg.text()
    #     self.companyVar = companyCol + str(self.count + 1)
    #     self.companyVar2 = 'A' + str(self.count + 2)
    #     self.companyTopVar = companyCol + '1'
    #     print(self.newCompanyName)
    #     # sheet_obj[companyVar].value = copyItem
    #     self.font = Font(color="FF0000")
    #     self.sheet_obj[self.companyVar].value = self.newCompanyName
    #     self.sheet_obj[self.companyVar].font = self.font
    #     self.sheet_obj[self.companyTopVar].value = 'New Company'
    #     # except:
    #     #     print('Enter a company name')
    #
    # def updateJobName(self):
    #     try:
    #         if self.titleColumn.text() != '':  # and titleBefore != title:
    #             titleCol = self.titleColumn.text()
    #         self.title = self.jobMsg.text()
    #         self.titleVar = titleCol + str(self.count + 1)
    #         self.titleTopVar = titleCol + '1'
    #         self.sheet_obj[self.titleVar].value = self.title
    #         self.sheet_obj[self.titleVar].font = self.font
    #         self.sheet_obj[self.titleTopVar].value = 'New Title'
    #         print(self.titleVar)
    #         self.wb.save(self.filepath)
    #         print('New title saved')
    #     except:
    #         print('Enter a job title')

    def _connectSignals(self):
        """Connect signals and slots."""
        # self._getData()
        # self._prepExcel()

        self.btn.clicked.connect(AnotherWindow.updateFilePath)
        self.btn.clicked.connect(_getData)
        self.btn.clicked.connect(_prepExcel)
        self.btn.clicked.connect(_prepData)
        self.btn.clicked.connect(self.nameStatus)
        self.btn.clicked.connect(self.companyNameStatus)
        self.btn.clicked.connect(self.job)
        self.btn.clicked.connect(self.counter)
        self.btn.clicked.connect(self.startSearch)

        self.google.clicked.connect(self.googleSearch)
        self.backbtn.clicked.connect(self.goback)
        self.backbtn.clicked.connect(_prepData)
        self.backbtn.clicked.connect(self.startSearch)
        self.backbtn.clicked.connect(self.nameStatus)
        self.backbtn.clicked.connect(self.companyNameStatus)
        self.backbtn.clicked.connect(self.job)
        self.backbtn.clicked.connect(self.counter)

        # self.startAt.returnPressed.connect(self.updateCount)
        companyMsg.returnPressed.connect(self.updateCompanyName)
        jobMsg.returnPressed.connect(self.updateJobName)



def main():
    """Main function."""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    window.setWindowTitle('Wellesley LinkedIn Searcher')
    window.setFixedSize(300, 360)
    app.exec()


#getConfigData()
getFileName()
_getData()
_prepExcel()
_prepData()
#getDataFromConfig()
main()

