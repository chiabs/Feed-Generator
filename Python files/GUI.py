from PyQt5.QtWidgets import QSizePolicy, QMenuBar
from PyQt5.QtGui import * 
from  ExcelReader import Reader
from Writer import Writer

from PyQt5.QtWidgets import (
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QFormLayout,
    QDialog,
    QMessageBox
)

#Main Window class to select type of process
class MainWindow(QDialog):
    def __init__(self):
        super().__init__(parent=None)
        self.setWindowTitle("CC/PCard Feed Generator")
        # setting geometry
        self.setGeometry(100, 100, 600, 500)
        dialogLayout = QVBoxLayout()

        # create menu
        menubar = QMenuBar()
        dialogLayout.addWidget(menubar)
        helpMenu = menubar.addMenu("&Help")
        helpAction = helpMenu.addAction("Help")
        helpAction.triggered.connect(launchMainHelpWindow)

        #New button triggers createWindow
        self.newbutton = QPushButton("")
        self.newbutton.setObjectName("Create New Feed")
        self.newbutton.setText("Create New Feed")
        self.newbutton.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.newbutton.setFont(QFont('Times font',15))
        self.newbutton.clicked.connect(self.launchCreateWindow)

        #Saved button triggers saveWindow
        self.savedbutton = QPushButton("")
        self.savedbutton.setObjectName("Create from Excel")
        self.savedbutton.setText("Create from Excel")
        self.savedbutton.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.savedbutton.setFont(QFont('Times font',15))
        self.savedbutton.clicked.connect(self.create_from_excel)

        dialogLayout.addWidget(self.newbutton)
        dialogLayout.addWidget(self.savedbutton)

        self.setLayout(dialogLayout)

    def launchCreateWindow(self):
        self.createWindow = CreateWindow()
        self.createWindow.setFont(QFont("Times font", 12))
        self.createWindow.show()
    
    #Function to create from excel
    def create_from_excel(self):
        generator_log = Reader.read_file()

        if generator_log == ["Please ensure that feed.xlsx is closed before trying again."]:
            showAlert(generator_log[0],"alert")
        else:
            showAlert(generator_log, "detail")
    
    

        

#Window for creation via data entry in app
class CreateWindow(QWidget): 
    def __init__(self):
        super().__init__(parent=None)
        self.setWindowTitle("CC/PCard Feed Generator")
        #self.setGeometry(100, 100, 400, 800)
        dialogLayout = QVBoxLayout()
        formLayout = QFormLayout()

        # create menu
        menubar = QMenuBar()
        dialogLayout.addWidget(menubar)
        helpMenu = menubar.addMenu("&Help")
        helpAction = helpMenu.addAction("Help")
        helpAction.triggered.connect(launchCreateHelpWindow)

        #For all names and fields
        nameLabel = QLabel('Employee Name*')
        nameLabel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.nameField = QLineEdit()
        self.nameField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        quantityLabel = QLabel('Number of Expenses')
        self.quantityField = QLineEdit()
        self.quantityField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        outputFileLabel = QLabel('Output File Name')
        self.outputFileField = QLineEdit()
        self.nameField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        cardLabel = QLabel('Card Number*')
        self.cardField = QLineEdit()
        self.cardField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        idLabel = QLabel('Employee ID*')
        self.idField = QLineEdit()
        self.idField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        currencyLabel = QLabel('Currency*')
        self.currencyField = QLineEdit()
        self.currencyField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        dateLabel = QLabel('Date (YYYYMMDD)')
        self.dateField = QLineEdit()
        self.dateField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        dateLabel = QLabel('Date (YYYYMMDD)')
        self.dateField = QLineEdit()
        self.dateField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        minimumLabel = QLabel('Minimum Amount')
        self.minimumField = QLineEdit()
        self.minimumField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        maximumLabel = QLabel('Maximum Amount')
        self.maximumField = QLineEdit()
        self.maximumField.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        #Add into the layout
        formLayout.addRow(nameLabel, self.nameField)
        formLayout.addRow(idLabel, self.idField)
        formLayout.addRow(cardLabel, self.cardField)
        formLayout.addRow(quantityLabel, self.quantityField)
        formLayout.addRow(outputFileLabel, self.outputFileField)
        formLayout.addRow(currencyLabel, self.currencyField)
        
        formLayout.addRow(dateLabel, self.dateField)
        formLayout.addRow(minimumLabel, self.minimumField)
        formLayout.addRow(maximumLabel, self.maximumField)
        dialogLayout.addLayout(formLayout)

        #Generate button triggers process
        self.button = QPushButton("")
        self.button.setObjectName("Generate")
        self.button.setText("Generate")
        self.button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.button.clicked.connect(self.create_feed)
        dialogLayout.addWidget(self.button)

        #Reset button clears the fields
        self.resetbutton = QPushButton("")
        self.resetbutton.setObjectName("Reset")
        self.resetbutton.setText("Reset")
        self.resetbutton.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.resetbutton.clicked.connect(self.reset)
        dialogLayout.addWidget(self.resetbutton)

        #For the generated Label
        self.msgLabel = QLabel("All * fields are compulsory")
        dialogLayout.addWidget(self.msgLabel)

        #Set layout
        self.setLayout(dialogLayout)


    #Function to create feed
    def create_feed(self):
        filename = self.outputFileField.text() 
        quantity = self.quantityField.text()
        card = self.cardField.text()
        employee_id = self.idField.text()
        employee_name = self.nameField.text()
        currency = self.currencyField.text()
        currency = currency.upper()
        date = self.dateField.text() 
        minimum = self.minimumField.text()
        maximum = self.maximumField.text()
        #Initialise Writer object
        write = Writer(card,employee_id,employee_name,currency,filename,quantity,date,minimum,maximum)
        #Call create_feed functions which returns a string
        result = write.create_feed()

        #Create alert dialog with result
        if "Generated" in result:
            showAlert(result,"info")
        else:
            showAlert(result,"alert")
    
    #Reset the fields
    def reset(self):
        self.outputFileField.setText("")
        self.cardField.setText("")
        self.nameField.setText("")
        self.idField.setText("")
        self.quantityField.setText("")
        self.currencyField.setText("")
        self.dateField.setText("")
        self.msgLabel.setText("Resetted!")

    #Save information
    def save(self):
        card = self.cardField.text()
        employee_id = self.idField.text()
        employee_name = self.nameField.text()
        currency = self.currencyField.text()
        currency = currency.upper()

        datalist = [employee_name, employee_id, card, currency]

        if card == "" or employee_id == "" or employee_name == "" or currency == "":
            showAlert("Please fill in currency, card, employee ID and employee name to save.")
        else:
            #Storage.Storage.save(datalist)
            showAlert("Saved successfully","info")

#Custom messagebox class for displaying alerts
class MessageBox(QMessageBox):
    def __init__(self, parent=None):
        super().__init__(parent) 
    
    def setAlertMsg(self,string):
        self.setWindowTitle("Error")
        self.setIcon(QMessageBox.Critical)
        self.setText("")
        self.setStandardButtons(QMessageBox.Ok)
        self.setText(string)
    
    def setInfoMsg(self,string):
        self.setWindowTitle("Information")
        self.setIcon(QMessageBox.Information)
        self.setText("")
        self.setStandardButtons(QMessageBox.Ok)
        self.setText(string)
    
    def setInfoDetailMsg(self,list):
        #Initialise counter
        success = 0
        failure = 0
        details_string = ""

        #Record success and failure, compile details into single string
        for item in list:
            if "Success" in item:
                success += 1
            else:
                failure += 1
            details_string += item + "\n"

        self.setIcon(QMessageBox.Information)
        self.setText("Successfully generated " + str(success) + " files. " + str(failure) + " failures.")
        self.setInformativeText("Click details to find out more.")
        self.setWindowTitle("Generation information")
        self.setDetailedText(details_string)
        self.setStandardButtons(QMessageBox.Ok)


#Shows error messages thru a new message box
def showAlert(MsgStr, type):
    dlg = MessageBox()
    if type == "info":
        dlg.setInfoMsg(MsgStr)
    elif type == "detail":
        dlg.setInfoDetailMsg(MsgStr)
    else:
        dlg.setAlertMsg(MsgStr)
    dlg.exec()


#Launch main help window
def launchMainHelpWindow():
    msgBox = QMessageBox()
    msgBox.setWindowTitle("Information")
    msgBox.setIcon(QMessageBox.Information)
    msgBox.setText("")
    msgBox.setStandardButtons(QMessageBox.Ok)

    helpstring = "Click on 'Create New Feed' if you wish to launch the create menu which will allow you to input your parameters for a credit card feed." + "\n" + "\n" + "Click on 'Create from Excel' to create feed(s) from data in feed.xlsx file, which has been automatically generated in the Excel_Feed folder." + "\n" + "\n" + "Before clicking 'Create from Excel', please ensure that you have filled in the necessary details: Name, ID, Card, Currency, Quantity and ensure that you have saved anc closed the file." + "\n" + "\n" + "The generated output will be located in the Generated_Files folder."
    msgBox.setText(helpstring)

    msgBox.exec()


#Launch create help window
def launchCreateHelpWindow():
    msgBox = QMessageBox()
    msgBox.setWindowTitle("Information")
    msgBox.setIcon(QMessageBox.Information)
    msgBox.setText("")
    msgBox.setStandardButtons(QMessageBox.Ok)
    helpstring = "test"
    helpstring = "Please fill in compulsory fields: Employee Name, Employee Id, Card Number, Currency." + "\n" + "\n" + "For the optional fields, these are the default values if left empty:" + "\n" + "Number of Expense : 5"  + "\n" + "Output File name : output" + "\n" + "Date : Current date" + "\n" + "Minimum Amount : 0" + "\n" + "Maximum Amount : 999" + "\n" + "\n" + "The generated output will be located in the Generated_Files folder."
    msgBox.setText(helpstring)

    msgBox.exec()



