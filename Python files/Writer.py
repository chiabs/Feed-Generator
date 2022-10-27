from random import randint
from datetime import date
import datetime
import os

class Writer:
    def __init__(self,card,employee_id,employee_name,currency,filename,quantity,date,minimum,maximum):
        self.card = card
        self.employee_id = employee_id
        self.employee_name = employee_name
        self.currency = currency
        self.filename = filename + ".txt"
        self.quantity = quantity
        self.date = date
        self.minimum = minimum
        self.maximum = maximum

    #Function to create feed
    def create_feed(self):
        card = self.card
        employee_id = self.employee_id
        employee_name = self.employee_name
        currency = self.currency
        filename = self.filename
        quantity = self.quantity
        date = self.date
        minimum = self.minimum
        maximum = self.maximum

        #Required data
        requiredList = [card,employee_id,employee_name,currency]
        #Autofill data if empty
        if date == "":
            today = datetime.datetime.today().strftime("%Y-%m-%d")
            date = today.replace("-","")
        if quantity == "":
            quantity = "5"
        if filename == ".txt":
            random_seq = str(randint(0,10000))
            filename = employee_name + "_" + random_seq + filename
        if minimum == "":
            minimum = "0"
        if maximum == "":
            maximum = "999"
        
        #First line
        first_line = "100" + card + employee_name + 13 * " " + employee_id

        #Data Validation
        if self.checkEmpty(requiredList): 
            return "Ensure that the following are filled: Card, Name, ID, Currency"
        elif not card.isnumeric():
            return "Please enter valid card number"
        elif not self.dateValidate(date):
            return "Please enter valid date in YYYYMMDD format"
        elif not quantity.isnumeric():
            return "Please enter valid number of expenses"
        elif not len(currency) == 3:
            return "Please enter valid currency (e.g. sgd, eur)"
        elif not employee_id.isnumeric():
            return "Please enter correct employee ID"
        elif self.has_numbers(employee_name):
            return "Employee name should not have numbers"
        elif not self.filenameValidate(filename):
            return "Filename should not contain | < > ? * /"
        elif not int(maximum) > int(minimum):
            return "Max amount must be greater than min amount"
        elif int(maximum) < 0 or int(minimum) < 0:
            return "Max/Min amounts cannot be negative"
        else:
            #Obtain path
            direct = os.path.dirname(__file__)
            direct = direct + "\\Generated_Files\\"

            #Check if directory exists else create folder
            if not os.path.exists(direct):
                os.mkdir(direct)

            #Open file and write
            with open(direct + filename, 'w') as f: 
                f = open(direct + filename, 'w')
                f.write(first_line)
                f.write('\n')
                f.write('\n')
                initial = randint(1000000000000000, 9999999999999999)
                count = 0

                quantity2 = int(quantity)
                for x in range(quantity2):
                    #Generate random amount from 100 to 500 up to 2 dp
                    amt = str(randint(int(minimum)*100,int(maximum)*100))
                    while len(amt) < 15:
                        amt = "0" + amt
                    attach = 34 * " "
                    attach2 = date + currency + amt + currency + amt + "Dummy Expense Desc and Expense Code"
                    f.write("200" + card + str(count + initial) + attach + date + attach2)
                    f.write('\n')
                    f.write('\n')
                    count += 1
                
            #Update in window to show completed feed 
            return "Generated file: " + filename

    #Function for date validation
    def dateValidate(self,date):
        if not len(date) == 8:
            return False
        
        isValidDate = True

        year = date[:4]
        month = date[4:6]
        day = date[6:]
        
        try:
            datetime.datetime(int(year), int(month), int(day))
        except ValueError:
            isValidDate = False

        return isValidDate

    #Check if string contains numbers
    def has_numbers(self,inputString):
        return any(char.isdigit() for char in inputString)
        
    #Validate file name
    def filenameValidate(self,name):
        invalidChar = ['|','/','*','?','>','<']
        for char in invalidChar:
            if char in name:
                return False
                
        return True

    #check if list has empty strings
    def checkEmpty(self,list):
        for str in list:
            if str == "":
                return True
        return False

