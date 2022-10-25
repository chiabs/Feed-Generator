import xlsxwriter 
from Writer import Writer
import os
import pandas as pd


class Reader():
    def __init__(self):
        pass

    #Function to check if excel feed exists, if not it will create
    def create_file():
        direct = os.path.dirname(__file__)
        direct = direct + "\\Excel_Feed"
        
        if not os.path.exists(direct):
            os.mkdir(direct)

        if not os.path.isfile(direct + '\\feed.xlsx'):
            #Create xlsx file
            workbook = xlsxwriter.Workbook(direct + '\\feed.xlsx')
            worksheet = workbook.add_worksheet()

            #Write on worksheet
            worksheet.write('A1', 'Quantity')
            worksheet.write('B1', 'Name')
            worksheet.write('C1', 'ID')
            worksheet.write('D1', 'Card')
            worksheet.write('E1', 'Currency')
            worksheet.write('F1', 'Date')
            worksheet.write('G1', 'Minimum Amount')
            worksheet.write('H1', 'Maximum Amount')
            worksheet.write('I1', 'Filename')

            #Close workbook
            workbook.close()

    #Function to read excel file and create feed
    def read_file():
        path = os.path.dirname(__file__) + '\\Excel_Feed\\feed.xlsx'
        try:    
            df = pd.read_excel(path)
            df =df.fillna('')
            df_dict = df.to_dict(orient='records')

            generator_log = []

            for item in df_dict:
                quantity = str(item['Quantity'])
                if int(quantity) > 0:
                    name = str(item['Name'])
                    emp_id = str(item['ID'])
                    card = str(item['Card'])
                    currency = str(item['Currency'])
                    date = str(item['Date'])
                    min_amt = str(item['Minimum Amount'])
                    max_amt = str(item['Maximum Amount'])
                    filename = str(item['Filename'])
                    
                    #Initialise Writer object
                    write = Writer(card,emp_id,name,currency,filename,quantity,date,min_amt,max_amt)

                    #Call create_feed functions which returns a string
                    result = write.create_feed()

                    #Details of run to be added to log
                    details = ""

                    #Separate for success and failure cases
                    if "Generated" in result:
                        details = "Success: " + name + " " + card
                    else:
                        details = "Failure: " + name + " " + card + ". " + result
                    
                    #Append details to log
                    generator_log.append(details)

            return generator_log

        except:
            return ["Please ensure that feed.xlsx is closed before trying again."]

        



    def delete(credentials):
        pass


        






