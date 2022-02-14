import sys
import os


applicationPath = os.environ['APPLICATION_PATH']


import Utils as Utils

class InvoicingApplication:
    "Time to "
    def __init__(self):
        print("init")

    def CheckExisting():
        InvoicingApplication.selectedContract = Utils.Contract()
        InvoicingApplication.selectedContract.SelectContract()
        filePath = InvoicingApplication.selectedContract.filePath
        projectName = ((filePath).rsplit('/')[2])[5:]
        filePath += r"\Commercial\Valuation\Invoices"
        
        invoices = os.listdir(filePath)
        if len(invoices) > 0:
            print("Preparing Invoice "+str(len(invoices))+" for "+projectName)
        else:
            print("Preparing first invoice for"+projectName)
        

       

    def Main():
        InvoicingApplication.CheckExisting()



InvoicingApplication.Main()
    

