import sys
import os

from numpy import choose

applicationPath = os.environ['APPLICATION_PATH']

class Contract():   
    "To be moved to a utils folder for other apps"

    def __init__(self):
        Contract.filePath = "Has Not Been Updated"
        Contract.clear = lambda: os.system('cls')

    def ChooseFile(self, filePath, selectionType):
        directoryNames = []
        for (dirpath, dirnames, filenames) in os.walk(filePath):
                directoryNames.extend(filenames)
        choices = ""
        for index in range(len(directoryNames)):
            choices += (str(index)+" - "+directoryNames[index]+"\n")
        choices += (str(len(directoryNames))+" - New "+ selectionType + " \n")
        print(choices)
        print(len(directoryNames))
        choiceSelection = int(input("Selection: "))
        if choiceSelection == len(directoryNames):
            print("Making new "+selectionType)
        else:
            Contract.filePath = directoryNames[choiceSelection]
            print(Contract.filePath)

    def SelectContract(self):
        #Variables
        filePath = "\\Contracts Folder"
        memory = open(applicationPath+"\Data\contractMemory.txt", "r")
        
        #Recent Payment Choices
        selectionMemory = memory.read()
        selectionMemory = selectionMemory.split("\n")
        if len(selectionMemory) > 3:
            selectionMemory.pop()

        choices = ""

        for index in range(len(selectionMemory)):
                choices += (str(index)+" - "+selectionMemory[index].rsplit('/', 1)[-1]+"\n")
        choices += str(len(selectionMemory))+" - Other"
        print("Recently Selected Contracts")
        print(choices)
        selection = int(input("Selection: "))
        Contract.clear()

        #All contracts choices
        if selection == 3:

            for i in range(2):
                if i == 2:
                    selectionMemory.insert(0,filePath)
                    with open(applicationPath+"\Data\contractMemory.txt", "w") as file:
                        file_lines = "\n".join(selectionMemory)
                        file.write(file_lines)
                    

                directoryNames = []
                directoryPaths = []
                for (dirpath, dirnames, filenames) in os.walk(filePath):
                    directoryNames.extend(dirnames)
                    break
                
                if i == 0:
                    filteredDirectoryNames = [k for k in directoryNames if 'Contracts' in k]
                    directoryNames = filteredDirectoryNames

                choices = ""
                for index in range(len(directoryNames)):
                    choices += (str(index)+" - "+directoryNames[index]+"\n")
                print(choices)
                companySelection = int(input("Selection: "))
                try:
                    filePath = filePath+"/"+directoryNames[companySelection]
                except:
                    x=1
        else:
            filePath = selectionMemory[selection]
        print(filePath)
        Contract.filePath = filePath

if __name__ == '__main__':
    sys.exit()