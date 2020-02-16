'''
User Hour Counter
'''
#Run Backup file
import Backup

#Excel/Python Module
import xlsxwriter 
import xlrd
import pandas as pd
from time import sleep
from subprocess import call 
import sys


#os module
import os


# Worksheet data list
userIDS = []
hours = []


def start(): 
    print("Program Starting...")
    sleep(0.7)
    #Create Backup
    Backup.main()
    #Open txt files  
    id = open('ID# List.txt','r+')
    hrs = open('Hour List.txt','r+')
    try:
        #Add txt data to list in program
        for line in id:
            userIDS.append(int(line.strip()))
        for line in hrs:
            hours.append(int(line.strip()))
    except ValueError:
        runBackup(hrs)
    print("IDs File Opened")
    print("Success.")
    sleep(0.4)
    #Return data
    return id, hrs

def clear():
    print("\n-----------------")
    print("Clearing console")
    print("-----------------")
    sleep(1.75)
    #Clear console so it's not as crammed 
    if os.name == 'nt': 
        _ = os.system('cls')
  
def userInput(id, hrs):
    valid = ""
    #Valid value not set as data is not yet checked
    uID = 0
    while valid != "y":
        try:
            valid = ""
            #Program try asking user for data input
            clear()
            uID1 = input("ID #: ")
            uID2 = int(abs(int(uID1)))
        #If data inputed is not the correct type of data inform user
        except ValueError:
            if uID1 == ".":
                emClose(id, hrs)
                sys.exit("Program Stopped.")
            else:
                print('Invalid')
                valid = "inv"
        #Can't input for 0
        if uID == 0:
            print('Invalid')
            valid = "inv"
        if len(str(uID)) != 7:
            print("Invalid")
            valid = "inv"
        #Ask user if their input is correct
        if valid != "inv":
            valid = str(input("Is " + str(uID) + " Correct? (y/n)"))
    #Return user verified input
    return uID


def userIDAdd(uID, id, hrs):
    valid = 0
    #Valid loop
    while valid != 1:
        #Ask user for new ID
        print("New ID = " + str(uID))
        newID = uID
        #Check if length of ID is 7 characters
        if len(str(newID)) == 7:
            valid += 1
        else:
            print("Invalid ID")
    #Add new data to list and doccuments
    userIDS.append(newID)
    hours.append(1)
    id.writelines(str(newID) + " \n")
    hrs.writelines(str(1) + " \n")
    

def uIDCheck(uID, addID, hrs):
    valid = 0
    #Check to see if input by user is already been logged
    for loggedID in userIDS:
        if uID == loggedID:
            valid += 2
    #If input has not been logged before
    if valid != 2:
        #Ask user if they want to register ID
        addID = input(str(uID) + " || ID Not Registered | Add (y/n) ?: ")
        #If yes
        if addID == 'y':
            #Method to log new ID
            userIDAdd(uID, id, hrs)
        #If no
        elif addID == 'n':
            #Prompt user to try again
            print("Try again")
            valid += 1
    return uID, valid


def uData(id, hrs):
    valid = 0
    #Valid used loop program until data is correct and verified
    while valid != 2:
        #Method to get user inputs
        uID = userInput(id, hrs)
        checkeduID, valid = uIDCheck(uID, id, hrs)
    return checkeduID


def hourAdd(ID, hrs):
    #X value set to list place of ID
    x = userIDS.index(ID)
    #Add an hour to the current value of same x value in hours list 
    sHrs = int(hours[x])
    sHrs += 1
    hours[x] = str(sHrs)
    #Method to clear document to log new hours
    deleteContent(hrs)
    #Rewrite hours.txt with new hour data
    for i in hours:
        hrs.writelines(str(i) + " \n")


def deleteContent(pfile):
    #Deletes contents in a docuement
    pfile.seek(0)
    pfile.truncate()


def excelHourLog():
    x = 0
    #Open excel doccoment
    book = xlrd.open_workbook("Info.xlsx")
    first_sheet = book.sheet_by_index(0)
    #If the amount of ID's is more than the amount of hours logged
    while len(userIDS) > len(hours):
        #Add new value found to corresponding value from excel in list
        cell = first_sheet.cell(x, 0)
        hours.append(cell.value)


def excelIDLog():
    #Create a Pandas dataframe from the data.
    df1 = pd.DataFrame({'ID': userIDS,
                        'Hours': hours})
    #Create a Pandas Excel writer using XlsxWriter as engine
    writer = pd.ExcelWriter('Info.xlsx', engine='xlsxwriter')
    #Convert the dataframe to an XlsxWriter Excel object
    df1.to_excel(writer, sheet_name='hourData')
    return writer

def runBackup(hrs):
    bakup = open("backups/backupQuantity.txt", "r+")
    for line in bakup:
        hours.append(line.strip())
    x = int(hours[-1])
    lastestBackup = "backup_" + str(x) + ".txt"
    backedHours = open("backups/" + lastestBackup, "r+")
    hrs.writelines(backedHours)
    bakup.close()
    backedHours.close()

def excelLogging():
    #Run logging method
    excelHourLog()
    writer = excelIDLog()
    #Save file
    writer.save()
    return writer

def program(hrs):
    dne = " "
    #While dne value is not a ., keep looping program
    while dne != ".":
        #Try to run program
        try:
            #Method to request user data
            ID = uData(id, hrs)
            #Method to add hours to user
            hourAdd(ID, hrs)
            #Methods to log data into excel
            writer = excelLogging()
        #If Hour List.txt file deleted or corrupt, run past backup
        except ValueError:
            runBackup(hrs)
        except IndexError:
            runBackup(hrs)

        #Display hours to user
        x = userIDS.index(ID)
        print("\n" + str(ID) + ":")
        print("Current Hours: " + str(hours[x]) + "\n")

        #Create a backup
        Backup.main()

        #Ask user if done
        #dne close value set to "." to prevent users to close program accidentally
        dne = input("Press enter to continue. \n")
    print("Program Stopped")
    return writer


def emClose(id, hrs):
    writer = excelLogging()
    closeProg(id, hrs, writer)

def closeProg(id, hrs, writer):
    #Close and save all files
    id.close()
    hrs.close()
    #Create a backup
    Backup.main()
    writer.save()


if __name__ == "__main__":
    print("\n------------------\nHour Logger\nAdrian Rodriguez-Arana\n------------------\n")
    #Start Program Method
    id, hrs  = start()

    writer = program(hrs)

    #Close Program method
    closeProg(id, hrs, writer)
