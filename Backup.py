'''
ExcelAuto Backups
'''
import os

backups = []

def importBackup():
    bkup = open("backups/backupQuantity.txt", "r+")
    for line in bkup:
        backups.append(int(line.strip()))
    return bkup
    bkup.close()

def createFile(bkup):
    backups.append(int(backups[-1]) + 1)
    bkup.writelines(str(backups[-1]) + " \n")
    x = int(backups[-1])
    bname = "backup_" + str(x) + ".txt"
    currentBkup = open("backups/" + bname, "w+")
    return currentBkup

def writeBackup(cB):
    currentHours = open("Hour List.txt", "r+")
    cB.writelines(currentHours)
    cB.close()


def main():
    bkup = importBackup()
    cB = createFile(bkup)
    writeBackup(cB)