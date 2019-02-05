import csv
import os
import shutil
import tkinter
import tkinter.filedialog

def getPath():
    root = tkinter.Tk()
    root.withdraw()
    iDir = os.getcwd()
    Path = tkinter.filedialog.askdirectory(initialdir = iDir, title="select a directry you want to make a FileList.")
    return Path

def getFilelist(Path):
    NameList = os.listdir(Path)

    table1 =[]
    table1.extend([Path,"Files or Dir.s",len(NameList)])
    table2 =[]
    table2.extend(["No", "File Name", "Path", "Dir to copy"])

    mainlist= []
    for i in range(len(NameList)):
        No = i+1
        FileName = NameList[i]
        hyperlink = '=hyperlink("{0}/{1}")'.format(Path, FileName) 
        mainlist.append([No, FileName, hyperlink, ''])
    return [table1,table2,mainlist]

def writeFilelist(table1, table2, mainlist):
    root = tkinter.Tk()
    root.withdraw()
    with open("FileList.csv","w") as f:
        writer = csv.writer(f, lineterminator='\n')
        writer.writerow(table1)
        writer.writerow(table2)
        writer.writerows(mainlist)
    root.filename =  tkinter.filedialog.asksaveasfilename(initialdir = Path,title = "select a directry to save Filelist", initialfile = "FileList.csv", filetypes = (("csv","*.csv"),("all files","*.*")))
    shutil.move(os.getcwd() + '/FileList.csv', root.filename)

Path = getPath()
FileList = getFilelist(Path)
writeFilelist(FileList[0], FileList[1], FileList[2])