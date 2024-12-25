import sys
''' 
    This sys.path.append() is for integration with IDEA purposes
    Make it point to the site packages directory in your
    Python virtual environment
    Comment it out if you're running it from terminal/Virtual Environment
'''
# sys.path.append('D:/MED_IDEA_Internsip/Managed_Projects/TestingCreation/Macros.ILB/virEnv/Lib/site-packages')
import tkinter as tk
import tkinter.font as tkFont
import pandas as pd
#import win32com.client as win32ComClient
import os
import time
import json
import traceback
from tkinter import *
from pathlib import PurePosixPath, Path
from tkinter import filedialog as fd
from edefter_clean_xbrl import clean_xbrl


class App:
    def __init__(self, root):
        # setting title
        root.title("E-Defter Module")
        # setting window size
        width = 864
        height = 569
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height,
                                    (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        # ====================================
        # =========== START OF UI ============
        # P.S You should not need to change
        # These parts, so you can skip them
        # =========== START OF UI =============
        # =====================================

        listOfAllChildren = []
        listOfAllChildren_var = []
        ft = tkFont.Font(family='Times',size=10)
        
        self.coveredDate_var = tk.IntVar()
        self.coveredDate_start_var = tk.IntVar()
        self.coveredDate_end_var = tk.IntVar()
        self.parentCoveredDate_var = tk.IntVar()
        
        self.fiscalYear_var = tk.IntVar()
        self.fiscalYear_start_var = tk.IntVar()
        self.fiscalYear_end_var = tk.IntVar()
        self.parentFiscalYear_var = tk.IntVar()
        
        self.accountant_var = tk.IntVar()
        self.accountant_name_var = tk.IntVar()
        self.accountant_type_desc_var = tk.IntVar()
        self.parentAccountant_var = tk.IntVar()
        
        self.entry_var = tk.IntVar()
        self.entries_comment_var = tk.IntVar()
        self.entry_number_var = tk.IntVar()
        self.parentEntry_var = tk.IntVar()
        
        self.line_var = tk.IntVar()
        self.line_number_var = tk.IntVar()
        self.line_counter_var = tk.IntVar()
        self.parentLine_var = tk.IntVar()
        
        self.document_var = tk.IntVar()
        self.documentReference_var = tk.IntVar()
        self.documentType_var = tk.IntVar()
        self.documentTypeDescription_var = tk.IntVar()
        self.documentNumber_var = tk.IntVar()
        self.documentDate_var = tk.IntVar()
        self.parentDocument_var = tk.IntVar()
        
        self.other_var = tk.IntVar()
        self.organizationIdentifier_var = tk.IntVar()
        self.businessDescription_var = tk.IntVar()
        self.entryNumberCounter_var = tk.IntVar()
        self.uniqueID_var = tk.IntVar()
        self.postingDate_var = tk.IntVar()
        self.EnteredBy_var = tk.IntVar()
        self.creationDate_var = tk.IntVar()
        self.enteredDate_var = tk.IntVar()
        self.sourceApplication_var = tk.IntVar()
        self.paymentMethod_var = tk.IntVar()
        self.parentOther_var = tk.IntVar()
        
        # Prepare the list of all children
        listOfAllChildren.append((self.organizationIdentifier_var, 1, 'organizationIdentifier'))
        listOfAllChildren.append((self.businessDescription_var, 2, 'businessDescription'))
        listOfAllChildren.append((self.fiscalYear_start_var, 3, 'fiscalYearStart'))
        listOfAllChildren.append((self.fiscalYear_end_var, 4, 'fiscalYearEnd'))
        listOfAllChildren.append((self.accountant_name_var, 5, 'accountantName'))
        listOfAllChildren.append((self.accountant_type_desc_var, 6, 'accountantEngagementTypeDescription'))
        listOfAllChildren.append((self.uniqueID_var, 9, 'uniqueID'))
        listOfAllChildren.append((self.creationDate_var, 10, 'creationDate'))
        listOfAllChildren.append((self.entries_comment_var, 11, 'entriesComment'))
        listOfAllChildren.append((self.coveredDate_start_var, 12, 'periodCoveredStart')) # The 1 here is column number in the cleaned version of the XML)
        listOfAllChildren.append((self.coveredDate_end_var, 13, 'periodCoveredEnd'))
        listOfAllChildren.append((self.sourceApplication_var, 14, 'sourceApplication'))
        listOfAllChildren.append((self.EnteredBy_var, 15, 'enteredBy'))
        listOfAllChildren.append((self.enteredDate_var, 16, 'enteredDate'))
        listOfAllChildren.append((self.entry_number_var, 17, 'entryNumber'))
        listOfAllChildren.append((self.documentType_var, 18, 'documentType'))
        listOfAllChildren.append((self.entryNumberCounter_var, 19, 'entryNumberCounter'))
        listOfAllChildren.append((self.line_number_var, 20, 'lineNumber'))
        listOfAllChildren.append((self.line_counter_var, 21, 'lineNumberCounter'))
        listOfAllChildren.append((self.postingDate_var, 22, 'postingDate'))
        listOfAllChildren.append((self.documentReference_var, 23, 'documentReference'))
        listOfAllChildren.append((self.documentTypeDescription_var, 24, 'documentTypeDescription'))
        listOfAllChildren.append((self.documentNumber_var, 25, 'documentNumber'))
        listOfAllChildren.append((self.documentDate_var, 26, 'documentDate'))
        listOfAllChildren.append((self.paymentMethod_var, 27, 'paymentMethod'))
        
        # Covered Date
        
        coveredDateChildren = [self.coveredDate_start_var, self.coveredDate_end_var]
        self.coveredDate=tk.Menubutton(root, text="PeriodCovered",relief=RAISED)
        
        self.coveredDate.menu =  Menu( self.coveredDate, tearoff = 0 )
        self.coveredDate.place(x=30,y=70,width=100)
        self.coveredDate["menu"] =  self.coveredDate.menu
        
        self.coveredDate.menu.add_checkbutton(label="Select/Deselect Group", variable = self.parentCoveredDate_var,
        command = lambda: self.parentPressed(self.parentCoveredDate_var,coveredDateChildren,True,False))
        self.coveredDate.menu.add_separator()
        self.coveredDate.menu.add_checkbutton(label="Period Covered Start", variable = self.coveredDate_start_var,
        command = lambda: self.checkboxPressed(self.coveredDate_start_var,self.parentCoveredDate_var,coveredDateChildren,listOfAllChildren))
        self.coveredDate.menu.add_checkbutton(label="Period Covered End", variable = self.coveredDate_end_var,
        command = lambda: self.checkboxPressed(self.coveredDate_end_var,self.parentCoveredDate_var,coveredDateChildren,listOfAllChildren))
      
        # Fiscal Year
      
        fiscalYearChildren = [self.fiscalYear_start_var, self.fiscalYear_end_var]
        self.fiscalYear=tk.Menubutton(root, text="Fiscal Year",relief=RAISED)
        
        self.fiscalYear.menu =  Menu( self.fiscalYear, tearoff = 0 )
        self.fiscalYear.place(x=140,y=70,width=100)
        self.fiscalYear["menu"] =  self.fiscalYear.menu
        
        self.fiscalYear.menu.add_checkbutton(label="Select/Deselect Group", variable = self.parentFiscalYear_var,
        command = lambda: self.parentPressed(self.parentFiscalYear_var,fiscalYearChildren,True,False))
        self.fiscalYear.menu.add_separator()
        self.fiscalYear.menu.add_checkbutton(label="Fiscal Year Start", variable = self.fiscalYear_start_var,
        command = lambda: self.checkboxPressed(self.fiscalYear_start_var,self.parentFiscalYear_var,fiscalYearChildren,listOfAllChildren))
        self.fiscalYear.menu.add_checkbutton(label="Fiscal Year End", variable = self.fiscalYear_end_var,
        command = lambda: self.checkboxPressed(self.fiscalYear_end_var,self.parentFiscalYear_var,fiscalYearChildren,listOfAllChildren))
        
        # Accountant
        
        accountantChildren = [self.accountant_name_var, self.accountant_type_desc_var]
        self.accountant=tk.Menubutton(root, text="Accountant",relief=RAISED)
        
        self.accountant.menu =  Menu( self.accountant, tearoff = 0 )
        self.accountant.place(x=250,y=70,width=100)
        self.accountant["menu"] =  self.accountant.menu
        
        self.accountant.menu.add_checkbutton(label="Select/Deselect Group", variable = self.parentAccountant_var,
        command = lambda: self.parentPressed(self.parentAccountant_var,accountantChildren,True,False))
        self.accountant.menu.add_separator()
        self.accountant.menu.add_checkbutton(label="Accountant Name", variable = self.accountant_name_var,
        command = lambda: self.checkboxPressed(self.accountant_name_var,self.parentAccountant_var,accountantChildren,listOfAllChildren))
        self.accountant.menu.add_checkbutton(label="Accountant Description", variable = self.accountant_type_desc_var,
        command = lambda: self.checkboxPressed(self.accountant_type_desc_var,self.parentAccountant_var,accountantChildren,listOfAllChildren))
        
        # Entries
        
        entryChildren = [self.entries_comment_var, self.entry_number_var]
        self.entry=tk.Menubutton(root, text="Entries",relief=RAISED)
        
        self.entry.menu =  Menu( self.entry, tearoff = 0 )
        self.entry.place(x=360,y=70,width=100)
        self.entry["menu"] =  self.entry.menu
        
        self.entry.menu.add_checkbutton(label="Select/Deselect Group", variable = self.parentEntry_var,
        command = lambda: self.parentPressed(self.parentEntry_var,entryChildren,True,False))
        self.entry.menu.add_separator()
        self.entry.menu.add_checkbutton(label="Entries Comment", variable = self.entries_comment_var,
        command = lambda: self.checkboxPressed(self.entries_comment_var,self.parentEntry_var,entryChildren,listOfAllChildren))
        self.entry.menu.add_checkbutton(label="Entry Number", variable = self.entry_number_var,
        command = lambda: self.checkboxPressed(self.entry_number_var,self.parentEntry_var,entryChildren,listOfAllChildren))
        
        # Line
       
        lineChildren = [self.line_number_var, self.line_counter_var]
        self.line=tk.Menubutton(root, text="Line",relief=RAISED)
        
        self.line.menu =  Menu( self.line, tearoff = 0 )
        self.line.place(x=470,y=70,width=100)
        self.line["menu"] =  self.line.menu
        
        self.line.menu.add_checkbutton(label="Select/Deselect Group", variable = self.parentLine_var,
        command = lambda: self.parentPressed(self.parentLine_var,lineChildren,True,False))
        self.line.menu.add_separator()
        self.line.menu.add_checkbutton(label="Line Number", variable = self.line_number_var,
        command = lambda: self.checkboxPressed(self.line_number_var,self.parentLine_var,lineChildren,listOfAllChildren))
        self.line.menu.add_checkbutton(label="Line Counter", variable = self.line_counter_var,
        command = lambda: self.checkboxPressed(self.line_counter_var,self.parentLine_var,lineChildren,listOfAllChildren))
        
        # Document
        
        documentChildren = [self.documentReference_var, self.documentType_var, self.documentTypeDescription_var, self.documentNumber_var, self.documentDate_var]
        self.document=tk.Menubutton(root, text="Document",relief=RAISED)
        
        self.document.menu =  Menu( self.document, tearoff = 0 )
        self.document.place(x=580,y=70,width=100)
        self.document["menu"] =  self.document.menu
        
        self.document.menu.add_checkbutton(label="Select/Deselect Group", variable = self.parentDocument_var,
        command = lambda: self.parentPressed(self.parentDocument_var,documentChildren,True,False))
        self.document.menu.add_separator()
        self.document.menu.add_checkbutton(label="Document Reference", variable = self.documentReference_var,
        command = lambda: self.checkboxPressed(self.documentReference_var,self.parentDocument_var,documentChildren,listOfAllChildren))
        self.document.menu.add_checkbutton(label="Document Type", variable = self.documentType_var,
        command = lambda: self.checkboxPressed(self.documentType_var,self.parentDocument_var,documentChildren,listOfAllChildren))
        self.document.menu.add_checkbutton(label="Document Type Description", variable = self.documentTypeDescription_var,
        command = lambda: self.checkboxPressed(self.documentTypeDescription_var,self.parentDocument_var,documentChildren,listOfAllChildren))
        self.document.menu.add_checkbutton(label="Document Number", variable = self.documentNumber_var,
        command = lambda: self.checkboxPressed(self.documentNumber_var,self.parentDocument_var,documentChildren,listOfAllChildren))
        self.document.menu.add_checkbutton(label="Document Date", variable = self.documentDate_var,
        command = lambda: self.checkboxPressed(self.documentDate_var,self.parentDocument_var,documentChildren,listOfAllChildren))
        
        # Other
        
        otherChildren = [self.organizationIdentifier_var, self.businessDescription_var,self.entryNumberCounter_var,self.uniqueID_var,self.postingDate_var,self.EnteredBy_var,self.creationDate_var,self.enteredDate_var,self.sourceApplication_var,self.paymentMethod_var]
        self.other=tk.Menubutton(root, text="Other",relief=RAISED)
        
        self.other.menu =  Menu( self.other, tearoff = 0 )
        self.other.place(x=690,y=70,width=100)
        self.other["menu"] =  self.other.menu
        
        self.other.menu.add_checkbutton(label="Select/Deselect Group", variable = self.parentOther_var,
        command = lambda: self.parentPressed(self.parentOther_var,otherChildren,True,False))
        self.other.menu.add_separator()
        self.other.menu.add_checkbutton(label="Organization Identifier", variable = self.organizationIdentifier_var,
        command = lambda: self.checkboxPressed(self.organizationIdentifier_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Business Description", variable = self.businessDescription_var,
        command = lambda: self.checkboxPressed(self.businessDescription_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Entry Number Counter", variable = self.entryNumberCounter_var,
        command = lambda: self.checkboxPressed(self.entryNumberCounter_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Unique ID", variable = self.uniqueID_var,
        command = lambda: self.checkboxPressed(self.uniqueID_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Posting Date", variable = self.postingDate_var,
        command = lambda: self.checkboxPressed(self.postingDate_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Entered By", variable = self.EnteredBy_var,
        command = lambda: self.checkboxPressed(self.EnteredBy_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Creation Date", variable = self.creationDate_var,
        command = lambda: self.checkboxPressed(self.creationDate_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Entered Date", variable = self.enteredDate_var,
        command = lambda: self.checkboxPressed(self.enteredDate_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Source Application", variable = self.sourceApplication_var,
        command = lambda: self.checkboxPressed(self.sourceApplication_var,self.parentOther_var,otherChildren,listOfAllChildren))
        self.other.menu.add_checkbutton(label="Payment Method", variable = self.paymentMethod_var,
        command = lambda: self.checkboxPressed(self.paymentMethod_var,self.parentOther_var,otherChildren,listOfAllChildren))
                
      
        # Bottom Part of the GUI
        self.text = Text(root, state='disabled', width=40, height=5, wrap=WORD)
        self.text.place(x = 400, y = 300)
      
        self.selectAll_var = tk.IntVar()
        
        ft = tkFont.Font(family='Times',size=16)
        self.selectAll=tk.Checkbutton(root, variable=self.selectAll_var, font=ft, fg="#000000", 
        justify="center", text="Select/Deselect All", offvalue="0", onvalue="1", 
        command=lambda: self.parentPressed(self.selectAll_var, listOfAllChildren_var, True, True))
        self.selectAll.place(x=360,y=40,width=200,height=20)
      
        self.statusLabel = Label(root, text = "Status:\nInactive\n(Import a file to start)", font=("Arial", 18))
        self.statusLabel.place(x = 80, y = 300)
      
        self.selectFile_button = tk.Button(root, text = "Choose File", command=self.selectFile)
        self.selectFile_button.place(x = 480, y = 390)
        self.ok_button = tk.Button(root, text = "OK", command=lambda: self.okButtonCommand(listOfAllChildren, root))
        self.ok_button.place(x = 620, y = 390)
        # ====================================
        # =========== END OF UI ============
        # P.S Check the comments below to
        # see what variables are available to you
        # =========== END OF UI =============
        # =====================================

        ''' 
            There is the list called listOfAllChildren, it contains 21 tuples
            representing the 21 columns, each of the tuple has the format of
            (checkBox varibale, column number, name of column )
            checkBox variable: it is used to read/write the value of the check box
            column number: the order of the column after the XML is cleaned
            name of the column: The EXACT name of the field as it appears in the .XML file
        '''

        ''' 
            The other variable is listOfAllChildren_var, it contains only the 
            check box variables used to access the checkboxes, this is only
            used for the selectAll button above
        '''

        listOfAllChildren_var = [(child[0]) for child in listOfAllChildren]

        self.file = None  # this will store the file that will be imported
        self.fileCounter = 1
        self.fileList = []

    def parentPressed(self, parentVar, children, isParent, isAll):
        if parentVar.get() == 1 and isParent == True:
            for child in children:
                child.set(1)
            if isAll == True:
                self.parentCoveredDate_var.set(1)
                self.parentFiscalYear_var.set(1)
                self.parentAccountant_var.set(1)
                self.parentAccountant_var.set(1)
                self.parentEntry_var.set(1)
                self.parentLine_var.set(1)
                self.parentDocument_var.set(1)
                self.parentOther_var.set(1)
        if parentVar.get() == 0 and isParent == True:
            for child in children:
                child.set(0)
            if isAll == True:
                self.parentCoveredDate_var.set(0)
                self.parentFiscalYear_var.set(0)
                self.parentAccountant_var.set(0)
                self.parentAccountant_var.set(0)
                self.parentEntry_var.set(0)
                self.parentLine_var.set(0)
                self.parentDocument_var.set(0)
                self.parentOther_var.set(0)
        for child in children:
            if child.get() == 0:
                parentVar.set(0)
                self.selectAll_var.set(0)
       
        if parentVar == self.parentCoveredDate_var:
            self.coveredDate.menu.post(root.winfo_rootx()+30,root.winfo_rooty()+95)
        elif parentVar == self.parentFiscalYear_var:
            self.fiscalYear.menu.post(root.winfo_rootx()+140,root.winfo_rooty()+95)
        elif parentVar == self.parentAccountant_var:
            self.accountant.menu.post(root.winfo_rootx()+250,root.winfo_rooty()+95)
        elif parentVar == self.parentEntry_var:
            self.entry.menu.post(root.winfo_rootx()+360,root.winfo_rooty()+95)
        elif parentVar == self.parentLine_var:
            self.line.menu.post(root.winfo_rootx()+470,root.winfo_rooty()+95)
        elif parentVar == self.parentDocument_var:
            self.document.menu.post(root.winfo_rootx()+580,root.winfo_rooty()+95)
        elif parentVar == self.parentOther_var:
            self.other.menu.post(root.winfo_rootx()+690,root.winfo_rooty()+95)

    def checkboxPressed(self, varName, parentVar, childList, allChildren):
        if varName.get() == 0:
            parentVar.set(0)
            self.selectAll_var.set(0)
        summ = 0
        for child in childList:
            if child.get() == 1:
                summ += 1
        if summ == len(childList):
            parentVar.set(1)
            
        summAll = 0
        for child2 in allChildren:
            if child2[0].get() == 1:
                summAll += 1
        if summAll == len(allChildren):
            self.selectAll_var.set(1)
        
        if parentVar == self.parentCoveredDate_var:
            self.coveredDate.menu.post(root.winfo_rootx()+30,root.winfo_rooty()+95)
        elif parentVar == self.parentFiscalYear_var:
            self.fiscalYear.menu.post(root.winfo_rootx()+140,root.winfo_rooty()+95)
        elif parentVar == self.parentAccountant_var:
            self.accountant.menu.post(root.winfo_rootx()+250,root.winfo_rooty()+95)
        elif parentVar == self.parentEntry_var:
            self.entry.menu.post(root.winfo_rootx()+360,root.winfo_rooty()+95)
        elif parentVar == self.parentLine_var:
            self.line.menu.post(root.winfo_rootx()+470,root.winfo_rooty()+95)
        elif parentVar == self.parentDocument_var:
            self.document.menu.post(root.winfo_rootx()+580,root.winfo_rooty()+95)
        elif parentVar == self.parentOther_var:
            self.other.menu.post(root.winfo_rootx()+690,root.winfo_rooty()+95)

    def selectFile(self):
        filetypes = (
            ('XML files', '*.xml'),
        )
        filename = fd.askopenfilename(
            title='Open a file',
            initialdir='D:\Bilkent Uni\MED_IDEA Internship\e-defter Module\Code\Sample Data',
            filetypes=filetypes
        )
        self.file = filename
        self.fileList.append(filename)
        self.text.configure(state='normal')
        if filename != "":
            self.text.insert('end', str(self.fileCounter) +
                             "- " + PurePosixPath(filename).name + "\n")
            self.fileCounter += 1
        self.text.configure(state='disabled')

    def okButtonCommand(self, childrenList, root):
        # Progress bar
        self.pb = Text(root, state='disabled', width=40, height=1)
        self.pb.place(x=270, y=450)
        
        self.square1 = Text(root, state='disabled', width=1, height=1, bg="#0000EE")
        self.square1.place(x=270, y=450)
        
        self.square2 = Text(root, state='disabled', width=1, height=1, bg="#0000EE")
        self.square2.place(x=285, y=450)
        
        self.square3 = Text(root, state='disabled', width=1, height=1, bg="#0000EE")
        self.square3.place(x=300, y=450)
        
        # Disabling buttons
        self.coveredDate.config(state = 'disabled')
        self.fiscalYear.config(state = 'disabled')
        self.accountant.config(state = 'disabled')
        self.entry.config(state = 'disabled')
        self.line.config(state = 'disabled')
        self.document.config(state = 'disabled')
        self.other.config(state = 'disabled')
        self.selectAll.config(state='disabled')
        self.selectFile_button.config(state='disabled')
        self.ok_button.config(state='disabled')
 
        listOfChosenColumns = [child[2] for child in childrenList if child[0].get() == 1]
        listOfNotChosenColumns = [child[2] for child in childrenList if child[0].get() == 0]
        
        a = 15
        for filePath in self.fileList:    
            self.statusLabel.configure(text = 'Status:\nCleaning XML File\nThis may take\nup to a minute')
            root.update()

            cleanedList = clean_xbrl(filePath)          
            filteredData = pd.DataFrame(cleanedList)

            self.statusLabel.configure(text = 'Status:\nRemoving Unselected Columns')
            root.update()

            filteredData.drop(listOfNotChosenColumns, axis = 1, inplace = True)
            # conversion of the DataFrame to a cleaned/filtered XML
            fileName = Path(filePath).stem
            fileName_clean = fileName + 'clean.xml'

            self.statusLabel.configure(text = 'Status:\nExporting To IDEA')
            root.update()

            filteredData.to_xml(fileName_clean, index = False)
            filePath_clean = os.getcwd() + '\\' + fileName_clean
            self.importXMLToIdea(filePath = filePath_clean, fileName = fileName, root = root)
            
            
            self.square1.place(x = 270+a, y = 450)
            self.square2.place(x = 285+a, y = 450)
            self.square3.place(x = 300+a, y = 450)
            time.sleep(0.5)
            root.update()
            a += 15
            if a == 20*15:
                self.square1.place(x = 270+20*15, y = 450)
                self.square2.place(x = 285+20*15, y = 450)
                self.square3.place(x = 270, y = 450)
                time.sleep(0.5)
                root.update()
                self.square1.place(x = 270+21*15, y = 450)
                self.square2.place(x = 270, y = 450)
                self.square3.place(x = 285, y = 450)
                time.sleep(0.5)
                root.update()
                self.square1.place(x = 270, y = 450)
                self.square2.place(x = 285, y = 450)
                self.square3.place(x = 300, y = 450)
                time.sleep(0.5)
                root.update()
                a = 15
            
        root.destroy()

    # filePath is the path to the XML file that holds the cleaned/filtered data
    # fileName is the name you want to be given to the new .IMD DB
    def importXMLToIdea(self, filePath=None, fileName=None, root=None):
        try:
            #    idea = win32ComClient.Dispatch(dispatch="Idea.IdeaClient")
            task = idea.GetImportTask("ImportXML")
            task.InputFileName = filePath
            task.OutputFileName = fileName
            projectFolder = idea.WorkingDirectory
            self.deleteIfExists(projectFolder + '\\' + fileName + '.IMD')
            task.PerformTask()

            # delete the temp cleaned XML file
            os.remove(filePath)
        finally:
            self.statusLabel.configure(
                text='Status:\nAn Error Occured\nWhile Exporting')
            root.update()
            task = None
            db = None
            idea = None

    def deleteIfExists(self, path=None):
        if os.path.exists(path):
            os.remove(path)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
