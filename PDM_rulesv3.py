# -*- coding: utf-8 -*-
"""
Created on Fri Mar 31 09:12:36 2017

@author: ald28843

#Borrar indices de todos los excel en columna A
#Error de no abrir fichero, salir sin elegirlo
#Error cuando está el fichero de salida abierto
"""

## Importing the interface
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
## Importing s to work with data frames
import pandas as pd
import  tkinter.filedialog, tkinter.constants 
from tkinter import *
import openpyxl
import numpy as np
import xlsxwriter
import xlrd, xlwt
from PIL import ImageTk, Image
from numpy.random import normal
import plotly.plotly as py
import matplotlib.pyplot as plt
import plotly.graph_objs as go
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.figure import Figure
import datetime



from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib.utils import ImageReader
from io import StringIO
from collections import Counter

import matplotlib

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg
from matplotlib.figure import Figure

# To handle exceptions
import sys

## Defining a class dialog with it propierties
class filedialogclass(tkinter.Frame):

  def __init__(self, root):

    tkinter.Frame.__init__(self, root)
    
    content = ttk.Frame(root)
    
    #image = Image.open('C:\\Users\\ald28843\\Desktop\\ALD\\00- Datamarts\\PDM\\140 - Adhoc\\ALD.JPG')
    #photo = ImageTk.PhotoImage(image)
        
    
    #frame = ttk.Frame(content, borderwidth=5, relief="sunken", width=200, height=100)
    #namelbl = ttk.Label(content, text="Name")
    self.name = ttk.Entry(content,width=78)
    self.name2 = ttk.Entry(content,width=78)
    self.name3 = ttk.Entry(content,width=78)
    self.name4 = ttk.Entry(content,width=78)
    self.fileholdingname='C:\\Users\\ald28843\\Desktop\\ALD\\00-Datamarts\\PDM\\140-Adhoc\\Check Rules IT.xlsx'
    self.outputfile='C:\\Users\\ald28843\\Desktop\\ALD\\00-Datamarts\\PDM\\140-Adhoc\\OutputFile.xlsx'
    self.set_text3("Holding file in --> "+ self.fileholdingname)
    self.name3.configure(state='disabled')    
    self.set_text4("Output file in --> "+ self.outputfile)
    self.name4.configure(state='disabled')
    self.onevar = BooleanVar()
    self.twovar = BooleanVar()
    self.threevar = BooleanVar()
    self.fourvar = BooleanVar()
    self.frame2 = ttk.Frame(content, borderwidth=10, relief="sunken", width=500, height=350)
    
    #Here we will place the frame for the summary and graphs
    #self.frame3= tk.Frame(content, borderwidth=10, relief="sunken", width=200, height=200, bg='grey')
    self.frame3= tk.Frame(content, borderwidth=10, relief="sunken", width=500, height=350,background='black')
    self.frame4= tk.Frame(content, borderwidth=5, relief="sunken", width=300, height=300,background='black')
    self.c = Canvas(self.frame3, bg='white', width=400, height=360)
    self.c1 = Canvas(self.frame4, bg='white', width=405, height=100)

    
    
   
   
    self.w = Text (self.frame2,bg='white')
    self.w.insert(INSERT,"Please start selecting a file...")
    self.one = ttk.Checkbutton(content, text="ReadSource UCS", variable=self.onevar, onvalue=True)
    self.two = ttk.Checkbutton(content, text="ChangeHVFile", variable=self.twovar,command=self.changeHoldingFile, onvalue=True)
    self.four = ttk.Checkbutton(content, text="ChangeOutputFile", variable=self.fourvar,command=self.changeOutputFile, onvalue=True)    
    self.three = ttk.Checkbutton(content, text="Checks completed", variable=self.threevar, onvalue=True)
    Bstep1 = ttk.Button(content, text="1.Select UCS file",command=self.askopenfilenameUCS)
    Bstep2 = ttk.Button(content, text="2.Select Fleet file",command=self.askopenfilenameUCS)
    Bstep3 = ttk.Button(content, text="3.Show Rejected Values",command=self.plot_in_canvas)
    self.Bstep4 = ttk.Button(content, text="4.Show Mapped Values",command=self.showmappedvalues)
    
    content.grid(column=0, row=0)
    #frame.grid(column=2, row=0, columnspan=1, rowspan=1)
#    namelbl.grid(column=3, row=0, columnspan=2)
    self.frame4.grid(column=5, row=1, columnspan=4, rowspan=4,sticky="nsew") 
    self.c1.grid(column=5,row=1, columnspan=4, rowspan=4,sticky="nsew") 
    self.frame2.grid(column=0, row=6, columnspan=4, rowspan=5,sticky="nsew") 
    self.frame3.grid(column=5, row=6, columnspan=4, rowspan=5,sticky="nsew") 
    self.c.grid(column=5, row=6, columnspan=4, rowspan=5,sticky="nsew") 
    #photo.grid(column=5, row=1,columnspan=4,rowspan=5,sticky="nsew")
    
    self.w.grid(column=1, row=1,sticky="nsew")
    self.name.grid(column=1, row=1,columnspan=3,sticky="nsew")
    self.name2.grid(column=1,row=2,columnspan=3,sticky="nsew")
    self.name3.grid(column=1,row=4,columnspan=3,sticky="nsew")
    self.name4.grid(column=1,row=3,columnspan=3,sticky="nsew")
    self.one.grid(column=0, row=5)
    self.two.grid(column=1, row=5)
    self.three.grid(column=3, row=5,sticky="nsew")
    self.four.grid(column=2, row=5,sticky="nsew")
    
    Bstep1.grid(column=0, row=1,sticky="nsew")
    Bstep2.grid(column=0, row=2,sticky="nsew")    
    Bstep3.grid(column=0, row=3,sticky="nsew") 
    self.Bstep4.grid(column=0, row=4,sticky="nsew")
    self.img = tkinter.PhotoImage(file="C:\\Users\\ald28843\\Desktop\\ALD\\00-Datamarts\\PDM\\140-Adhoc\\ALD.GIF")
    self.c1.create_image(200,50,  image=self.img)
    
    self.file_opt = options = {}
     # define options for opening or saving a file
    #options['defaultextension'] = 'xlsx'
    options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
    options['initialdir'] = 'C:\\Users\\ald28843\\Desktop\\ALD\\00-Datamarts\\PDM\\140-Adhoc'
    options['initialfile'] = 'myfile.txt'
    options['parent'] = root
    options['title'] = 'This is a title'    
    
  def showmappedvalues(self):
      
    self.img2 = tkinter.PhotoImage(file="C:\\Users\\ald28843\\Desktop\\ALD\\00-Datamarts\\PDM\\140-Adhoc\\prueba_figura2.png")
    self.c.create_image(200,175,  image=self.img2)
    
    return
    
  def changeHoldingFile(self):
      if self.twovar.get() == 0:
            self.name3.configure(state='disabled')
      else:
            self.name3.configure(state='normal')
            self.fileholdingname = filedialog.askopenfilename()
            self.set_text3(self.fileholdingname)
    
  def changeOutputFile(self):
      if self.fourvar.get() == 0:
            self.name4.configure(state='disabled')
            self.outputfile=self.name4.get()
            print(self.outputfile)
      else:
            self.name4.configure(state='normal')
            self.set_text4(self.outputfile)
    
  def set_text(self,text):
    self.name.delete(0, 'end')
    self.name.insert(0,text)
    self.onevar.set(True)
    
    return
    
  def set_text2(self,text):
    #self.name.delete()
    self.name2.insert(0,text)
    self.twovar.set(True)
    
    return

  def set_text3(self,text):
    self.name3.delete(0,'end');
    self.name3.insert(0,text)
    #self.threevar.set(True)
    
    return
    
  def set_text4(self,text):
    self.name4.delete(0,'end');
    self.name4.insert(0,text)
    #self.threevar.set(True)
    
    return
    
  def plot_in_canvas(self):
      
    self.img2 = tkinter.PhotoImage(file="C:\\Users\\ald28843\\Desktop\\ALD\\00-Datamarts\\PDM\\140-Adhoc\\prueba_figura1.png")
    self.c.create_image(200,175,  image=self.img2)
        
     #self.c = FigureCanvasTkAgg(f, )
     #self.c.draw()    
    return
    
  def askopenfile(self):

    """Returns an opened file in read mode."""

    return filedialog.askopenfile(mode='r', **self.file_opt)

  def askopenfilenameUCS(self):
       

    """Returns an opened file in read mode.
    This time the dialog just returns a filename and the file is opened by your own code.
    """

    # get filename
    self.filename = tkinter.filedialog.askopenfilename(**self.file_opt)
    self.ColumnstoSearch=read_UCS(self,self.filename)    
    
    # open file on your own
    if self.filename:
      return open(self.filename, 'r')

## method to read UCS file, here are define headers ,it returns a dataframe 

def read_UCS(self,Source_in):
    
    
    start = time.time()   
    source=Source_in
    print("Reading the file: "+Source_in)
    df=pd.read_excel(source,header=None)
    

    print("Rows processed:"+ str(len(df.index)) )
    end = time.time()
    print("File was read in :"+str(round(end - start,2)) +" seconds")
    df = df.reset_index(drop=True)
    #print(df)
    df_copy=df.copy(deep=True)
    ColumnsToTest=['COUNTRY','CONTRACT NUMBER','COMPANY','SUBSIDIARY CODE','PLATE','MAKE','GENERIC MODEL','GENERIC MODEL + YEAR OF FACELIFT','DETAILED MODEL','GEAR BOX','CC','DIN','DOORS NUMBER','BODY GROUP','VEHICLE TYPE','FUEL TYPE','VAT','LIST PRICE OF VEHICLE EXCLUDING VAT, OPTIONAL AND DISCOUNT','LIST PRICE OF OPTIONAL EXCLUDING VAT AND DISCOUNT','CURRENCY','TOTAL DISCOUNTED PRICE','DATE OF FIRST REGISTRATION','CONTRACT START DATE','CONTRACTUAL RETURN DATE','CONTRACTUAL DURATION','CONTRACTUAL KM','START KM','REAL RETURN DATE','REAL RETURN KM','USED CAR SALE DATE','USED CAR SALE AMOUNT WITHOUT VAT','UCS REFURB COSTS INVOICED TO THE CUSTOMER WITHOUT VAT','+/- KM INVOICED OR CREDITED TO THE CUSTOMER WITHOUT VAT','AMOUNT SPEND BY ALD TO PUT THE CAR BACK INTO SELLING CONDITION WITHOUT VAT','EARLY TERMINATION FEE WITHOUT VAT','TRANSPORTATION COSTS WITHOUT VAT','NET BOOK VALUE AT RETURN DATE WITHOUT VAT','CONTRACTUAL RESIDUAL VALUE EXCLUDING VAT','LEASE BACK Y/N','SALE CHANNEL',' BUY BACK Y/N','NET GUIDE BOOK VALUE','COLOR','CO2 EMISSIONS ','MARKET SECTOR']
    # To avoid reading rows before the header, code will get the row with less n/a values    
     
    
    null_values=df.T.isnull().sum()
    df_copy['SumNulls']=null_values
      
    
    InitRow=df_copy[df_copy['SumNulls']==min(null_values)].index[0]
    
    
    #print(InitRow)
    
    #print(df.index.max) 
   
    df.columns=df.loc[InitRow,:]
    
    df=df[InitRow+1:][:]
    filasTotales=max(df.index)-InitRow
  
    filamax=max(df.index)
   
    filas=range(filasTotales)
    
    ColumsRead=df.columns.values
    ColumsRead_2=np.array(ColumsRead)[np.newaxis]
    
   
    #If rows have Unnamed value them we might read the next row
    for i in filas:
              
        if any('Unnamed' in s for s in ColumsRead_2):
           
            df.columns=df.loc[i+1,:]
            df=df[i+2:][:]
            #this is to delete nan in column names
            df.dropna(self, axis=0, how='any', thresh=None, subset=None)
            ColumsRead=df.columns.values
            ColumsRead_2=np.array(ColumsRead,dtype=np.float64)[np.newaxis]
            
            
    #print(ColumsRead)
        else:
            
            i=filamax     
            #print('entra en parte 2: '+i)
            #print(ColumsRead)
    
#    try:
        
        # if columns are correctly define
   
    #This is to avoid nan values in the headers
    
    print("Test1. Checking fields header on input file ")
    DiffColumns=np.setdiff1d(ColumnsToTest,ColumsRead, assume_unique=False)
    if len(DiffColumns) == 0:       
        
        self.set_text(self.filename)  
        self.w.delete('0.0','end');        
        self.w.insert(INSERT,'\n1. OK- Columns were correctly defined in input file'+'\n')
        UCSruleChecks(self,df)
    else:
        self.w.delete('0.0','end')        
        self.set_text("ERROR: The UCS file selected does not contain the correct fields, please check in below panel")
             #KO
        
             #OK
        CommonColumns=np.intersect1d(ColumnsToTest,ColumsRead_2, assume_unique=False)

        self.w.tag_configure('bold_italics', font=('Arial', 10, 'bold', 'italic'))
        self.w.insert(INSERT,"COLUMNS CORRECTLY DEFINED:"+'\n','bold_italics')        
        self.w.insert(INSERT,CommonColumns)
        self.w.insert(INSERT,'\n\n')
        self.w.insert(INSERT,"COLUMNS MISSING:"+'\n','bold_italics')  
        self.w.insert(INSERT,DiffColumns)
        self.w.insert(INSERT,'\n')
            
           
        
   # except: # catch *all* exceptions
   #     e = sys.exc_info()[0]
   #     print(e)
   #     self.set_text("ERROR: There was an error trying to read UCS file selected, please try again...")
        
   #    self.w.delete('0.0','end');
   #Now we start to test every rows in ColumnsToTest_df
def diff_month(d1_in, d2_in):
    #d1=datetime.datetime.strptime(d1_in, "%d/%m/%Y")    
    #d2=datetime.datetime.strptime(d2_in, "%d/%m/%Y") 
    d1=d1_in
    d2=d2_in
    
    
    mes1= (d1.year - d2.year) * 12
    mes2=0
    if d1.day >= d2.day:
            mes2= d1.month - d2.month
    else:
            mes2= d1.month - d2.month + 1
    return mes1 + mes2        
       
def readHoldingExcel(fileholdingname):
    source=fileholdingname
    df=pd.read_excel(source,sheetname='HOLDING VALUES')
    df = df.reset_index(drop=True)
    
    df_copy=df.copy(deep=True)
    return df_copy
    
    
    
def UCSruleChecks(self,UCSruleChecks):
    
#    # First it is needed to read HOLDING VALUES sheet    
     # if we are going to use default file then proceed else check for new file
     #We need to add the sheet before close the writer
     workbook = xlsxwriter.Workbook(self.outputfile)
     workbook.add_worksheet('UCS_Corrected&Mapped')
     workbook.add_worksheet('DuplicateContracts')
     workbook.add_worksheet('Wrong CompanyName')
     workbook.add_worksheet('Wrong SubsCodes')  
     workbook.add_worksheet('DuplicatePLATES') 
     workbook.add_worksheet('Wrong MAKES')
     workbook.add_worksheet('Wrong Generic Model')
     workbook.add_worksheet('Wrong GEAR BOX')
     workbook.add_worksheet('Wrong CC')
     workbook.add_worksheet('Wrong DOORS NUMBER')
     workbook.add_worksheet('Wrong BODY GROUPS')
     workbook.add_worksheet('Wrong VEHICLE TYPE')
     workbook.add_worksheet('Wrong FUEL TYPE')
     workbook.add_worksheet('Wrong LIST PRICE')
     workbook.add_worksheet('Wrong LIST PRICE OF OPTIONAL') 
     workbook.add_worksheet('TOTAL DISCOUNTED PRICE') 
     workbook.add_worksheet('Wrong CONTRACT START DATE')
     workbook.add_worksheet('Wrong DURATION')
     
     
     
     workbook.close()
        
       
        
     if(self.twovar.get()):
         self.fileholdingname= self.name3.get()  
     
     print("Test2. Checking contract number duplicated ")
    #2. Contract Number cannont be duplicated
     if len(set(UCSruleChecks['CONTRACT NUMBER']))==len(UCSruleChecks['CONTRACT NUMBER']):
        self.w.insert(INSERT,'2. OK- There are not duplicates contracts'+'\n')
     else:
        self.w.insert(INSERT,'2. KO- There are duplicates contracts --> DuplicateContracts sheet'+'\n')
        
        # if we have duplicates contracts then we should put them into a excel column (first in an output df)
        #df_aux=UCSruleChecks.groupby(['CONTRACT NUMBER','SUBSIDIARY CODE']).size()
     duplicates_rows=UCSruleChecks.duplicated(subset=['CONTRACT NUMBER','SUBSIDIARY CODE']).to_frame().rename(columns = {0: 'Duplicated'})
     dupSeries=duplicates_rows['Duplicated']
        
     IndexDuplicated=duplicates_rows.where(dupSeries).dropna().index.values
       
     contractsDuplicated=UCSruleChecks.loc[IndexDuplicated]
     contractsDuplicated_df=pd.DataFrame(contractsDuplicated)
        

       
#        xfile=openpyxl.load_workbook(self.outputfile)
#        sheet = xfile.get_sheet_by_name('DuplicateContracts')

     contractsDuplicated_df=pd.DataFrame(contractsDuplicated)
        #Writer to allow us creating more then one sheet
     writer = pd.ExcelWriter(self.outputfile, engine='xlsxwriter')           
     contractsDuplicated_df.to_excel(writer,'DuplicateContracts')
    
     contractsDuplicated_df['ERROR']='CONTRACT DUPLICATED'
     WrongLines=contractsDuplicated_df.copy(deep=True)
     
     print("Test3. Checking company ")
    #Rule 3. COMPANY FROM HOLDING VALUES
     HoldingValues_df=readHoldingExcel(self.fileholdingname)
        
     Company_test=UCSruleChecks.merge(HoldingValues_df[['COMPANY','SUBSIDIARY CODE']], left_on='COMPANY',right_on='COMPANY',how='left',indicator=True)       
        
     Company_wrong=Company_test[Company_test['_merge']=='left_only']
     Company_wrong = Company_wrong.drop(['SUBSIDIARY CODE_y','_merge'], 1)
     Company_wrong=Company_wrong.rename(columns = {'SUBSIDIARY CODE_x': 'SUBSIDIARY CODE'})
     
     Company_wrong['ERROR']='COMPANY NAME'
     WrongLines=WrongLines.append(Company_wrong)  
        
     if len(Company_wrong) > 0:     
            self.w.insert(INSERT,'3. KO- There are company names not in HOLDING VALUES --> Wrong CompanyName sheet'+'\n')        
            
            Company_wrong.to_excel(writer,'Wrong CompanyName')
            
     else:
            self.w.insert(INSERT,'3. OK- Company names are correctly populated '+'\n')    
          
     print("Test4. Checking Subsidiary ")    
     #Rule 4. SUBSIDIARY CODE FROM HOLDING VALUES 
        
     Subs_test=UCSruleChecks.merge(HoldingValues_df[['COMPANY','SUBSIDIARY CODE']], left_on='SUBSIDIARY CODE',right_on='SUBSIDIARY CODE',how='left',indicator=True)       
        
     Subs_wrong=Subs_test[Subs_test['_merge']=='left_only']
     Subs_wrong = Subs_wrong.drop(['COMPANY_y','_merge'], 1)
     Subs_wrong['ERROR']='SUBSIDIARY CODE'
     Subs_wrong=Subs_wrong.rename(columns = {'COMPANY_x': 'COMPANY'})
     WrongLines=WrongLines.append(Subs_wrong)
      
    
        
     if len(Subs_wrong) > 0:     
            self.w.insert(INSERT,'4. KO- There are subs codes not in HOLDING VALUES --> Wrong SubsCodes sheet'+'\n')        
            
            Subs_wrong.to_excel(writer,'Wrong SubsCodes')
            
     else:
            self.w.insert(INSERT,'4. OK- Subsidiary codes are correctly populated '+'\n')  
     print("Test5. Checking Plates duplicated ")    
        #Rule 5. PLATE IS NOT DUPLICATED         
     if len(set(UCSruleChecks['PLATE']))==len(UCSruleChecks['PLATE']):
            self.w.insert(INSERT,'5. OK- There are not duplicates PLATES'+'\n')
     else:
            self.w.insert(INSERT,'5. KO- There are duplicates PLATES --> DuplicatePLATES sheet'+'\n')
            
            # if we have duplicates contracts then we should put them into a excel column (first in an output df)
            #df_aux=UCSruleChecks.groupby(['CONTRACT NUMBER','SUBSIDIARY CODE']).size()
     Pduplicates_rows=UCSruleChecks.duplicated(subset=['PLATE','SUBSIDIARY CODE']).to_frame().rename(columns = {0: 'Duplicated'})
    
     PdupSeries=Pduplicates_rows['Duplicated']
    
     PIndexDuplicated=Pduplicates_rows.where(PdupSeries).dropna().index.values
    
     PLATESDuplicated=UCSruleChecks.loc[PIndexDuplicated]
     PLATESDuplicated_df=pd.DataFrame(PLATESDuplicated)  
     PLATESDuplicated_df['ERROR']='PLATES DUPLICATED'
     WrongLines=WrongLines.append(PLATESDuplicated_df)  
     
    #Writing in a new sheet
    
     PLATESDuplicated_df.to_excel(writer,'DuplicatePLATES')
     print("Test6. Checking Makes ")    
        #Rule 6.MAKE from holding values
            
     Subs_test=UCSruleChecks.merge(HoldingValues_df[['Make']], left_on='MAKE',right_on='Make',how='left',indicator=True)       
    
     Subs_wrong=Subs_test[Subs_test['_merge']=='left_only']
                
     Subs_wrong = Subs_wrong.drop(['Make','_merge'], 1)
            
            
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'6. KO- There are MAKES not in HOLDING VALUES --> Wrong SubsCodes sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong MAKES')
            
     else:
                self.w.insert(INSERT,'6. OK- MAKES are correctly populated '+'\n')          
     Subs_wrong['ERROR']='MAKE NOT FROM HOLDING VALUES'  
     WrongLines=WrongLines.append(Subs_wrong)
     print("Test7. Checking Generic Models ")
         #Rule 7.GENERIC MODEL from holding values and Generic model sheet rules
            
     Subs_test=UCSruleChecks.merge(HoldingValues_df[['Generic Model']], left_on='GENERIC MODEL',right_on='Generic Model',how='left',indicator=True)       
            
     Subs_wrong=Subs_test[Subs_test['_merge']=='left_only']
                        
     Subs_wrong = Subs_wrong.drop(['Generic Model','_merge'], 1)
     Subs_wrong['ERROR']='GENERIC MODEL'
     WrongLines=WrongLines.append(Subs_wrong)    
     
            
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'7. KO- There are Generic Model not in HOLDING VALUES --> Wrong SubsCodes sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong Generic Model')
           
     else:
                self.w.insert(INSERT,'7. OK- Generic Model are correctly populated '+'\n')          
                
            # Now we nedd to modify the Source with mapping 
     df=pd.read_excel(self.fileholdingname,sheetname='GENERIC MODEL')
     df_GM = df.reset_index(drop=True)
     UCSruleChecks=UCSruleChecks.merge(df_GM, left_on=['MAKE','GENERIC MODEL','DETAILED MODEL'],right_on=['MAKE','GENERIC MODEL','DETAILED MODEL'],how='left',indicator=True)
  
     
     UCSruleChecks['GENERIC MODEL'] = UCSruleChecks['GENERIC MODEL'].where(UCSruleChecks['NEW GENERIC MODEL'].isnull(), UCSruleChecks['NEW GENERIC MODEL'])
     if any (t==False for t in UCSruleChecks['NEW GENERIC MODEL'].isnull()):
             self.w.insert(INSERT,'   * NEW GENERIC MODEL MAPPED (MAKE,GENERIC MODEL & DETAILED MODEL CONDITION) '+'\n')
     MappedContract_aux=UCSruleChecks[UCSruleChecks['NEW GENERIC MODEL'].isnull()==False].copy(deep=True)
     MappedContract_aux['Mapped']='NEW GENERIC MODEL'  
     MappedContract=MappedContract_aux.copy(deep=True)     
     
     UCSruleChecks = UCSruleChecks.drop(['NEW GENERIC MODEL' , '_merge'], 1)
     MappedContract = MappedContract.drop(['NEW GENERIC MODEL' , '_merge'], 1)
    
     
            
    
     print("Test8. Checking GearBox ")
         #Rule 8. GEAR BOX AUTOMATIC MANUAL OR OTHER (in holding values) also there might be transformation
     
     Subs_test=UCSruleChecks.merge(HoldingValues_df[['GEAR BOX']].dropna(), left_on='GEAR BOX',right_on='GEAR BOX',how='left',indicator=True)       
            
     Subs_wrong=Subs_test[Subs_test['_merge']=='left_only']
                        
     Subs_wrong = Subs_wrong.drop(['_merge'], 1)
     Subs_wrong['ERROR']='GEAR BOX'
     
     WrongLines=WrongLines.append(Subs_wrong)        
       
         
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'8. KO- There are GEAR BOX not in HOLDING VALUES --> Wrong GEAR BOX sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong GEAR BOX')
           
     else:
                self.w.insert(INSERT,'8. OK- GEAR BOX are correctly populated '+'\n')          


    
     #print("Mapping fields... ")
            # Now we nedd to modify the Source with mapping 
     df=pd.read_excel(self.fileholdingname,sheetname='GEAR BOX')
     df_GB = df.reset_index(drop=True)
     UCSruleChecks=UCSruleChecks.merge(df_GB, left_on=['MAKE','DETAILED MODEL'],right_on=['MAKE','DETAILED MODEL'],how='left',indicator=True)
     UCSruleChecks['GEAR BOX'] = UCSruleChecks['GEAR BOX'].where(UCSruleChecks['NEW GEAR BOX'].isnull(), UCSruleChecks['NEW GEAR BOX'])
     if any (t==False for t in UCSruleChecks['NEW GEAR BOX'].isnull()):
             self.w.insert(INSERT,'   * NEW GEAR BOX MAPPED (MAKE & DETAILED MODEL CONDITION) '+'\n')
             MappedContract_aux=(UCSruleChecks[UCSruleChecks['NEW GEAR BOX'].isnull()==False]).copy(deep=True)
             MappedContract_aux['Mapped']='NEW GEAR BOX' 
             MappedContract_aux=MappedContract_aux.drop(['FUEL TYPE_y' , '_merge','NEW GEAR BOX'], 1).rename(columns = {'FUEL TYPE_x': 'FUEL TYPE'})
             
             MappedContract=MappedContract.append(MappedContract_aux)
          
              
     
     
     
     UCSruleChecks = UCSruleChecks.drop(['FUEL TYPE_y' , '_merge','NEW GEAR BOX'], 1).rename(columns = {'FUEL TYPE_x': 'FUEL TYPE'})
     
     
    
            # Again one left join is needed with Fuel type
     
     UCSruleChecks=UCSruleChecks.merge(df_GB[['FUEL TYPE','NEW GEAR BOX']].dropna(), left_on='FUEL TYPE',right_on='FUEL TYPE',how='left',indicator=True)    
    
     #if UCSruleChecks['NEW GEAR BOX'].isnull() then UCSruleChecks['GEAR BOX'] else  UCSruleChecks['NEW GEAR BOX']
     UCSruleChecks['GEAR BOX'] = UCSruleChecks['GEAR BOX'].where(UCSruleChecks['NEW GEAR BOX'].isnull(), UCSruleChecks['NEW GEAR BOX']) 
     
     
     if any (t==False for t in UCSruleChecks['NEW GEAR BOX'].isnull()):
             self.w.insert(INSERT,'   * NEW GEAR BOX MAPPED (FUEL TYPE CONDITION) '+'\n')
             MappedContract_aux=(UCSruleChecks[UCSruleChecks['NEW GEAR BOX'].isnull()==False]).copy(deep=True)
             MappedContract_aux['Mapped']='NEW GEAR BOX' 
             
            
             MappedContract=MappedContract.append(MappedContract_aux)
             

     UCSruleChecks = UCSruleChecks.drop(['NEW GEAR BOX' , '_merge'], 1)
     
     
     #Rule 9. CC, only 0 if electronic
     
     #for cc in UCSruleChecks['CC']:
     #    if cc == 0:
     
     
     print("Test9. Checking CC ")
     Subs_wrong=Subs_wrong[0:0]     
     if 'AUTOMATIC'  in (UCSruleChecks[UCSruleChecks['CC']==0]['GEAR BOX'].values) or 'MANUAL' in (UCSruleChecks[UCSruleChecks['CC']==0]['GEAR BOX'].values) :
             
             Subs_test_C=(UCSruleChecks[['CONTRACT NUMBER','COMPANY']].where(UCSruleChecks[UCSruleChecks['CC']==0]['GEAR BOX']=='AUTOMATIC').dropna())
             Subs_test_C.append((UCSruleChecks[['CONTRACT NUMBER','COMPANY']].where(UCSruleChecks['CC'].isnull()).dropna()))
             
     
             Subs_test=UCSruleChecks.merge(Subs_test_C, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left',indicator=True)         
             Subs_wrong=Subs_test[Subs_test['_merge']=='both']
                        
             Subs_wrong = Subs_wrong.drop(['_merge'], 1)
             Subs_wrong['ERROR']='CC'
     
             WrongLines=WrongLines.append(Subs_wrong)          
             
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'9. KO- There are vehicles with CC=0 that are not ELECTRONIC --> Wrong CC sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong CC')
                   
     
           
     else:
                self.w.insert(INSERT,'9. OK- CC are correctly populated '+'\n')  
     
     #--------------------------------------------------
     
           #Rule 10.DOORS NUMBER from DOOR NUMBER values sheet rules
           
     print("Test10. Checking DOORS NUMBER ")
     # First test - check null values
     Subs_wrong1=UCSruleChecks[['CONTRACT NUMBER','COMPANY']].where(UCSruleChecks['DOORS NUMBER'].isnull()).dropna().copy(deep=True)
     
     Subs_test=UCSruleChecks.merge(Subs_wrong1, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left',indicator=True)         
     Subs_wrong=Subs_test[Subs_test['_merge']=='both'].copy(deep=True)       
     Subs_wrong['ERROR']='DOORS NUMBER'
     WrongLines=WrongLines.append(Subs_wrong)    
     
            
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'10.KO- There are empty DOORS NUMBER --> Wrong DOORS NUMBER sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong DOORS NUMBER')
           
     else:
                self.w.insert(INSERT,'10.OK- DOORS NUMBER are correctly populated '+'\n')          
    
            
            # Now we nedd to modify the Source with mapping 
     
     df_DN=pd.read_excel(self.fileholdingname,sheetname='DOORS NUMBER')
     df_DN = df_DN.reset_index(drop=True).copy(deep=True)
    
     
     #print(df_DN)
     UCSruleChecks_aux=UCSruleChecks.merge(df_DN, left_on=['BODY GROUP'],right_on=['BODY GROUP'],how='left',indicator=True)
  
   
     # Contracts to change     
     
     UCSruleChecks['DIFF']= UCSruleChecks_aux['DOORS NUMBER']-UCSruleChecks_aux['NEW DOORS NUMBER']
     NoChange=UCSruleChecks_aux[['CONTRACT NUMBER','COMPANY','DOORS NUMBER']].where(UCSruleChecks['DIFF']==0).dropna()
     
    
     
    
     #Now we will merge UCSruleChecks with NoChange to identify on UCS which of them will change
    
    #UCSruleChecks=UCSruleChecks.merge(NoChange, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left',indicator=False)
     NoChange=NoChange.reset_index()
     
     ToChange=set(NoChange['CONTRACT NUMBER']).intersection(UCSruleChecks['CONTRACT NUMBER'])
     #print(ToChange)     
     
     ToChange=pd.merge(UCSruleChecks_aux[['CONTRACT NUMBER','COMPANY','DOORS NUMBER','NEW DOORS NUMBER']],NoChange,how='left',on=['CONTRACT NUMBER','COMPANY'],indicator=True)     
     
     ToChange=ToChange[ToChange['_merge']=='left_only']
     
     #In case we have more than one option then we choose the min one
     ToChange = ToChange.groupby(['CONTRACT NUMBER','COMPANY'], as_index=False)['NEW DOORS NUMBER'].min()
     ToChange=ToChange.rename(columns={'NEW DOORS NUMBER':'DOORS NUMBER'})
     DoorsMapped=NoChange[['CONTRACT NUMBER','COMPANY','DOORS NUMBER']]
     
     DoorsMapped=DoorsMapped.append(ToChange)
    
     
     UCSruleChecks2=UCSruleChecks.merge(DoorsMapped, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left')         
     
     
    
     #if UCSruleChecks['NEW GEAR BOX'].isnull() then UCSruleChecks['GEAR BOX'] else  UCSruleChecks['NEW GEAR BOX']
     UCSruleChecks['DOORS NUMBER'] = UCSruleChecks2['DOORS NUMBER_y']
     
     UCSruleChecks = UCSruleChecks.drop(['DIFF'], 1)
          
       
     if len(ToChange) > 0:          
             self.w.insert(INSERT,'   * NEW DOORS NUMBER MAPPED '+'\n')
             MappedContract_aux=UCSruleChecks.merge(ToChange, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left')
             
             MappedContract_aux['Mapped']='NEW DOORSNUMBER' 
             
             MappedContract_aux = MappedContract_aux.drop('DOORS NUMBER_y', 1)
             MappedContract=MappedContract.append(MappedContract_aux)             
            
     print("Test11. Checking BODY GROUP ") 
     #-------------------------------------------------
           #Rule 11.BODY GROUP from holding values and Generic model sheet rules
      
     Subs_test=UCSruleChecks.merge(HoldingValues_df[['BODY GROUP']], left_on='BODY GROUP',right_on='BODY GROUP',how='left',indicator=True)       
            
     Subs_wrong=Subs_test[Subs_test['_merge']=='left_only']
                        
     Subs_wrong = Subs_wrong.drop(['BODY GROUP','_merge'], 1)
     Subs_wrong['ERROR']='BODY GROUP'
     WrongLines=WrongLines.append(Subs_wrong)    
     
            
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'11.KO- There are BODY GROUPs not in HOLDING VALUES --> Wrong SubsCodes sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong BODY GROUPS')
           
     else:
                self.w.insert(INSERT,'11.OK- BODY GROUPs are correctly populated '+'\n')          
                
    # Now we nedd to modify the Source with mapping 
     df=pd.read_excel(self.fileholdingname,sheetname='BODY GROUP')
     df_BG = df.reset_index(drop=True)
     
     Aux_BG=UCSruleChecks.merge(df_BG, left_on=['MAKE','GENERIC MODEL'],right_on=['MAKE','GENERIC MODEL'],how='left',indicator=True)
  
    # Depending of how columns are populated in excel file (BG) we need to mapp them or not  
     
     #1. If rest of rows except Make and Gen Model are populated then mapped
     Case1=Aux_BG[['CONTRACT NUMBER','COMPANY','MAKE','GENERIC MODEL','DETAILED MODEL_y','REGISTRATION DATE','COUNTRY_y','MARKET SECTOR_y','BODY GROUP_y']].copy(deep=True)
     
     Case1['New BODY GROUP']= Case1[(Case1['REGISTRATION DATE'].isnull()) & (Case1['COUNTRY_y'].isnull()) & (Case1['MARKET SECTOR_y'].isnull()) & (Case1['DETAILED MODEL_y'].isnull())]['BODY GROUP_y']
     
     
     #Case1['DETAILED MODEL_y'].isnull()).where(Case1['REGISTRATION DATE'].isnull()).where(Case1['COUNTRY_y'].isnull()).where(Case1['MARKET SECTOR_y'].isnull())
     
          
     Case1_toChange=Case1[['CONTRACT NUMBER','COMPANY','New BODY GROUP']].drop_duplicates()
             # Again one left join is needed with Fuel type
    
        
     UCSruleChecks=UCSruleChecks.merge(Case1_toChange, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left',indicator=True)    
    
     #if UCSruleChecks['NEW GEAR BOX'].isnull() then UCSruleChecks['GEAR BOX'] else  UCSruleChecks['NEW GEAR BOX']
     UCSruleChecks['BODY GROUP'] = UCSruleChecks['BODY GROUP'].where(UCSruleChecks['New BODY GROUP'].isnull(), UCSruleChecks['New BODY GROUP']) 
     
     UCSruleChecks = UCSruleChecks.drop(['New BODY GROUP' , '_merge'], 1)
     
     
     # stil pending rules to be implemented
     #-------------------------------------
     
     if len(Case1_toChange) > 0:          
             self.w.insert(INSERT,'   * NEW BODY GROUP MAPPED '+'\n')
             MappedContract_aux=UCSruleChecks.merge(Case1_toChange, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left')
             MappedContract_aux=(MappedContract_aux[MappedContract_aux['New BODY GROUP'].notnull()]).copy(deep=True)
             
             MappedContract_aux['Mapped']='NEW BODY GROUP' 
             
             MappedContract_aux = MappedContract_aux.drop(['New BODY GROUP'], 1)
             MappedContract=MappedContract.append(MappedContract_aux)  
             
     
     print("Test12. Checking Vehicle Type ") 
     #-------------------------------------------------
          
     Subs_test=UCSruleChecks.merge(HoldingValues_df[['VEHICLE TYPE']], left_on='VEHICLE TYPE',right_on='VEHICLE TYPE',how='left',indicator=True)       
    
     Subs_wrong=Subs_test[Subs_test['_merge']=='left_only']
                
     Subs_wrong = Subs_wrong.drop(['VEHICLE TYPE','_merge'], 1)
            
            
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'12.KO- There are VEHICLE TYPE not in HOLDING VALUES --> Wrong SubsCodes sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong VEHICLE TYPE')
            
     else:
                self.w.insert(INSERT,'12.OK- VEHICLE TYPE are correctly populated '+'\n')          
     Subs_wrong['ERROR']='VEHICLE TYPE FROM HOLDING VALUES'  
     WrongLines=WrongLines.append(Subs_wrong)
     
     print("Test13. Checking Fuel Type ") 
     #-------------------------------------------------
          
     Subs_test=UCSruleChecks.merge(HoldingValues_df[['FUEL TYPE']], left_on='FUEL TYPE',right_on='FUEL TYPE',how='left',indicator=True)       
    
     Subs_wrong=Subs_test[Subs_test['_merge']=='left_only']
                
     Subs_wrong = Subs_wrong.drop(['FUEL TYPE','_merge'], 1)
            
            
     if len(Subs_wrong) > 0:     
                self.w.insert(INSERT,'13.KO- There are FUEL TYPE not in HOLDING VALUES --> Wrong SubsCodes sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong FUEL TYPE')
            
     else:
                self.w.insert(INSERT,'13.OK- FUEL TYPE are correctly populated '+'\n')          
     Subs_wrong['ERROR']='FUEL TYPE FROM HOLDING VALUES'  
     WrongLines=WrongLines.append(Subs_wrong)
     
         
     #Now we need to mapped the values, taking into account detail model desc
     # First we indentify the detail model inside the description then we merge with fuel type
     df=pd.read_excel(self.fileholdingname,sheetname='FUEL TYPE')
     df_FT = df.reset_index(drop=True)
     df_FT['DETAILED MODEL']=df_FT['DETAILED MODEL'].str.replace('"','')
     
     Match =UCSruleChecks[['CONTRACT NUMBER','COMPANY']].copy(deep=True) 
     Match['Mapped']=False
     Match['ValueToMapped']=None
     Match2=Match[0:0]
     Final=Match[0:0]
    
     for t in df_FT['DETAILED MODEL']:
        
        #if some of the details model contain value to mapped
        Match['Mapped']=(UCSruleChecks['DETAILED MODEL'].str.contains(t, case=True, flags=0, regex=True).copy(deep=True))                 
        Match2=(Match.where(Match['Mapped']).dropna(how='all'))
        Match2['ValueToMapped']=t
        #Match.loc[[Match['Mapped']==True],['ValueToMapped']]=t
        Final=Final.append(Match2)
      
     Final=Final.merge(df_FT, left_on=['ValueToMapped'],right_on=['DETAILED MODEL'],how='left')    
     if len(Final) > 0:          
             self.w.insert(INSERT,'   * NEW FUEL TYPE MAPPED '+'\n')
             MappedContract_aux=UCSruleChecks.merge(Final, left_on=['CONTRACT NUMBER','COMPANY'],right_on=['CONTRACT NUMBER','COMPANY'],how='left')
            
             
             UCSruleChecks['FUEL TYPE'] = MappedContract_aux['FUEL TYPE'].where(MappedContract_aux['NEW FUEL TYPE'].isnull(), MappedContract_aux['NEW FUEL TYPE']) 
             
                          
             MappedContract_aux=(MappedContract_aux[MappedContract_aux['NEW FUEL TYPE'].notnull()]).copy(deep=True)
                       
             MappedContract_aux=MappedContract_aux.reset_index()
                      
             
             MappedContract_aux1=MappedContract_aux.drop(['Mapped', 'ValueToMapped', 'DETAILED MODEL_y', 'NEW FUEL TYPE'],1).copy()
             
             MappedContract_aux1['Mapped']='NEW FUEL TYPE' 
             #MappedContract_aux=MappedContract_aux.drop(['NEW FUEL TYPE'])
            
             MappedContract=MappedContract.append(MappedContract_aux1)       
            
             
             #Now we modify the input with mapped values
                

             
             #UCSruleChecks=UCSruleChecks.merge(df_GB, left_on=['MAKE','DETAILED MODEL'],right_on=['MAKE','DETAILED MODEL'],how='left',indicator=True)
             #UCSruleChecks['GEAR BOX'] = UCSruleChecks['GEAR BOX'].where(UCSruleChecks['NEW GEAR BOX'].isnull(), UCSruleChecks['NEW GEAR BOX'])
   
             
     #print(Match[['CONTRACT NUMBER','COMPANY','Mapped2']])         
        
                    
            
     print("Test14. List price of vehicle > 0") 
     #-------------------------------------------------
     
     Subs_wrong=UCSruleChecks.where(UCSruleChecks['LIST PRICE OF VEHICLE EXCLUDING VAT, OPTIONAL AND DISCOUNT']<=0).dropna(how='all')
          
     if len(Subs_wrong)>0:
                self.w.insert(INSERT,'14.KO- There are LIST PRICE lower or equal to 0 --> Wrong LIST PRICE sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong LIST PRICE')
            
     else:
                self.w.insert(INSERT,'14.OK- LIST PRICES are correctly populated '+'\n')          
     Subs_wrong['ERROR']='LIST PRICE'  
     WrongLines=WrongLines.append(Subs_wrong)   
        
     print("Test15. List price of optional < List price of vehicle") 
     #-------------------------------------------------
     
     Subs_wrong=UCSruleChecks.where(UCSruleChecks['LIST PRICE OF VEHICLE EXCLUDING VAT, OPTIONAL AND DISCOUNT']<=UCSruleChecks['LIST PRICE OF OPTIONAL EXCLUDING VAT AND DISCOUNT']).dropna(how='all')
          
     if len(Subs_wrong)>0:
                self.w.insert(INSERT,'15.KO- There are LP OF OPT higher than the LP --> Wrong LP optional sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong LIST PRICE OF OPTIONAL')
            
     else:
                self.w.insert(INSERT,'15.OK- LIST PRICE of optional is correctly populated '+'\n')          
     Subs_wrong['ERROR']='LIST PRICE OF OPTIONAL'  
     WrongLines=WrongLines.append(Subs_wrong)   
        
     
     print("Test16. TOTAL DISCOUNTED PRICE <  LIST PRICE + OPTIONS") 
     #-------------------------------------------------
     
     Subs_wrong=UCSruleChecks.where(UCSruleChecks['TOTAL DISCOUNTED PRICE']>UCSruleChecks['LIST PRICE OF OPTIONAL EXCLUDING VAT AND DISCOUNT']+UCSruleChecks['LIST PRICE OF VEHICLE EXCLUDING VAT, OPTIONAL AND DISCOUNT']).dropna(how='all')
          
     if len(Subs_wrong)>0:
                self.w.insert(INSERT,'16.KO- TOTAL DISCOUNTED PRICE higher than the LP + OPT --> Wrong TDPrice sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong TOTAL DISCOUNTED PRICE')
            
     else:
                self.w.insert(INSERT,'16.OK- TOTAL DISCOUNTED PRICE is correctly populated '+'\n')          
     Subs_wrong['ERROR']='TOTAL DISCOUNTED PRICE'  
     WrongLines=WrongLines.append(Subs_wrong)   
        
     print("Test17. RETURN DATE > START DATE >  DATE OF FIRST REGISTRATION ") 
     #-------------------------------------------------
     
     Subs_wrong=UCSruleChecks.where((UCSruleChecks['CONTRACTUAL RETURN DATE']<UCSruleChecks['CONTRACT START DATE']) | (UCSruleChecks['CONTRACT START DATE'] < UCSruleChecks['DATE OF FIRST REGISTRATION'] ) ).dropna(how='all')
          
     if len(Subs_wrong)>0:
                self.w.insert(INSERT,'17.KO- START DATE vs REGISTRATION & RETURN DATE  --> Wrong START DATE sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong CONTRACT START DATE')
            
     else:
                self.w.insert(INSERT,'17.OK- CONTRACT START DATE is correctly populated '+'\n')          
     Subs_wrong['ERROR']='CONTRACT START DATE'  
     WrongLines=WrongLines.append(Subs_wrong)   
     
     
     
     
     print("Test18. DURATION <= (RETURN DATE - REGISTRATION DATE)") 
     #-------------------------------------------------
     UCSruleChecks['MonthsCalc']=UCSruleChecks.apply(lambda row: diff_month(row['CONTRACTUAL RETURN DATE'],row['DATE OF FIRST REGISTRATION']),axis=1)
     Subs_wrong=UCSruleChecks.where((UCSruleChecks['CONTRACTUAL DURATION'])>UCSruleChecks['MonthsCalc']).dropna(how='all')
     
     Subs_wrong.drop('MonthsCalc',axis=1)
     UCSruleChecks.drop('MonthsCalc',axis=1)
     
     if len(Subs_wrong)>0:
                self.w.insert(INSERT,'18.KO- DURATION <= (RETURN DATE - REGISTRATION DATE)  --> Wrong DURATION sheet'+'\n')        
                
                Subs_wrong.to_excel(writer,'Wrong DURATION')
            
     else:
                self.w.insert(INSERT,'18.OK- DURATION is correctly populated '+'\n')          
     Subs_wrong['ERROR']='CONTRACT DURATION'  
     WrongLines=WrongLines.append(Subs_wrong)   
           
        
            
          
     
     # Depending of how columns are populated in excel file (BG) we need to mapp them or not  
     
     #--------------------------------------------------
     
     #This last part manage output file + graphs
     #In order to print output contracts it has to be removed the ones rejected
     
     UCSoutput=UCSruleChecks.merge(WrongLines[['CONTRACT NUMBER','COUNTRY']].dropna(), left_on=['CONTRACT NUMBER','COUNTRY'],right_on=['CONTRACT NUMBER','COUNTRY'],how='left',indicator=True) 
     UCSoutput=(UCSoutput[UCSoutput['_merge']=='left_only'])
     
     UCSoutput.drop(['_merge'],1)
     UCSoutput.to_excel(writer,'UCS_Corrected&Mapped')
     print("Fields mapped ")
     
     TotalLines=len((UCSruleChecks[['COUNTRY', 'CONTRACT NUMBER']].groupby(['COUNTRY', 'CONTRACT NUMBER']).agg('count')))
     TotalLinesRejected=len((WrongLines[['COUNTRY', 'CONTRACT NUMBER']].groupby(['COUNTRY', 'CONTRACT NUMBER']).agg('count')))
     TotalLinesMapped=len(MappedContract)
     self.w.insert(INSERT,'-------------------------------------------------------------------------\nSummary:'+'\n')
     self.w.insert(INSERT,'Contracts lines: '+ str(TotalLines)+'   '+'Contracts Rejected: '+str(TotalLinesRejected)+'  '+'Total fields mapped: '+str(TotalLinesMapped)+'\n-------------------------------------------------------------------------\n')
     
    #Finally we plot the results
    # UCSruleChecks
     MappedContract_bars=(MappedContract[['CONTRACT NUMBER','Mapped']].groupby('Mapped').agg('count'))
     WrongLines_bars=(WrongLines[['CONTRACT NUMBER','ERROR']].groupby('ERROR').agg('count').reset_index().set_index('ERROR'))
     
     WrongLines_bars['CONTRACT NUMBER']=WrongLines_bars['CONTRACT NUMBER']*100/len(UCSruleChecks['CONTRACT NUMBER'])
     MappedContract_bars['CONTRACT NUMBER']=MappedContract_bars['CONTRACT NUMBER']*100/len(UCSruleChecks['CONTRACT NUMBER'])
     
     f = Figure(figsize=(200,100))
     ax = f.add_subplot()
     ay=  f.add_subplot()
     canvas = FigureCanvas(f)
     
     # Set color transparency (0: transparent; 1: solid)
     a = 0.7
     # Create a colormap
     #customcmap = [(x/24.0,  x/48.0, 0.05) for x in range(len(df))]    
      
      
           
     
     #Rejected= WrongLines_bars.plot(kind='barh',legend=None,color=customcmap,ax=ax) 
     Rejected= WrongLines_bars.plot(kind='barh',legend=None,ax=ax) 
     
     
     
     Rejected.set_xlabel("'#contracts impacted %'")  
     Rejected.set_ylabel("") 
     
     
    # Remove grid lines (dotted lines inside plot)
     Rejected.grid(False)
     # Remove plot frame
     Rejected.set_frame_on(False)
     # Position x tick labels on top
     #ax.xaxis.tick_top()
     # Pandas trick: remove weird dotted line on axis
     
     Rejected.yaxis.set_ticks_position('none')
     Rejected.xaxis.set_ticks_position('none')    
     
     f.axes.append(Rejected)  
     #
     #f.savefig('C:\\Users\\ald28843\\Desktop\\ALD\\00- Datamarts\\PDM\\140 - Adhoc\\rejected&mapped.gif',dpi=100)     
     #plt.show()
     fig1 = plt.gcf() 
     fig1.suptitle('Rejected Rows', fontsize=13)
     plt.tight_layout()
     plt.draw()
     #dpi = 68 when executing from spider
     fig1.savefig('prueba_figura1.png', dpi=50)  
     plt.close()
     
     #second= MappedContract_bars.plot(kind='barh',legend=None,title='Mapped Rows',color=customcmap,ax=ay)
     second= MappedContract_bars.plot(kind='barh',legend=None,title='Mapped Rows',ax=ay)
     second.set_ylabel("")
     second.set_xlabel("'#contracts mapped %'")  
    # Remove grid lines (dotted lines inside plot)
     second.grid(False)
     # Remove plot frame
     second.set_frame_on(False)     
     second.yaxis.set_ticks_position('none')
     second.xaxis.set_ticks_position('none')
     
     
     
     
     f.axes.append(second)  
     
     fig2 = plt.gcf()
     plt.tight_layout()    
     plt.draw()
     fig2.savefig('prueba_figura2.png', dpi=50)
    
     #ax = f.add_subplot(Rejected)
     #ay = f.add_subplot(second)
     
     print("Output file generated: "+  self.outputfile)
     self.plot_in_canvas()
     plt.close()
     
     #Rejected.savefig('temp.png2',dpi=Rejected.dpi)
     
     #plt.show()
     #dataPlot = FigureCanvasTkAgg(f, master=self.frame3)
     #dataPlot.show()
     #dataPlot.get_tk_widget().pack(expand=1)
     
    
def subplot(data, fig=None, index=111):
    if fig is None:
        fig = plt.figure()
    ax = fig.add_subplot(index)
    ax.plot(data)     
     
def main():
    
   root = tkinter.Tk()
   root.title('PDM Rules')
   filedialogclass(root)
  
   root.mainloop()  

main()









































































