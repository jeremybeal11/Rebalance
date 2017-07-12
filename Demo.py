from tkinter import *
import tkinter as tk
import pandas as pd 
from pandas import ExcelWriter
import xlsxwriter
#import numpy


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()

        # create widget for 1Q button
    def create_widgets(self):
        self.sQ = tk.Button(self, text = "Rebalance", command= self.RB)
        self.sQ.pack(side = 'top')

        # create a button for the second QTR
        self.sQ = tk.Button(self, text = "No rebalance", command= self.NRB)
        self.sQ.pack(side = 'top')

        #create button to quit
        self.quit = tk.Button(self, text="QUIT", fg="red",
                              command=root.destroy)
        self.quit.pack(side="bottom")

    def RB(self): 
         #ceated a dataframe that reads an excel file, specifying the sheet and index options 
         df = pd.read_excel('/Users/PFG/Desktop/Rebalance/reb.xlsx', 'Sheet1', index_col=None, na_values=['NA'])  #replace the path to where you save the folder

         #then i filted the Rebal colum to only show who needs to be rebalanced by 
         #looking for 'X' and 'x' incase someone forgets to captalize the x.
         Rebal = df[df['Rebal'].isin(['X','x'])]
         writer = pd.ExcelWriter('/Users/PFG/Desktop/Rebalance/RB_Clients.xlsx', engine='xlsxwriter')  #replace the path to where you save the folder

         Rebal.to_excel(writer, sheet_name='sheet1')
    
         return writer.save()
        

    def NRB(self):
         df = pd.read_excel('/Users/     /Desktop/Rebalance/reb.xlsx', 'Sheet1', index_col=None, na_values=['NA'])   #replace the path to where you save the folder

        #then i filted the Rebal colum to only show who needs to be rebalanced by 
        #looking for 'X' and 'x' incase someone forgets to captalize the x.

         NoRebal = df[df['NoRebal'].isin(['X','x'])]

         writer = pd.ExcelWriter('/Users/     /Desktop/Rebalance/NRB_Clients.xlsx', engine='xlsxwriter')    #replace the path to where you save the folder

         NoRebal.to_excel(writer, sheet_name='sheet1')
    
         return writer.save()

        

root = tk.Tk()
app = Application(master=root)
app.mainloop()


