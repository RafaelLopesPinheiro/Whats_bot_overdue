# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 10:31:26 2022

@author: rafael
"""

import pywhatkit
import keyboard
from datetime import datetime
import pandas as pd
from tkinter import *
import time
import pyautogui 

### PROGRAM BOT WHATS ###
def read_format_file(path_name):
    excel = pd.read_excel(path_name)
    dados_df = pd.DataFrame(excel)
    dados_df['NÚMERO'] = dados_df['NÚMERO'].astype(str)
    dados_df['NÚMERO'] = '+'+dados_df['NÚMERO'].astype(str)
    # print(dados_df['NÚMERO'])
    return dados_df
    

def filtering_overdue(df):
    overdue_df = [df.iloc[[i]] for i,j in enumerate(df['VENCIMENTO']) if j < datetime.now()]
    
    return overdue_df 


def overdue_days(data):
    today = datetime.today()
   
    d1 = data[0]['VENCIMENTO']
    diff = today - d1
    
    return str(diff)[4:6]


## SET VENCIMENTO AS INDEX


def send_msg(overdue_df):  
    while len(overdue_df) >= 1: 
        for i,j in enumerate(overdue_df):
            for v, x in enumerate(j): 
                if x == 'NÚMERO':
                    mask = pd.DataFrame(overdue_df[i])
                    name = mask.iloc[0][0]
                    numb = mask.iloc[0]['NÚMERO']  ## get numbers
                    delta_days = datetime.today() - overdue_df[i]['VENCIMENTO']
                    overdue_days = str(delta_days)[4:7].strip(' ') 

                    
                    text = f"Sr(a). {name} sua fatura está atrasada á {overdue_days} dias, fale conosco para lhe enviarmos o boleto." ## create text msg

                    pywhatkit.sendwhatmsg(numb , text,
                                          datetime.now().hour, datetime.now().minute+1)
                    
                    del overdue_df[i]
                    pyautogui.click(1050, 950)
                    time.sleep(2)
                    keyboard.press_and_release('enter')
                    keyboard.press_and_release('ctrl + w')
                    # keyboard.press_and_release('ctrl + w') ## IF NEED TWO CLOSE COMMANDS FOR CONFIRM CLOSE TAB IN BROWSER
                    
            break




def main():
    path_name = 'C:/Users/rafae/Desktop/teste_dudu.xlsx'
    dados_df = read_format_file(path_name)
    print(dados_df.head())
    overdue_df = filtering_overdue(dados_df)
    
    
    send_msg(overdue_df)
    




if (__name__ == "__main__"):
     main() 


