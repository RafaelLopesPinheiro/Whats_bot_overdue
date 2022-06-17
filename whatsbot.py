# -*- coding: utf-8 -*-
"""
Created on Fri Jun  3 09:33:39 2022

@author: rafae
"""
import pywhatkit
import keyboard
from datetime import datetime
import pandas as pd



### PROGRAMA BOT WHATS ###
def read_format_file(path_name):
    excel = pd.read_excel(path_name)
    dados_df = pd.DataFrame(excel)
    dados_df['NÚMERO'] = dados_df['NÚMERO'].astype(str)
    dados_df['NÚMERO'] = '+'+dados_df['NÚMERO']
    return dados_df
    

def filtering_overdue(df):
    overdue_df = [df.iloc[[i]] for i,j in enumerate(df['VENCIMENTO']) if j < datetime.now()]
    
    return overdue_df 


def overdue_days(data):
    today = datetime.today()
    #delta = datetime.today()
    
    d1 = data[0]['VENCIMENTO']
    diff = today - d1
    
    return str(diff)[4:6]


## SET VENCIMENTO AS INDEX


def send_msg(overdue_df, mensage):  
    while len(overdue_df) >= 1: 
        for i,j in enumerate(overdue_df):
            for v, x in enumerate(j): 
                if x == 'NÚMERO':
                    mask = pd.DataFrame(overdue_df[i])
                    numb = mask.iloc[0]['NÚMERO']  ## get numbers
                    delta_days = datetime.today() - overdue_df[i]['VENCIMENTO']
                    text = f"Sr(a). {mask.iloc[0][0]} sua fatura está atrasada á {str(delta_days)[4:6]} dias, fale conosco para lhe enviarmos o boleto." ## create text msg

                    pywhatkit.sendwhatmsg(numb , text,
                                          datetime.now().hour, datetime.now().minute+1)
    
                    del overdue_df[i]
                    keyboard.press_and_release('ctrl + w')
                    #keyboard.press_and_release('ctrl + w') ## NEED TWO CLOSE COMMANDS FOR CONFIRMATION TO CLOSE TAB IN BROWSER
                    
            break







def main():
    path_name = 'C:/Users/rafae/Desktop/teste_dudu.xlsx'
    #mensage = "sua fatura está atrasada"
    dados_df = read_format_file(path_name)
    overdue_df = filtering_overdue(dados_df)
    
    d0 = datetime.today()
    #d1 = overdue_df[2]['VENCIMENTO']
    #diff = d0 - d1
    mensage = overdue_days(overdue_df)
    #send_msg(overdue_df, mensage)
    overdue_df.index = overdue_df['VENCIMENTO']
    print(overdue_df)

if (__name__ == "__main__"):
    main()

