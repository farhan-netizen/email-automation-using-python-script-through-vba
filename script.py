df = pd.read_excel('Desktop/sales.xlsx', sheet_name='Sheet3')
html_body = df.to_html()

import pandas as pd
import os
df = pd.read_excel('Desktop/sales.xlsx', sheet_name='Sheet3')
html_body = df.to_html()

def emailer(html_body, subject, recipent):
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail= outlook.createItem(0)
    mail.To = recipent
    mail.Subject = subject
    mail.HtmlBody = "Hi," + "\n"  +  html_body + "\n" + "\n" + "\nRegards"
    mail.Display(True)
    
if __name__ == '__main__':
    emailer(html_body, 'Testing automation', 'someone@anymail.com')
