""" Script designed to return the contents of the html table of a specific email folder """

import win32com.client
import pandas as pd
from bs4 import BeautifulSoup as Bs


conn_outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
inbox = conn_outlook.GetDefaultFolder(6)  # 6 mean accesses the default "Inbox" folder in Outlook
specific_folder = inbox.Folders['Specific Folder Name']

emails = specific_folder.Items
emails = emails.Restrict("[Subject] = 'Specific subject'")

if emails.Count == 0:
    print("E-mail not found.")
else:
    email = emails.GetLast()

    html_body = email.HTMLBody
    soup = Bs(html_body, features='html_parser')
    table = soup.find('table')

    if table:
        rows = table.find_all('tr')
        table_content = []

        for row in rows:
            cols = row.find_all(['td', 'th'])
            cols = [col.get_text(strip=True) for col in cols]
            table_content.append(cols)

        if len(table_content) > 1:
            df_specific_data = pd.DataFrame(table_content[1:], columns=table_content[0])
