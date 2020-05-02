#importing modules
import pandas as pd
from openpyxl import load_workbook
pd.options.mode.chained_assignment = None

wb = load_workbook('/.../analytics-22158-credit.xlsx')
creditSheet = wb.active

#cleaing credit card data
unwantedM = " 00:00:00"
unwantedY = "2020-"
dash = '-'
#01/12/2011
#DD/MM/YYYY
#cleaning columns A, C
for i in range(273):
    if type(creditSheet.cell(row= i+1 , column = 3).value) == str:
        creditSheet.cell(row= i+1 , column = 3).value = float(creditSheet.cell(row= i+1 , column = 3).value.strip(' CR'))

    if unwantedM and unwantedY and dash in str(creditSheet.cell(row= i+1 , column = 1).value):
        creditSheet.cell(row= i+1 , column = 1).value = str(creditSheet.cell(row= i+1 , column = 1).value).replace(unwantedY,'')
        creditSheet.cell(row= i+1 , column = 1).value = str(creditSheet.cell(row= i+1 , column = 1).value).replace(unwantedM,'')
        creditSheet.cell(row= i+1 , column = 1).value = str(creditSheet.cell(row= i+1 , column = 1).value).replace(dash,'/')

#adding year 2019
for j in range(224):
    creditSheet.cell(row= j+1 , column = 1).value = str(creditSheet.cell(row= j+1 , column = 1).value) + "/2019"

#adding year 2020
for k in range(272):
    if "/2019" not in str(creditSheet.cell(row= k+1 , column = 1).value):
        creditSheet.cell(row= k+1 , column = 1).value = str(creditSheet.cell(row= k+1 , column = 1).value) + "/2020"

wb.save('/.../analytics-22158-credit.xlsx')

df = pd.read_excel('/.../analytics-22158-credit.xlsx')

def find_clean_tran(tran):
    all_clean_transaction = {
        "Masked Bank Transaction" : "TOTAL FUEL Gas Station",
        "Masked Bank Transaction" : "EL EZABY",
        "Masked Bank Transaction" : "H&M-CAIRO FESTIVAL",
        "Masked Bank Transaction" : "AMERICAN EAGLE OUTFIT",
        "Masked Bank Transaction" : "CARREFOUR",
        "Masked Bank Transaction" : "MASTER EXPRESS",
        "Masked Bank Transaction" : "Spotify",
        "Masked Bank Transaction" : "EL EZABY",
        "Masked Bank Transaction" : "IKEA",
        "Masked Bank Transaction" : "HOLMES BURGERS",
        "Masked Bank Transaction" : "AMZN"
    }

    clean_transaction = all_clean_transaction[tran]

    return clean_transaction


for i in range(len(df['Transaction'])):
    df['Transaction'].loc[i] = find_clean_tran(df['Transaction'].loc[i])


df.to_csv(r'/.../analytics-22158-credit.csv', index = False)

#cleaing debit card data
df = pd.read_csv('/.../analytics-22158-debit.csv')
def get_date(date):
    date_by_day = date[0:2]
    search_key = date[2:5]
    
    months = {
        # 27/06/2019
        # 02Jun19
        "Jun" : "/06/2019",
        "Jul" : "/07/2019",
        "Aug" : "/08/2019",
        "Sep" : "/09/2019",
        "Oct" : "/10/2019",
        "Nov" : "/11/2019",
        "Dec" : "/12/2019",
        "Jan" : "/01/2020",
        "Feb" : "/02/2020",
        "Mar" : "/03/2020"
    }
    for key, val in months.items():
            if search_key in key:
                clean_date = date_by_day + val
                return clean_date
    # res = [val for key, val in months.items() search_key for search_key in search_keys if search_key in key]
    # print(res)

def find_clean_tran(tran):
    transaction_lst = [
        "PAPA JOHNS",
        "Uber",
        "Spotify",
        "WITHDRAWAL",
        "myfawry",
        "PIZZAHUT",
        "Swvl",
        "Other",
        "EMARATMISR",
        "Go Bus",
        "STARBUCKS",
        "HEART ATTACK",
        "DUKES",
        "IKEA",
        "CARREFOUR",
        "KFC",
        "RADIOSHACK",
        "CILANTRO",
        "EGYPTRAILWAYS",
        "MCDONALD"
    ]
    for i in range(len(transaction_lst)):
        if transaction_lst[i] in tran:
            return transaction_lst[i]

idx = df[df['Transaction'].str.contains('SALARY|Internet Transfer|OPENING|CLOSING|DEPOSIT', na=False)]
idx_list = list(idx.index)
# df = df.drop(df.index[idx_list])
df = df.dropna()

for i in range(len(df)):
    #df['Transaction'].loc[i] = find_clean_tran(df['Transaction'].loc[i])
    print(df['Date'].loc[i])
    #df['Date'].loc[i] = get_date(df['Date'].loc[i])

df.to_csv(r'/.../analytics-22158-debit.csv', index = False)