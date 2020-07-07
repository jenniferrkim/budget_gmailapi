from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from dateutil import parser
from money_parser import price_str
import pandas as pd
from datetime import datetime

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def query_builder():
    source = input("Choose your source: \n1. boa travel \n2. barclays \n3. citi \n4. venmo \n5. all \n")
    date = input("Enter starting date of transactions: m/d/yyyy \n")
 
    queryli = ['from:onlinebanking@ealerts.bankofamerica.com subject:Credit card transaction exceeds alert limit you set after:' + date,
              'from:alerts@service.barclaysus.com subject:Account Alert: Purchase Activity after:' + date,
              'from:alerts@citibank.com subject:transaction was made on your Citi® Double Cash Card account after:' + date,
              'from:venmo@venmo.com subject:completed OR subject:paid after:' + date]
    if source == '1' or source == 'boa travel':
        choice = 1
        query = queryli[0]
        #query = 'from:onlinebanking@ealerts.bankofamerica.com subject:Credit card transaction exceeds alert limit you set after:' + date
    elif source == '2' or source == 'barclays':
        choice = 2
        query = queryli[1]
        #query = 'from:alerts@service.barclaysus.com subject:Account Alert: Purchase Activity after:' + date
    elif source == '3' or source == 'citi':
        choice = 3
        query = queryli[2]
        #query = 'from:alerts@citibank.com subject:transaction was made on your Citi® Double Cash Card account after:' + date
    elif source == '4' or source == 'venmo':
        choice = 4
        query = queryli[3]
        #query = 'from:venmo@venmo.com subject:completed OR subject:paid after:' + date
    elif source == '5' or source == 'all':
        choice = 5
        query = queryli
        
    return choice, query

def budget_with_gmail(query):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    try:
        response = service.users().messages().list(userId='me',
                                                   q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId='me', q=query,
                                             pageToken=page_token).execute()
            messages.extend(response['messages'])
        
        #messages returns list of dictionaries {'id': 'qwerty', 'threadId': 'qwerty'}
        
        msgli = []
      
        if 'Double Cash Card' in query:
            raw_or_nah = 'raw'
        else:
            raw_or_nah = 'metadata'
        
        for item in messages:
            msg = service.users().messages().get(userId='me', id=item['id'], format = raw_or_nah).execute()
            msgli.append(msg)

        return msgli
                    
    except (errors.HttpError, error):
        print ('An error occurred: %s' % error)
        
def parse_boa(msgli):    
    dateli = []
    descli = []
    amtli = []
    
    for msg in msgli:
        # getting dates
        headerli = msg['payload']['headers']
        for nestdic in headerli:
            if nestdic['name'] == 'Date':
                date = parser.parse(nestdic['value']).date()
                dateli.append(date)
        
        # getting amounts and description of transaction
        snippet = msg['snippet']
        amount_index = snippet.find("Amount")
        date_index = snippet.find("Date")
        where_index = snippet.find("Where")
        end_index = snippet.find("View details")
        if end_index == -1:
            end_index = snippet.find("This may")

        amt_string = snippet[amount_index:date_index]
        where_string = snippet[where_index + 7:end_index]
        
        amtli.append(price_str(amt_string))
        descli.append(where_string)
            
    boa_df = pd.DataFrame(data = {'Date': dateli, 'Description': descli, 'Amount': amtli, 'Source': ['BoA Travel']*len(dateli)})
    
    return boa_df

def parse_barclays(msgli):    
    dateli = []
    descli = []
    amtli = []
    
    for msg in msgli:
        # getting dates
        headerli = msg['payload']['headers']
        for nestdic in headerli:
            if nestdic['name'] == 'Date':
                date = parser.parse(nestdic['value']).date()
                dateli.append(date)
                
        snippet = msg['snippet']
        
        # getting amounts
        amount_index = snippet.find("purchase")
        amt_string = snippet[amount_index:]
        amtli.append(price_str(amt_string))
        
    barclays_df = pd.DataFrame(data = {'Date': dateli, 'Description': ['n/a']*len(dateli), 'Amount': amtli, 'Source': ['Barclays']*len(dateli)})
    
    return barclays_df

def parse_citi(msgli):
    import base64
    dateli = []
    descli = []
    amtli = []
    
    for msg in msgli:
        # decode raw base64url encoded string to bytes
        decoded = base64.urlsafe_b64decode(msg['raw'] + '=' * (4 - len(msg['raw']) % 4))
   
        # search for relevant snippet with info
        og = True
        
        start_index = decoded.find(b"Account #: XXXX")
        end_index = decoded.find(b"exceeds the $0.00 transaction amount")
        
        if start_index == -1:
            og = False
            start_index = decoded.find(b"Citi Alert:")
            end_index = decoded.find(b"on card ending in")
            date_start_index = decoded.find(b"<jennifer.kim7@gmail.com>;")
            
        snippet = decoded[start_index:end_index]
        snippet = snippet.decode('utf-8') # convert snippet to string
      
        # amount, description, date
        if og == True:
            amount_index = snippet.find("Account #: XXXX")
            where_index = snippet.find("at")
            date_index = snippet.find("on")

            amt_string = snippet[amount_index+20:where_index]
            where_string = snippet[where_index+2:date_index]
            date_string = snippet[date_index+2: date_index + 13]
            
        if og == False:
            amount_index = snippet.find("A ")
            where_index = snippet.find("at")
            
            amt_string = snippet[amount_index:where_index]
            where_string = snippet[where_index+2:end_index]
            
            date_string = decoded[date_start_index+26: date_start_index + 43]
            date_string = date_string.decode('utf-8')
            
        amtli.append(price_str(amt_string))
        descli.append(where_string)
        date = parser.parse(date_string).date() 
        dateli.append(date)
            
    citi_df = pd.DataFrame(data = {'Date': dateli, 'Description': descli, 'Amount': amtli, 'Source': ['Citi Double Cash']*len(dateli)})

    return citi_df

def parse_venmo(msgli):
    dateli = []
    descli = []
    amtli = []
    
    for msg in msgli:
        headerli = msg['payload']['headers']
        for nestdic in headerli:
            #print(nestdic)
            # getting date of transactions
            if nestdic['name'] == 'Date':
                date = parser.parse(nestdic['value']).date()
                dateli.append(date) 
            # getting amount of transactions
            if nestdic['name'] == 'Subject':
                amtli.append(price_str(nestdic['value']))
        
        # getting descriptions
        descli.append(msg['snippet'])
    
    # change sign of amount based on whether you're paying or being paid
    for desc in descli:
        if ('You charged' in desc) or ('paid You' in desc):
            amtli[descli.index(desc)] = float(amtli[descli.index(desc)]) * -1
        else:
            amtli[descli.index(desc)] = float(amtli[descli.index(desc)])

    venmo_df = pd.DataFrame(data = {'Date': dateli, 'Description': descli, 'Amount': amtli, 'Source': ['Venmo']*len(dateli)})
    
    return venmo_df

def to_gsheets(df):
    import pygsheets
    #authorization
    gc = pygsheets.authorize(service_file='C:\\Users\\example_sheet.json')

    #open the google spreadsheet (where 'Budget v2' is the name of my sheet)
    sh = gc.open('Budget v2')

    #select the n-th sheet, counter from 0 
    wks = sh[4]

    #update the first sheet with df, starting at cell A1. 
    wks.set_dataframe(df,(1,1))

def main():
    choice, query = query_builder()

    if choice == 1:
        msgli_boa = budget_with_gmail(query)
        df = parse_boa(msgli_boa)
    elif choice == 2:
        msgli_barclays = budget_with_gmail(query)
        df = parse_barclays(msgli_barclays)
    elif choice == 3:
        msgli_citi = budget_with_gmail(query)
        df = parse_citi(msgli_citi)
    elif choice == 4:
        msgli_venmo = budget_with_gmail(query)
        df = parse_venmo(msgli_venmo)
    elif choice == 5:
        msgli_boa = budget_with_gmail(query[0])
        msgli_barclays = budget_with_gmail(query[1])
        msgli_citi = budget_with_gmail(query[2])
        msgli_venmo = budget_with_gmail(query[3])

        df_boa = parse_boa(msgli_boa)
        df_barclays = parse_barclays(msgli_barclays)
        df_citi = parse_citi(msgli_citi)
        df_venmo = parse_venmo(msgli_venmo)

        df = pd.concat([df_boa, df_barclays, df_citi, df_venmo])

    to_gsheets(df)
        
if __name__ == '__main__':
    main()