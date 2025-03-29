# Import liabrairies

import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
import xlsxwriter
from io import StringIO

# Define the Company

company_name='PLTR'

# Define the urls

urls={
   'Income Statement Annualy':f'https://stockanalysis.com/stocks/{company_name}/financials/' ,
   'Balanace Sheet Annualy': f'https://stockanalysis.com/stocks/{company_name}/financials/balance-sheet/',
   'Cash Flow Annualy': f'https://stockanalysis.com/stocks/{company_name}/financials/cash-flow-statement/',
   'Ratio Annualy': f'https://stockanalysis.com/stocks/{company_name}/financials/ratios/',
   'Income Statement Q':f'https://stockanalysis.com/stocks/{company_name}/financials/?p=quarterly',
   'Balance Sheet Q': f'https://stockanalysis.com/stocks/{company_name}/financials/balance-sheet/?p=quarterly',
   'Cash Flow Q': f'https://stockanalysis.com/stocks/{company_name}/financials/cash-flow-statement/?p=quarterly',
   'Ratio Q':f'https://stockanalysis.com/stocks/{company_name}/financials/ratios/?p=quarterly'

}


# Define the Headers

Header={
    'Agent_User':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# Create an Excel File
with pd.ExcelWriter(f'Financial Statement({company_name}).xlsx',engine='xlsxwriter') as xlwriter:
    for key,url in urls.items():
        try:
            response=requests.get(url,headers=Header)
            response.raise_for_status()

            soup= BeautifulSoup(response.content,'html.parser') #replace html.parser with lxml

            df = pd.read_html(StringIO(str(soup)), attrs={'data-test': 'financials'})[0]

            df.to_excel(xlwriter,sheet_name=key,index=True,engine='xlsxwriter')
        except requests.exceptions.RequestException as e:
            print(f'Failed to fetch {key}:{e}')
        
        except Exception as e :
            print(f'An Error occured while processing {key}:{e}')
        
print(f"Financial Statement saved to 'Financial Statement({company_name}).xlsx'")