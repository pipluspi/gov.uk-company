import pandas as pd
import requests as rq
from bs4 import BeautifulSoup

company_df = pd.read_excel('Company.xlsx')

company_df.columns = company_df.columns.str.strip()
company_df['CompanyNumber'] = company_df['CompanyNumber'].astype('str').str.strip()
company_df.rename(columns={'Unnamed: 29' : 'Current Status'},inplace=True)

company_director_df = pd.DataFrame()
company_status_df = pd.DataFrame()
for company_number in company_df['CompanyNumber'] :
    print(company_number)
    
    #### Company Director ####
    company_director_temp_df = pd.DataFrame()
    company_director_temp_df['CompanyNumber'] = [company_number]
    company_director_url = "https://find-and-update.company-information.service.gov.uk/company/"+company_number+"/officers"
    response = rq.get(url=company_director_url)
    if response.status_code != 200 :
        if len(company_number) != 8 :
            zeor_add = 8 - len(company_number)
            real_company_number = '0'*zeor_add + company_number
            company_director_url = "https://find-and-update.company-information.service.gov.uk/company/"+real_company_number+"/officers"
            response = rq.get(url=company_director_url)
    print(company_director_url)
    soup = BeautifulSoup(response.text,features="lxml")
    all_div = soup.find_all('div',{"class" : "appointments-list"})
    div_soup = BeautifulSoup(str(all_div),features="lxml")
    div_soup = soup.find_all('div',{"class" : "appointment-1"})
    try :
        for a in div_soup[0].find_all('a', href=True):
            company_director_temp_df['Company director'] = [a.text]
    except :
            company_director_temp_df['Company director'] = ['']
    company_director_df = pd.concat([company_director_df,company_director_temp_df],axis=0,ignore_index=True)
    #### Company Director ####
    
    #### Company Status ####
    company_status_temp_df = pd.DataFrame()
    company_status_temp_df['CompanyNumber'] = [company_number]
    company_status_url = "https://find-and-update.company-information.service.gov.uk/company/"+company_number
    response = rq.get(url=company_status_url)

    if response.status_code != 200 :
        if len(company_number) != 8 :
            zeor_add = 8 - len(company_number)
            real_company_number = '0'*zeor_add + company_number
            company_status_url = "https://find-and-update.company-information.service.gov.uk/company/"+real_company_number
            response = rq.get(url=company_status_url)
    print(company_status_url)
    soup = BeautifulSoup(response.text,features="lxml")
    all_div = soup.find_all('div',{"class" : "grid-row"})
    div_soup = BeautifulSoup(str(all_div),features="lxml")
    div_soup = soup.find_all('dl',{"class" : "column-two-thirds"})
    try :
        for text in div_soup[0].find_all('dd') :
            company_status_temp_df['Current Status'] = [text.text.strip().replace('\n','').replace('  ','')]
    except :
            company_status_temp_df['Current Status'] = ['']
    company_status_df = pd.concat([company_status_df,company_status_temp_df],axis=0,ignore_index=True)
    #### Company Status ####

company_df.update(company_director_df,join='left')
company_df.update(company_status_df,join='left')
company_df.to_excel('Company_new.xlsx',index=False)
