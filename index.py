import pandas as pd
from bs4 import BeautifulSoup
import requests


# BASE_URL = 'https://vnrvjietexams.net/eduprime3exam/Results/Results?htno={}&examId=6189&_=1711987366280'
BASE_URL = 'https://vnrvjietexams.net/eduprime3exam/Results/Results?htno={}&examId=6271&_=1727441986873'

df = pd.read_excel('roll.xlsx')
rolls = df['Roll'].tolist()

df2 = {'Roll':[], 'Name':[], 'SGPA':[]}
for roll in rolls:

    url = BASE_URL.format(roll)

    res = requests.get(url)
    soup = BeautifulSoup(res.text, 'html.parser')

    sgpa_tag = soup.find('td', string='SGPA')
    student_name = soup.find('td', string='Student Name')

    sgpa_value = sgpa_tag.find_next_sibling('td').get_text(strip=True)
    student_name = student_name.find_next_sibling('td').get_text(strip=True)

    df2['Roll'].append(roll)
    df2['Name'].append(student_name[1:].strip().capitalize())
    df2['SGPA'].append(sgpa_value[1:] if len(sgpa_value.strip()) > 0 else 'F')

df2 = pd.DataFrame(df2)
df2["rank"] = df2["SGPA"].rank(ascending=False, method='min')
df2.to_excel('result.xlsx', index=False)