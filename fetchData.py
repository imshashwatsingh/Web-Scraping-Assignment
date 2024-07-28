import pandas as pd
import requests
import os
from bs4 import BeautifulSoup

wd = os.getcwd()

df = pd.read_excel(wd+"\\IOFiles\\Input.xlsx")
# df.head()I
os.mkdir(wd+"\\articles")

for i in range(df.shape[0]):
    url = df.iloc[i,1]
    page = requests.get(url)
    soup = BeautifulSoup(page.text,'html')
    content = soup.find('div',class_="td-post-content")
    if content != None:
        with open("articles/"+str(df.iloc[i,0])+".txt","w",encoding='utf-8') as f:
            for x in content:
                f.write(str(x.get_text()))
    else:
        print("File ",i+1,"Not Done because it has no content !")
        continue

