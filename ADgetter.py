import requests, os, bs4
caaUrl = 'https://www.caa.gov.tw/FlightInstruction.aspx?a=240&lang=1'
os.makedirs('ad', exist_ok = True)
res = requests.get(caaUrl)
res.raise_for_status()
adSoup = bs4.BeautifulSoup(res.text)

#adFile = open('ad.html', encoding="utf-8")
#adSoup = bs4.BeautifulSoup(adFile.read())
elems = adSoup.select('.download-filebase')

for i in range(0,len(elems) - 1):
    url = 'https://www.caa.gov.tw/' + elems[i].attrs['href']
    print(url)

    downLoad = requests.get(url)
    pdf = downLoad.content
    
    with open(elems[i].attrs['title'],'wb') as f:
        f.write(pdf)
        


    


    
