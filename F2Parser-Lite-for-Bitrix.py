import json
import csv
import pandas as pd
import requests
import os
import re
import datetime
import cssutils
from bs4 import BeautifulSoup
from transliterate import translit
from datetime import date
from json2html import *

sheet_names = pd.ExcelFile('Конкуренты.xlsx').sheet_names
rivals = pd.read_excel('Конкуренты.xlsx', sheet_name=None)
print(rivals)
print("")

data = []
n = 1
k = 0

site = {
    'https://domotehnika.by/': {
        'pagen': 'PAGEN_3',
        'page_in': 'yes',
        'itemLink': '.product-card__name',
        'itemClass': ['div','catalog-item'],        
        'itemNameClass': ['h1','.product-page-card__title'],
        'itemDescript': ['div','.product-property-item'],
        'itemPriceClass': ['div','.product-page-card__price'],
        'itemImages': ['span','.lazy-bg.product-page-card-slider-item__img.js-fly-img'],
        'itemImgUrl': 'style',
        'lastPageNumClass': '.w100p'
    },
    'http://bft.by/': {
        'pagen': 'PAGEN_1',
        'page_in': 'yes',
        'itemLink': '.bxr-element-name > a',
        'itemClass': ['div','t_2'],    
        'itemNameClass': ['h1','h1'],
        'itemDescript': ['td','.bxr-props-table'],
        'itemPriceClass': ['span','.bxr-market-current-price'],
        'itemImages': ['a','.fancybox:first-child'],
        'itemImgUrl': 'href',
        'lastPageNumClass': '.navigation-pages > a:nth-last-child(-n+2)'
    }
}

'''
ADD SITE EXAMPLE
'https://www._SITE_NAME.by/': {
    'pagen': '_PAGEN_VAL',
    'page_in': 'yes',
    'itemLink': '.product-card__name',
    'itemClass': ['_ELEMENT_DIV','_ELEMENT_CLASS'],    
    'itemNameClass': ['_ELEMENT_DIV','_ELEMENT_CLASS'],
    'itemDescript': '_ELEMENT_CLASS',
    'itemPriceClass': ['_ELEMENT_DIV','_ELEMENT_CLASS'],
    'itemImages': ['_ELEMENT_DIV','_ELEMENT_CLASS'],
    'lastPageNumClass': '_ELEMENT_CLASS'
}
'''

for s in range(0, len(sheet_names)):
    #print(rivals[sheet_names[s]])
    for r in range(0, rivals[sheet_names[s]].shape[0]):
        main_url = rivals[sheet_names[s]]['Сайт'][r]
        #print("Сайт: " + main_url)
        category = rivals[sheet_names[s]]['Категория'][r]
        category2 = rivals[sheet_names[s]]['Подкатегория 1'][r]
        category3 = rivals[sheet_names[s]]['Подкатегория 2'][r]
        category4 = rivals[sheet_names[s]]['Подкатегория 3'][r]
        category5 = rivals[sheet_names[s]]['Подкатегория 4'][r]
        print("Категория: " + category)
        url = rivals[sheet_names[s]]['Ссылка'][r]
        print(url)
        if (pd.isna(url)):
            continue       
        
        params = {site[main_url]['pagen']: 1}    
                
        last_page_num = 999             
        
        while params[site[main_url]['pagen']] <= last_page_num:
            
            if (site[main_url]['pagen'] == ''):                
                response = requests.get(url + '/' + str(params[site[main_url]['pagen']]))
            else:
                response = requests.get(url, params=params)
            print(response.url)
            print(response)
            soup = BeautifulSoup(response.text, 'html.parser')        
            
            try:
                last_page_num = soup.select_one(site[main_url]['lastPageNumClass']).text.strip()                
                try: 
                    last_page_num = int(last_page_num)
                except:
                    last_page_num = 999
            except:
                last_page_num = 1
                
            if site[main_url]['page_in'] == 'yes':
                urls = []
                for tag in soup.select(site[main_url]['itemLink']):
                    href = tag.attrs['href']
                    tag_url = main_url[:-1]+format(href)
                    #print(url)
                    urls.append(tag_url)
            else:
                items = soup.find_all(site[main_url]['itemClass'][0], class_=site[main_url]['itemClass'][1])
            #print(items)
            
            if (site[main_url]['page_in'] == 'yes'):            
                for page_url in urls:
                    #print(page_url)
                    page_response = requests.get(page_url)
                    #print(page_response)
                    page_soup = BeautifulSoup(page_response.text, 'html.parser')
                    try:
                        itemName = page_soup.select_one(site[main_url]['itemNameClass'][1]).text.strip()
                    except:
                        print('ERROR itemName: url=' + page_url)
                        continue
                    itemCode = translit(itemName, 'ru', reversed=True)
                    itemCode = itemCode.replace(' ', '-')                   
                    itemDescript = {}
                    try:                        
                        #itemDescript = page_soup.select_one(site[main_url]['itemDescript'][1])
                        for row in page_soup.select(site[main_url]['itemDescript'][1]):
                            cols = row.select(site[main_url]['itemDescript'][0])
                            cols = [c.text.strip() for c in cols]
                            itemDescript[cols[0]] = cols[1]
                            #print(itemDescript)  
                    except:
                        print('ERROR itemDescript: url=' + page_url)
                        continue                       
                    
                    try:
                        itemPrice = page_soup.select_one(site[main_url]['itemPriceClass'][1]).text.strip()
                        itemPrice = itemPrice.replace(",", ".")
                        #itemPrice = re.findall(r'\d*?[\s ]?\d*[,.][^ \n?]\d*', itemPrice)[0]
                    except:
                        print('ERROR itemPrice: url=' + page_url)
                        continue     
                    try:
                        if (site[main_url]['itemImgUrl'] == 'href'):
                            itemImage = page_soup.select_one(site[main_url]['itemImages'][1]).get('href')  
                        elif (site[main_url]['itemImgUrl'] == 'style'):
                            itemImage = page_soup.select_one(site[main_url]['itemImages'][1])['style']                                                                                                                
                            itemImage = re.findall(r'\(.*\),', str(itemImage))[0]
                            itemImage = itemImage.replace('(', '') 
                            itemImage = itemImage.replace(')', '') 
                            itemImage = itemImage.replace(',', '')                                              
                    except:
                        print('ERROR itemImage: url=' + page_url)
                        continue       
                    if itemImage is None:
                        print('ERROR itemImage: url=' + page_url)
                        continue
                    print(f'{n + k}: {itemName} за {itemPrice}')
                    item = {
                        'Название': itemName,
                        'Символьный код': itemCode,
                        'Цена': itemPrice,
                        'Описание': json2html.convert(json = itemDescript),
                        'Формат описания': 'html',
                        'Раздел (уровень 1)': 'Мебельный Компас',                      
                        'Раздел (уровень 2)': category,
                        'Раздел (уровень 3)': category2,
                        'Раздел (уровень 4)': category3,
                        'Раздел (уровень 5)': category4,
                        'Раздел (уровень 6)': category5,
                        'Картинка для анонса': main_url + itemImage,
                        'Детальная картинка': main_url + itemImage
                    }  
                    '''                    
                    img_data = []
                    itemImages = page_soup.select(site[main_url]['itemImages'][1])                                          
                    for img in itemImages:
                        itemImage = re.findall(r'\(.*\),', str(img))[0]
                        itemImage = itemImage.replace('(', '') 
                        itemImage = itemImage.replace(')', '') 
                        itemImage = itemImage.replace(',', '')
                        itemImage = main_url[:-1] + itemImage
                        img_data.append(itemImage)    
                    item['Картинки товара'] = img_data
                    print(item)
                    '''
                    if item in data:
                        last_page_num = params[site[main_url]['pagen']]
                    #print(last_page_num)
                    data.append(item)
                    k += 1
                    #break
            else:
                for n, i in enumerate(items, start=n):
                    try:
                        itemName = i.find(site[main_url]['itemNameClass'][0], class_=site[main_url]['itemNameClass'][1]).text.strip()
                    except:
                        #print('ERROR itemName: url=' + page_url)
                        continue
                    itemCode = translit(itemName, 'ru', reversed=True)
                    itemCode = itemCode.replace(' ', '-')
                    #print(itemName)
                    try:
                        itemPrice = i.find(site[main_url]['itemPriceClass'][0], class_=site[main_url]['itemPriceClass'][1]).text.strip()
                    except:
                        #print('ERROR itemPrice: url=' + url)
                        continue
                    #itemPrice = re.findall(r'\d*? ?\d*[,.][^ \n?]\d*', itemPrice)[0]
                    itemPrice = itemPrice.replace(" ", "")
                    for row in page_soup.select('.product-property-item'):
                        cols = row.select('div')
                        cols = [c.text.strip() for c in cols]
                        itemDescript[cols[0]] = cols[1]
                        #print(itemDescript)                       
                    try:
                        itemPrice = page_soup.select_one(site[main_url]['itemPriceClass'][1]).text.strip()
                    except:
                        print('ERROR itemPrice: url=' + page_url)
                        continue
                    #itemPrice = re.findall(r'\d*?[\s ]?\d*[,.][^ \n?]\d*', itemPrice)[0]
                    itemPrice = itemPrice.replace(",", ".")
                    try:                       
                        itemImage = page_soup.select_one(site[main_url]['itemImages'][1])
                        #itemImage = re.findall(r'\(.*\),', str(itemImage))[0]
                        #itemImage = itemImage.replace('(', '') 
                        #itemImage = itemImage.replace(')', '') 
                        #itemImage = itemImage.replace(',', '')
                    except:
                        print('ERROR itemImage: url=' + page_url)
                        continue
                    print(f'{n}: {itemName} за {itemPrice}')
                    item = {
                        'Название': itemName,
                        'Символьный код': itemCode,
                        'Цена': itemPrice,
                        'Описание': json2html.convert(json = itemDescript),
                        'Формат описания': 'html',
                        'Раздел (уровень 1)': 'Мебельный Компас',                      
                        'Раздел (уровень 2)': category,
                        'Раздел (уровень 3)': category2,
                        'Раздел (уровень 4)': category3,
                        'Раздел (уровень 5)': category4,
                        'Раздел (уровень 6)': category5,
                        'Картинка для анонса': main_url + itemImage,
                        'Детальная картинка': main_url + itemImage
                    }  
                    data.append(item)
              
            print('End page.')
            params[site[main_url]['pagen']] += 1
        #break
        
            
    print("Парсинг завершён!")

    dt_now = str(date.today())

    with open('Товары конкурентов от '+dt_now+'.json', 'w', encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)

    df = pd.read_json('Товары конкурентов от '+dt_now+'.json')
    df.to_csv('Товары конкурентов от '+dt_now+'.csv', index=False)
    df.to_excel('Товары конкурентов от '+dt_now+'.xlsx', index=False)
    print("Данные сохраненны!")
    os.remove('Товары конкурентов от '+dt_now+'.json')