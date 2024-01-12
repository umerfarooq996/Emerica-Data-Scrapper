from datetime import datetime
import re
import json
import requests
import ast
import traceback
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver

import openpyxl

import os
LOOKUP_FOLDER="lookups"
TEMPLATE=os.path.join(LOOKUP_FOLDER,"Template.xlsx")
LOOKUP_TABLE=os.path.join(LOOKUP_FOLDER,"Etnies_Lookup_Table.xlsx")

from script import get_walmart_product_data,get_shopify_product_data,get_ebay_product_data,get_amazon_product_data
from helper import remove_double_spaces,switch_words,extract_style_code,singularize

skus=[]
product_links = []
products_data = []
reviews_data = []
all_product_links=[]
raw_data=[]

def write_file(file_name, content):
    with open(f"{file_name}.html", 'w', encoding='utf-8') as file:
        file.write(str(content))

def correct_link(main_link: str, link: str):
    idx1 = main_link.find("/products/")
    idx2 = link.find("/products/")
    if idx1 != -1 and idx2 != -1:
        link = main_link[:idx1]+link[idx2:]
    return link

def extract_product_info(page_source, link):
    soup = BeautifulSoup(page_source, 'html.parser')
    write_file('temp', soup.prettify())
    js = None
    match = re.search('afterpay_product = ({.*?});', soup.prettify(),re.DOTALL)
    if match:
        product_pricing_data = match.group(1)
        js = json.loads(product_pricing_data)
    else:
        print("Error -> ", link)
        return

    def get_reviews():
        reviews = []
        dt = {}
        for div in soup.find_all("div", class_="jdgm-divider-top"):
            dt['Author'] = div.find("span", class_="jdgm-rev__author").text
            dt['Title'] = div.find("b", class_="jdgm-rev__title").text
            dt['Body'] = div.find("div", class_="jdgm-rev__body").text
            dt['Created At'] = div.find(
                "span", class_="jdgm-rev__timestamp").text
            dt['Rating'] = div.find(
                "span", class_="jdgm-rev__rating")["data-score"]
            reviews.append(dt.copy())
        return reviews

    def get_variants():
        variants = []
        for var in js['variants']:
            variants.append({'Barcode': var['barcode'], 'Sku': var['sku'], 'Color': var['option1'], 'Size': var['option2'],'Quantity':var['available']})
        return variants

    def get_color_ways():
        color_ways = []
        cont = soup.find("ul", class_="swatch-view-image")
        if not cont:
            return color_ways
        for div in cont.find_all("div", class_='swatch-group-selector'):
            url = correct_link(link, div["swatch-url"])
            if url not in product_links:
                color_ways.append(url)
        return color_ways
    price = str(js['price'])
    price=f"{price[:-2]}.{price[-2:]}"
    images = js['images']
    for i in range(len(images)):
        images[i] = images[i].replace("//", "https://")
    description = js["description"]
    return {'Title': js['title'].title(), 'Handle': js['handle'], 'Price': price, 'Description': description, 'Variants': get_variants(), 'ColorWays': get_color_ways(), 'Reviews': get_reviews(), 'Images': images}

def getPrice(p):
    if p:
        p = p.replace('$', '').strip()
        p = round(float(p))
        p = int(p)
        return p

def fix_text(text):
    # Split the text into sentences using regular expressions (excluding HTML tags)
    sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)(?<!<\/\w>)\s', text)

    # Capitalize the first letter of each sentence and lowercase the rest
    fixed_sentences = [sentence[0].capitalize() + sentence[1:].lower() for sentence in sentences]

    # Join the sentences back together to form the fixed text
    fixed_text = ' '.join(fixed_sentences)

    return fixed_text

def scrap_product(link):
    headers = {
    'authority': 'emerica.com',
    'accept': '*/*',
    'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
    'if-none-match': 'W/"cacheable:64b1e0c3c153e3a6c3b9d6e9c83cc73d"',
    'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'x-requested-with': 'XMLHttpRequest',
}
    resp=requests.get(link,headers=headers)
    return extract_product_info(resp.content,link)
    # driver.get(link)
    # driver.implicitly_wait(20)
    # return extract_product_info(driver.page_source, link)

def get_page_links():
    links=[]
    page_data=[]
    resp = requests.get('https://emerica.com')
    soup=BeautifulSoup(resp.content,'html.parser')
    for div in soup.find_all("div",class_="has-dropdown--horizontal")[:3]:
        label=div.find("label").text.strip()
        dt=(label,[])
        for a in div.find_all("a"):
            lnk='https://emerica.com{0}'.format(a["href"])
            if lnk not in links:
                dt[1].append(lnk)
                links.append(lnk)
        page_data.append(dt)
    return page_data

def get_product_type(item,link,dt):
    df = pd.read_excel(LOOKUP_TABLE)
    gender="Male"
    for type in df.values.tolist():
        if item=="FOOTWEAR":
            if type[0] in "Shoe":
                return {
                    'Link':link,
                    'Standardized Product Type': type[1],
                    'Custom Product Type': type[2],
                    'WEIGHT GRAMS': type[3],
                    'Gender': gender
                }
        else:
            if "Belt" in dt['Title']:
                if type[0] in "Accessorie":
                    return {
                        'Link':link,
                        'Standardized Product Type': type[1],
                        'Custom Product Type': type[2],
                        'WEIGHT GRAMS': type[3],
                        'Gender': gender
                    }
            else:
                if type[0] in dt['Title']:
                    return {
                        'Link':link,
                        'Standardized Product Type': type[1],
                        'Custom Product Type': type[2],
                        'WEIGHT GRAMS': type[3],
                        'Gender': gender
                    }

def get_product_links(link):
    links=[]
    try:
        count = 1
        while count != -1:
            params = {
                'page': str(count),
            }
            resp = requests.get(link, params=params)
            soup = BeautifulSoup(resp.content, 'html.parser')
            con_ls = soup.find_all("div", class_="product__imageContainer")
            for con in con_ls:
                p_link = 'https://emerica.com{0}'.format(con.find("a")["href"])
                if p_link not in all_product_links:
                    links.append(p_link)
                    # links.append([p_link,"Male"])
                    all_product_links.append(p_link)
            # break
            count += 1
            if len(con_ls) == 0:
                count = -1
    except Exception as exc:
        print(exc)
    return links

def add_prod_info(prod_data:dict):
    image_list = prod_data['Images']
    link: str = prod_data['Link']
    dt = {}
    description = prod_data['Description']
    current_variant = prod_data['Variants'][0]
    sku, color = current_variant['Sku'], current_variant['Color']
    color = color.replace("/"," ").title()
    color_code = extract_style_code(image_list, sku)
    if color_code==None:
        print(prod_data)
        return
    sku = f'{sku}-{color_code}'
    size = current_variant['Size']
    style = prod_data['Title'].title() # name of product
    item_type = prod_data['Custom Product Type'] # what is product
    gender = prod_data['Gender']
    item_type_s=item_type
    item_type_s=singularize(item_type)
    style=style.replace(item_type_s,"")
    title=""
    if "Emerica" in style:
        title = f'{style} {gender} {color} {item_type_s}'
    else:
        title = f'Emerica {gender} {style} {color} {item_type_s}'
    title = title.title().replace('Kids', 'Boys').replace(
        "Male", "Mens").replace("Female", "Womens").replace("Youth","Boys")
    if "Boys" in title and "Mens" in title:
        title=title.replace("Mens ","")
    for v in [["Shoes","Shoe"],["Socks","Sock"],["Pants","Pant"]]:
        if v[1] in title and not v[0] in title:
            title=title.replace(v[1],v[0])
            break
    title=remove_double_spaces(title)
    if 'Boys' in title:
        title=switch_words(title)
    handle=f'{title} {sku}'
    handle=handle.lower().replace(' ', '-')
    dt['Handle'] = handle
    if handle in skus:
        print("Duplicate Sku")
        return
    skus.append(handle)
    price=prod_data['Price']
    bullet_points=[]
    soup=BeautifulSoup(description,'lxml')
    description=''
    try:
        description=soup.find("p").text.strip("")
    except:
        pass
    try:
        for li in soup.find("ul").find_all("li"):
            bullet_points.append(li.text.strip())
    except:
        pass
    new_dt={
        'handle':handle,
        "new_title":title,
        "url":"",
        "title":prod_data["Title"],
        "images":image_list,
        "description":description,
        "gender":{"gender": prod_data["Gender"],
        "age_group": 'adult',
        "title_gender": prod_data["Gender"]},
        "type":item_type_s,
        "type_p": item_type,
        "color":color,
        "style_code":color_code,
            "price":price,
            "cost":getPrice(price)/2,
            "features":[],
            "bullet_points":bullet_points,
            "widths":"",
            "category":prod_data['Standardized Product Type'],
            "weight":prod_data['WEIGHT GRAMS'],
        "stock":[],
            "sizes":[]
    }
    for var in prod_data["Variants"]:
        size=var["Size"]
        qty=var['Quantity']
        if qty:
            qty="4"
        else:
            qty=0
        new_dt["stock"].append({"SKU":f"{sku}-{size}","Quantity":qty,"Upc":var["Barcode"],"size":size,'code':sku})
        new_dt["sizes"].append(size)

    products_data.append(new_dt.copy())

    for row in prod_data['Reviews']:
        rev={
            'product_handle': handle,
            'state': 'published',
            'rating': row['Rating'],
            'title': row['Title'],
            'author': row['Author'],
            'email': '',
            'body': row['Body'],
            'created_at': row['Created At']
            }
        duplicate=False
        for l in reviews_data:
            if l['author']==rev['author'] and l['body']==rev['body']:
                duplicate=True
                break
        if not duplicate:
            reviews_data.append(rev.copy())

def scrap_site():
    try:
        product_links=[]
        for item in get_page_links():
            dt=[item[0],[]]
            for link in item[1]:
                dt[1].extend(get_product_links(link))
                # dt[1].extend(get_product_links(link))
            product_links.append(dt)
        for item in product_links:
            for link in item[1]:
                dt=scrap_product(link)
                if dt:
                    type=get_product_type(item[0],link,dt)
                    if type==None:
                        print(item[0],link)
                    else:
                        dt.update(type)
                        raw_data.append(dt.copy())
                        add_prod_info(dt)
                else:
                    print("erroer ", link)
    except:
        traceback.print_exc()

def read_existing_data():
    df=pd.read_excel("output_raw.xlsx")
    df['Price'] = df['Price'].fillna("")
    df['Price'] = df['Price'].astype(str)
    df['Variants'] = df['Variants'].apply(ast.literal_eval)
    df['Reviews'] = df['Reviews'].apply(ast.literal_eval)
    df['Images'] = df['Images'].apply(ast.literal_eval)
    # df['widths'] = df['widths'].fillna("")
    for index, row in df.iterrows():
        dt = row.to_dict()
        add_prod_info(dt)

def main():
    scrap_site()
    pd.DataFrame(raw_data).to_excel("output_raw.xlsx",index=False)
    # read_existing_data()
    file_path = TEMPLATE # Replace with the path to your existing Excel file
    workbook = openpyxl.load_workbook(file_path) 
    vendor="Emerica"
    pd.DataFrame(products_data).to_excel("output.xlsx",index=False)
    get_shopify_product_data(products_data,vendor,workbook)
    get_ebay_product_data(products_data,vendor,workbook)
    get_walmart_product_data(products_data,vendor,workbook)
    get_amazon_product_data(products_data,vendor,workbook)

    current_date = datetime.now().strftime("%Y-%m-%d")
    workbook.save(f'{vendor}_{current_date}.xlsx')
    workbook.close()


    output_file = f"{current_date} -EmericaProduct.xlsx"

    df2=pd.DataFrame(reviews_data)
    df2 = df2.drop_duplicates()
    with pd.ExcelWriter(output_file) as writer:
        df2.to_excel(writer, sheet_name='Reviews',index=False)

if __name__ == '__main__':
    main()
