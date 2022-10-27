import requests
from bs4 import BeautifulSoup
import json
import xlsxwriter
import re
import os
from tqdm import tqdm

PAGES_COUNT = 6
OUT_JSON_FILENAME = 'out.json'
OUT_XLSX_FILENAME = 'out.xlsx'
main_link = 'https://cosca.ru'
headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 "
                  "Safari/537.36 "
}


def get_product_links(pages_count):
    links = []
    for i in range(pages_count):
        url = f'https://cosca.ru/internet-magazin/cosca-ecopolimer/p/{i}'

        req = requests.get(url=url, headers=headers)
        page = req.text

        # with open(f'data/index_{i}.html', 'w', encoding='utf-8') as file:
        #     file.write(page)

        # with open(f'data/index_{i}.html', 'r', encoding='utf-8') as file:
        #     src = file.read()

        soup = BeautifulSoup(page, 'lxml')
        for link in soup.find_all(class_='product-name'):
            links.append(main_link + link.find('a').get('href'))
    return links


def product_parse(links):
    data = []
    for link in links:
        req = requests.get(url=link, headers=headers)
        # with open(f'data/product.{links.index(link)}.html', 'w', encoding='utf-8') as file:
        #     file.write(req.text)
        # with open(f'data/product.{links.index(link)}.html', 'r', encoding='utf-8') as file:
        #     src = file.read()
        try:
            product = dict()
            soup = BeautifulSoup(req.text, 'lxml')
            product_id = soup.find(class_='shop2-product-article').text.split(': ')[1]
            product_name = soup.find('h1').text.split(', ')[0]
            # name = soup.find(class_='site-path').find('href')
            # price = soup.find(class_='price-current').find('strong').text
            price = re.sub(r'\s+', ' ', soup.find(class_='price-current').text.split('руб.')[0])
            product["Артикул"] = product_id
            product["Наименование"] = product_name
            product["Цена"] = price

            # get parametrs
            params = soup.find(class_='shop2-product-options')
            for i in params:
                if i.find('th').text != 'Спецификация':
                    product[i.find('th').text.split()[0]] = i.find('td').text.strip()
            data.append(product)

            # get images
            get_images(soup, product_name)

        except Exception as ex:
            data.append(product)
            print(ex)
            print('No data')
    return data


# get images
def get_images(soup, product_name):
    all_images = soup.find(class_='product-image').find_all('a')
    image_count = 1
    for image in all_images[1:]:
        images_links = image.get('href')
        image_bytes = requests.get(f'{main_link}{images_links}', headers=headers).content
        subdir = product_name
        if not os.path.exists('images/' + subdir):
            os.mkdir('images/' + subdir)
        with open(f'images/{subdir}/{product_name}-{image_count}.jpg', 'wb') as file:
            file.write(image_bytes)
        # print(f'Image - {product_name}-{image_count} successfully downloaded')
        image_count += 1


def json_dump(data):  # dump data to json
    with open(OUT_JSON_FILENAME, 'w') as file:
        json.dump(data, file, ensure_ascii=False, indent=1)


def xlsx_dump(filename, data):  # dump data to .xlsx(excel table)
    if not len(data):
        return None

    with xlsxwriter.Workbook(filename) as workbook:
        ws = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        headers = list(data[0].keys())
        for col, h in enumerate(headers):
            ws.write_string(0, col, h, cell_format=bold)

        for row, item in enumerate(data, start=1):
            for name, value in item.items():
                col = headers.index(name)
                ws.write_string(row, col, value)


def main():
    links = get_product_links(PAGES_COUNT)
    data = product_parse(links)
    json_dump(data)
    xlsx_dump(OUT_XLSX_FILENAME, data)


if __name__ == '__main__':
    main()
