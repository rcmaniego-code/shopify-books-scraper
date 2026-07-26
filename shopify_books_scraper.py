import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import xlsxwriter

def clean_text(text):
    if not text:
        return ""

    return str(text).strip()

def create_shopify_handle(title):
    handle = title.lower()
    handle = re.sub(r'[^a-z0-9]+', '-', handle)

    return handle.strip('-')

def save_to_excel_with_styling(df, filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Shopify_Import', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Shopify_Import']
    header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#4472C4', 'border': 1, 'align': 'center'})

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    for i, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, column_len)

    worksheet.freeze_panes(1, 0)
    writer.close()

    print(f"Excel Saved: {filename}")

def process_and_save(products):
    if not products:
        return

    shopify_data = []
    for p in products:
        title = p.get('title', 'No Title')
        shopify_data.append(
            {
                'Handle': create_shopify_handle(title),
                'Title': title,
                'Body (HTML)': f"<p>{clean_text(p.get('description'))}</p>",
                'Vendor': 'Books Demo',
                'Type': 'Book',
                'Tags': 'scraped, book',
                'Published': 'TRUE',
                'Option1 Name': 'Title',
                'Option1 Value': 'Default Title',
                'Variant SKU': p.get('sku', ''),
                'Variant Price': p.get('price', '0.00'),
                'Image Src': p.get('image_url', ''),
                'Status': 'active'
            }
            )

    df = pd.DataFrame(shopify_data)
    df['Variant Price'] = df['Variant Price'].astype(str).str.replace('£','').astype(float)

    df.to_csv("shopify_books.csv", index=False, encoding='utf-8')

    save_to_excel_with_styling(df, "shopify_books.xlsx")

    print(f"DONE! {len(df)} products saved.")

def scrape_books():
    products = []
    url = "https://books.toscrape.com/catalogue/page-1.html"

    print(f"Scraping: {url}")
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers, timeout=10)

    soup = BeautifulSoup(response.content, 'html.parser')
    for item in soup.select('article.product_pod'):
        title = item.select_one('h3 a')['title']
        price = item.select_one('.price_color').text
        image = "https://books.toscrape.com/" + item.select_one('img')['src']
        products.append(
            {
                'title': title,
                'price': price,
                'description': 'Sample book description',
                'sku': f"BK-{len(products)+1:03d}",
                'image_url': image
            }
            )

    return products

if __name__ == "__main__":
    all_products = scrape_books()
    process_and_save(all_products)
