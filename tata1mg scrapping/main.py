from bs4 import BeautifulSoup
import pandas as pd
import openpyxl

def extract_data(html_file):
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, 'html.parser')

    product_boxes = soup.find_all('div', class_='style__product-box___3oEU6')

    product_links = []
    product_sizes = []
    product_names = []
    product_prices = []
    product_reviews = []

    for product_box in product_boxes:
        product_size_div = product_box.find('div', class_='style__pack-size___3jScl')
        product_size = product_size_div.text.strip() if product_size_div else ''

        product_name_div = product_box.find('div', class_='style__pro-title___3G3rr')
        product_name = product_name_div.text.strip() if product_name_div else ''

        product_price_div = product_box.find('div', class_='style__price-tag___KzOkY')
        product_price = product_price_div.text.strip() if product_price_div else ''

        product_review_span = product_box.find('span', class_='CardRatingDetail__weight-700___27w9q')
        product_review = product_review_span.text.strip() if product_review_span else ''

        product_link = product_box.find('a')['href'] if product_box.find('a') else ''

        product_links.append(product_link)
        product_sizes.append(product_size)
        product_names.append(product_name)
        product_prices.append(product_price)
        product_reviews.append(product_review)

    data = {
        'Product_link': product_links,
        'Product_size': product_sizes,
        'Product_name': product_names,
        'Product_price': product_prices,
        'Product_review': product_reviews
    }
    df = pd.DataFrame(data)

    return df


def main():
    # Specify your HTML files
    files = ['htmldata/fever.html', 'htmldata/diabetes.html', 'htmldata/Painkiller.html', 'htmldata/blood_pressure.html', 'htmldata/ayurvedic.html',
             'htmldata/cancer.html', 'htmldata/constipation.html', 'htmldata/headache.html', 'htmldata/infection.html', 'htmldata/loosemotion.html',
             'htmldata/anxiety.html','htmldata/arthritis.html','htmldata/asthma.html','htmldata/breathingproblem.html','htmldata/calcium.html',
             'htmldata/cholestrol.html','htmldata/fracture.html','htmldata/kidney.html','htmldata/liverinfection.html','htmldata/mindrelax.html',
             'htmldata/skincare.html','htmldata/vitamins.html','htmldata/pregnancy.html','htmldata/cough.html','htmldata/cold.html',
             'htmldata/menstruation.html','htmldata/babycare.html']

    # Read existing data from the output file
    try:
        existing_df = pd.read_excel('final_output.xlsx')
    except FileNotFoundError:
        existing_df = pd.DataFrame()

    # Process each HTML file and concatenate the results
    for file in files:
        df = extract_data(file)
        existing_df = pd.concat([existing_df, df], ignore_index=True)

    # Write to Excel file
    existing_df.to_excel('final_output.xlsx', index=False)
    print("Data has been successfully extracted and added to 'final_output.xlsx'.")


if __name__ == "__main__":
    main()

