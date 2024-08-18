import os
import requests
from bs4 import BeautifulSoup
import pandas as pd

input_data = pd.read_excel('./input/Input.xlsx')
output_directory = './output/extracted'

if not os.path.exists(output_directory):
    os.makedirs(output_directory)

for index, row in input_data.iterrows():
    file_name = row['URL_ID'] + '.txt' 
    url = row['URL']
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    article = soup.find('div', class_='td-post-content')

    if article:
        for script in article(["script", "style", "aside"]):
            script.decompose()

        text = article.get_text(separator="\n").strip()
        output_path = os.path.join(output_directory, file_name)
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(text)

        print(f"Text has been extracted and saved to '{output_path}'")
    else:
        print(f"Could not find the article content for URL: {url}")

print("All articles have been processed.")
