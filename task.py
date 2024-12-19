import os
import requests
from bs4 import BeautifulSoup
from zipfile import ZipFile
from openpyxl import Workbook
from typing import List, Dict, Tuple
import time
from concurrent.futures import ThreadPoolExecutor

url = "https://www.scrapethissite.com/pages/forms/"
output_folder = "Output"
zip_file_name = "HTML_Pages.zip"
excel_file = "Hockey_Stats.xlsx"

#create the output folder
def make_output_dir():
    """Ensure the output directory exists."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

#fetch the url and return as string
def fetch_html(url: str) -> str:
    response = requests.get(url)
    response.raise_for_status()
    return response.text

#fetc all html pages from url
def fetch_all_pages() -> List[str]:
    html = fetch_html(url)
    soup = BeautifulSoup(html, 'html.parser')

    # Check the pagination
    pagination = soup.find('ul', class_='pagination')
    last_page = int(pagination.find_all('a')[-2].text.strip())

    # Fetch pages concurrently- It will help to fetch html pages faster
    page_urls = [f"{url}?page={i}" for i in range(1, last_page + 1)]
    with ThreadPoolExecutor() as executor:
        html_pages = list(executor.map(fetch_html, page_urls))

    return html_pages

#Save phtml pages to zip file under the output folder
def save_html_to_zip(html_pages: List[str]):
    zip_path = os.path.join(output_folder, zip_file_name)
    with ZipFile(zip_path, 'w') as zip_file:
        for i, html in enumerate(html_pages, start=1):
            file_name = f"{i}.html"
            zip_file.writestr(file_name, html)
    print(f"HTML pages saved in {zip_path}")

#Parsing hockey data to save in the excel sheets with summary
def parse_hockey_data(html_pages: List[str]) -> Tuple[List[Dict], Dict[int, Tuple[str, int, str, int]]]:
    all_data = []
    summary = {}

    for html in html_pages:
        soup = BeautifulSoup(html, 'html.parser')
        rows = soup.find_all('tr', class_='team')
        for row in rows:
            year = int(row.find('td', class_='year').text.strip())
            team_name = row.find('td', class_='name').text.strip()
            wins = int(row.find('td', class_='wins').text.strip())

            all_data.append({'Year': year, 'Team': team_name, 'Wins': wins})

            if year not in summary:
                summary[year] = (team_name, wins, team_name, wins)
            else:
                winner, max_wins, loser, min_wins = summary[year]
                if wins > max_wins:
                    winner, max_wins = team_name, wins
                if wins < min_wins:
                    loser, min_wins = team_name, wins

                summary[year] = (winner, max_wins, loser, min_wins)

    return all_data, summary

#Save data to the excel sheet
def save_to_excel(data: List[Dict], summary: Dict[int, Tuple[str, int, str, int]]):
    wb = Workbook()

    # Sheet 1: NHL Stats 1990-2011
    sheet1 = wb.active
    sheet1.title = "NHL Stats 1990-2011"
    sheet1.append(["Year", "Team", "Wins"])
    for row in data:
        sheet1.append([row['Year'], row['Team'], row['Wins']])

    # Sheet 2: Winner and Loser per Year
    sheet2 = wb.create_sheet(title="Winner and Loser per Year")
    sheet2.append(["Year", "Winner", "Winner Num. of Wins", "Loser", "Loser Num. of Wins"])
    for year, (winner, max_wins, loser, min_wins) in sorted(summary.items()):
        sheet2.append([year, winner, max_wins, loser, min_wins])

    file_path = os.path.join(output_folder, excel_file)
    wb.save(file_path)
    print(f"Excel file saved at {file_path}")


def main():
    start_time = time.time()
    make_output_dir()

    print("Fetching HTML pages...")
    html_pages = fetch_all_pages()

    print("Saving HTML pages to ZIP file...")
    save_html_to_zip(html_pages)

    print("Parsing hockey data...")
    all_data, summary = parse_hockey_data(html_pages)

    print("Saving data to Excel...")
    save_to_excel(all_data, summary)

    print(f"Task successfully completed in {time.time() - start_time:.2f} seconds.")


if __name__ == "__main__":
    main()
