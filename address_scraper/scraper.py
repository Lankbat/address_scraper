import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

BASE_URL = "https://lookups.melissa.com/home/addresssearch/"

# Prompting user for input
street_name = input("Enter street name: ")
city_name = input("Enter city: ")
state_abbrev = input("Enter state abbreviation: ")
zipcode = input("Enter zipcode: ")

# Constructing the payload
payload = {
    "street": street_name,
    "city": city_name,
    "state": state_abbrev,
    "zip": zipcode
}

# Using requests.Session() to maintain any session cookies
with requests.Session() as session:
    response = session.post(BASE_URL, data=payload)

    # Check for successful request
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Extract address details
        address_results = soup.find_all("tr", class_="item")

        # Create a new workbook and select the active worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Address Data"
        ws.append(["Address", "City", "State", "Zipcode", "Type"])

    # extraction logic
    def safe_extract_text(tag_list, index):
        return tag_list[index].text.strip() if len(tag_list) > index else "N/A"

    for result in address_results:
        td_texts = result.find_all("td", class_="text-left capitalize")
    
        address = safe_extract_text(td_texts, 0)
        city = safe_extract_text(td_texts, 1)

        state_tags = result.find_all("td", class_="text-center")
        state = safe_extract_text(state_tags, 0)
        zipcode = safe_extract_text(state_tags, 1).split()[0]  # Splitting to get just the zip, not the additional link text
        type_ = safe_extract_text(state_tags, 2)

    # Append data to the Excel sheet
        ws.append([address, city, state, zipcode, type_])


        # Save the data to an Excel file
        wb.save("address_data.xlsx")
        print("Data saved to address_data.xlsx")
        
    else:
        print(f"Error {response.status_code}: Unable to fetch data from the website.")

