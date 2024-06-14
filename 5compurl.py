import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook

def get_comparison_urls(product_name):
    compare_urls = []  # Initialize an empty list
    try:
        chrome_service = Service("C:/Users/Admin/Desktop/5compare page url/chromedriver.exe")
        chrome_service.start()
        driver = webdriver.Chrome(service=chrome_service)

        driver.get("https://www.google.com/shopping")
        search_box = driver.find_element(By.NAME, 'q')
        search_box.send_keys(product_name)
        search_box.submit()

        # Wait for the search results to load
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "iXEZD")))

        # Find the compare prices links and get their URLs
        compare_prices_links = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(text(), 'Compare prices')]")))
        
        # Extract URLs from the first 5 links
        for link in compare_prices_links[:5]:
            compare_urls.append(link.get_attribute("href"))

    except Exception as e:
        print(f"Error occurred while searching for '{product_name}':", e)

    return compare_urls

# Read data from input Excel sheet
input_path = "C:/Users/Admin/Desktop/5compare page url/input.xlsx"
input_df = pd.read_excel(input_path)

# Initialize an empty list to store URLs
url_list = []

# Initialize an empty list to store rows for output Excel
output_rows = []

# Iterate over each row in the input DataFrame
for index, row in input_df.iterrows():
    product_name = row["ProductName"]
    
    # Get comparison page URLs
    compare_urls = get_comparison_urls(product_name)
    
    # Append collected URLs to the list
    url_list.append(compare_urls)

    # Write product details for each URL
    for url in compare_urls:
        output_row = row.copy()  # Create a copy of the input row
        output_row["URL"] = url  # Add the URL to the row
        output_rows.append(output_row)

# Combine all rows into a DataFrame
output_df = pd.DataFrame(output_rows)

# Write the DataFrame to Excel without hyperlinks
output_path = "C:/Users/Admin/Desktop/5compare page url/output.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
    output_df.to_excel(writer, index=False)
    

print("Output written to:", output_path)









