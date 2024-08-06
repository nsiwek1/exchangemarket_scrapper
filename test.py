from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup
import pandas as pd
import time

# Path to the ChromeDriver executable
chrome_driver_path = 'C:\\Users\\natal\\chromedriver\\chromedriver.exe'  # Update this path

# Initialize the Chrome WebDriver
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service)

# URL of the list of companies
url = "https://gpw.pl/spolki"

def get_company_details(driver, company_url):
    """Retrieve detailed information for a given company URL."""
    driver.get(company_url)
    
    # Wait for the page to load and the "Profil" link to be clickable
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Profil")))

    # Click on the "Profil" section
    try:
        profile_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Profil")))
        profile_link.click()
        
        # Wait for the profile page to load
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'table.footable')))
        
        # Parse the profile page HTML content
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        table = soup.find('table', class_='footable')
        if not table:
            print(f"No data table found for {company_url}.")
            return None
        
        details = {}
        rows = table.find_all('tr')
        for row in rows:
            th = row.find('th')
            td = row.find('td')
            if th and td:
                key = th.get_text(strip=True)
                value = td.get_text(strip=True)
                
                if key == "Nazwa:":
                    details['Name'] = value
                elif key == "Nazwa pełna:":
                    details['Full Name'] = value
                elif key == "Adres siedziby:":
                    details['Address'] = value
                elif key == "Prezes Zarządu:":
                    details['CEO'] = value
                elif key == 'Skrót:':
                    details['Skrót'] = value
                elif key == "E-mail:":
                    email_tag = td.find('a')
                    if email_tag:
                        details['Email'] = email_tag.get_text(strip=True)
        
        return details
    
    except (StaleElementReferenceException, TimeoutException) as e:
        print(f"Error processing company profile at {company_url}: {e}")
        return None

def click_show_more_until_done(driver):
    """Click the 'Pokaż więcej' link until all companies are loaded or until it doesn't add new items."""
    previous_count = 0
    max_attempts = 5
    attempts = 0

    while attempts < max_attempts:
        try:
            show_more_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.more[data-type="pager"]'))
            )
            driver.execute_script("arguments[0].click();", show_more_link)
            time.sleep(2)  # Adjust this delay as needed
            
            # Check if new items were added
            current_count = len(driver.find_elements(By.CSS_SELECTOR, 'table.footable tbody tr'))
            if current_count == previous_count:
                attempts += 1
            else:
                attempts = 0  # Reset attempts if new items are loaded
            previous_count = current_count
        except TimeoutException:
            # If the link is no longer found, break the loop
            break
        except Exception as e:
            print(f"Error clicking 'Pokaż więcej': {e}")
            break

try:
    # Open the main page
    driver.get(url)
    
    # Wait for the company list to load
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'table.footable')))

    # Click "Pokaż więcej" link until all companies are loaded
    click_show_more_until_done(driver)

    # Extract company links from the list
    companies = []
    links = driver.find_elements(By.CSS_SELECTOR, 'table.footable tbody tr a')
    company_links = [(link.text, link.get_attribute('href')) for link in links]

    for company_name, company_url_suffix in company_links:
        company_url = company_url_suffix if company_url_suffix.startswith('http') else f'https://gpw.pl{company_url_suffix}'
        
        retry_count = 0
        max_retries = 3
        success = False
        
        while retry_count < max_retries and not success:
            try:
                print(f"Processing: {company_name} ({company_url})")
                
                # Get company details
                details = get_company_details(driver, company_url)
                if details:
                    details['Company Name'] = company_name
                    details['Company URL'] = company_url
                    companies.append(details)
                    success = True
            except StaleElementReferenceException:
                retry_count += 1
                print(f"Retrying for company {company_name}, attempt {retry_count}")
                time.sleep(1)  # Small delay to allow page to stabilize
            except Exception as e:
                print(f"Error processing company {company_name}: {e}")
                break

    # Save to Excel
    if companies:
        df = pd.DataFrame(companies)
        df.to_excel('companies_list.xlsx', index=False)
        print("Data has been successfully saved to companies_list.xlsx")
    else:
        print("No company data was extracted.")

finally:
    # Close the WebDriver
    driver.quit()
