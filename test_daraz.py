from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import pandas as pd
import os

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-notifications')
    # Use a standard Chrome User-Agent to avoid getting flagged as a bot
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    driver = webdriver.Chrome(options=options)
    return driver

def scrape_daraz_laptops(max_pages=5):
    driver = setup_driver()
    all_products = []
    
    try:
        print("[*] Navigating to Daraz...")
        driver.get("https://www.daraz.com.bd/")
        
        # Wait for the search box to load
        print("[*] Waiting for the search box to load...")
        search_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "q"))
        )
        
        print("[*] Searching for 'laptop'...")
        search_box.send_keys("laptop")
        
        # Click search
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "search-box__button--1oH7"))
        )
        search_button.click()
        
        # Wait for initial results
        print("[*] Waiting for search results to render...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "[data-qa-locator='product-item']"))
        )
        
        current_page = 1
        
        while current_page <= max_pages:
            print(f"\n[--- Scraping Page {current_page} ---]")
            time.sleep(3) # Let the page load completely
            
            # Scroll down slowly to load lazy-loaded images and elements
            print("[*] Scrolling down to load all products...")
            scroll_height = driver.execute_script("return document.body.scrollHeight")
            for i in range(1, scroll_height, 500):
                driver.execute_script(f"window.scrollTo(0, {i});")
                time.sleep(0.3)
                
            time.sleep(2) # Extra wait after scrolling
            
            # Find all product containers on the page
            products = driver.find_elements(By.CSS_SELECTOR, "[data-qa-locator='product-item']")
            print(f"[*] Found {len(products)} products on page {current_page}")
            
            for index, product in enumerate(products, 1):
                try:
                    # Extract Title and URL
                    try:
                        title_elem = product.find_element(By.CSS_SELECTOR, "a[title]")
                        title = title_elem.get_attribute("title")
                        product_url = title_elem.get_attribute("href")
                    except NoSuchElementException:
                        # Fallback for slightly different HTML structure
                        try:
                            title_elem = product.find_element(By.CSS_SELECTOR, ".title--wFj93")
                            title = title_elem.text
                            product_url = "N/A"
                        except NoSuchElementException:
                            title = "N/A"
                            product_url = "N/A"
                        
                    # Extract Current Price
                    try:
                        price_elem = product.find_element(By.CSS_SELECTOR, ".price--NVB62")
                        price = price_elem.text
                    except NoSuchElementException:
                        price = "N/A"
                        
                    # Extract Original Price (if there is a discount)
                    try:
                        orig_price_elem = product.find_element(By.CSS_SELECTOR, ".origPrice--kZqwe")
                        original_price = orig_price_elem.text
                    except NoSuchElementException:
                        original_price = price # Set to current price if no discount exists
                        
                    # Store in our list
                    all_products.append({
                        "Title": title,
                        "Current Price": price,
                        "Original Price": original_price,
                        "URL": product_url
                    })
                    
                except Exception as e:
                    print(f"[-] Error extracting details for product {index}: {e}")
                    
            print(f"[+] Total products extracted so far: {len(all_products)}")
            
            # Pagination Logic: Check if we should go to the next page
            if current_page >= max_pages:
                print(f"[*] Reached maximum page limit of {max_pages}. Stopping.")
                break
                
            try:
                # Find the "Next" page button by looking for the specific page number
                next_page_num = str(current_page + 1)
                
                # Ant-design pagination usually sets the title of the <li> to the page number
                next_button_container = driver.find_element(By.XPATH, f"//ul[contains(@class, 'ant-pagination')]//li[@title='{next_page_num}']")
                
                print(f"[*] Navigating to page {next_page_num}...")
                
                # The clickable part is usually the <a> inside the <li>
                try:
                    clickable_next = next_button_container.find_element(By.TAG_NAME, "a")
                except NoSuchElementException:
                    # If there's no <a>, the <li> itself might be clickable
                    clickable_next = next_button_container
                
                # Scroll the button into view before clicking
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", clickable_next)
                time.sleep(1)
                
                # Use JavaScript click as fallback if normal click is intercepted
                driver.execute_script("arguments[0].click();", clickable_next)
                
                # Wait for the next page to load by waiting for the new page number to become active
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, f"//ul[contains(@class, 'ant-pagination')]//li[@title='{next_page_num}' and contains(@class, 'ant-pagination-item-active')]"))
                )
                
            except NoSuchElementException:
                print("[-] Could not find the next page number button. Stopping.")
                break
            except Exception as e:
                print(f"[-] Error navigating to next page: {e}")
                break
                
            current_page += 1
            
    except Exception as e:
        print(f"[-] A major error occurred during scraping: {e}")
        
    finally:
        print("[*] Closing the browser...")
        driver.quit()
        
    # Export data to Excel
    if all_products:
        print(f"\n[*] Preparing to save {len(all_products)} products to Excel...")
        df = pd.DataFrame(all_products)
        output_file = "daraz_laptops.xlsx"
        
        # Save in the exact same folder as this script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(script_dir, output_file)
        
        # Write to Excel without the index column
        df.to_excel(output_path, index=False)
        print(f"[+] Success! Data saved to {output_path}")
    else:
        print("[-] No data was extracted. Excel file was not created.")

if __name__ == "__main__":
    print("=== Daraz Multi-Page Laptop Scraper ===")
    print("NOTE: Scraping 5 pages by default to avoid anti-bot blocks.")
    # You can change max_pages=5 to a higher number to scrape more pages.
    scrape_daraz_laptops(max_pages=15)
