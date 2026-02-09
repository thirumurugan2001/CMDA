import re
import os
import json
import time
import requests
import traceback
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, ElementNotInteractableException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
load_dotenv()

class ZohoCRMAutomatedAuth:

    def __init__(self):
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.redirect_uri = os.getenv("REDIRECT_URL")
        self.org_id = os.getenv("ORG_ID")
        self.email = os.getenv("EMAIL_ADDRESS")
        self.password = os.getenv("PASSWORD")
        self.auth_url = os.getenv("AUTH_URL")
        self.token_url = os.getenv("TOKEN_URL")
        self.api_base_url = os.getenv("API_BASE_URL")
        self.zoho_model_name = os.getenv("ZOHO_MODEL_NAME")
        self.access_token = None
        self.refresh_token = None
        self.token_expires_at = None
        self.token_file = os.getenv("TOKEN_FILE_NAME")
    
    def truncate_field(self, value, max_length):
        """Truncate string to max_length, adding '...' if truncated."""
        if not value or pd.isna(value):
            return ""
        
        value_str = str(value).strip()
        if len(value_str) <= max_length:
            return value_str
        
        # Leave room for ellipsis
        if max_length > 3:
            return value_str[:max_length-3] + "..."
        else:
            return value_str[:max_length]
    
    def setup_driver(self, headless=False):
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless=new")        
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--disable-images")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--start-maximized")        
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")        
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.implicitly_wait(10)
            return driver
        except WebDriverException as e:
            print(f"Error setting up driver: {e}")
            return None
        
    def wait_and_find_element(self, driver, selectors, timeout=30):
        wait = WebDriverWait(driver, timeout)
        for selector_type, selector_value in selectors:
            try:
                if selector_type == "id":
                    element = wait.until(EC.element_to_be_clickable((By.ID, selector_value)))
                elif selector_type == "name":
                    element = wait.until(EC.element_to_be_clickable((By.NAME, selector_value)))
                elif selector_type == "xpath":
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, selector_value)))
                elif selector_type == "css":
                    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector_value)))
                if element.is_displayed() and element.is_enabled():
                    return element, selector_type, selector_value
            except (TimeoutException, ElementNotInteractableException, NoSuchElementException):
                continue
        return None, None, None
    
    def wait_and_find_element_present(self, driver, selectors, timeout=10):
        wait = WebDriverWait(driver, timeout)
        for selector_type, selector_value in selectors:
            try:
                if selector_type == "id":
                    element = wait.until(EC.presence_of_element_located((By.ID, selector_value)))
                elif selector_type == "name":
                    element = wait.until(EC.presence_of_element_located((By.NAME, selector_value)))
                elif selector_type == "xpath":
                    element = wait.until(EC.presence_of_element_located((By.XPATH, selector_value)))
                elif selector_type == "css":
                    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector_value)))
                return element, selector_type, selector_value
            except (TimeoutException, NoSuchElementException):
                continue
        return None, None, None
    
    def safe_click(self, driver, element, description="element"):
        try:
            element.click()
            return True
        except ElementNotInteractableException:
            try:
                driver.execute_script("arguments[0].click();", element)
                return True
            except Exception as e:
                try:
                    ActionChains(driver).move_to_element(element).click().perform()
                    return True
                except Exception as e:
                    print(f"Failed to click {description}: {e}")
                    return False
    
    def safe_send_keys(self, driver, element, text, description="field"):
        try:
            element.clear()  
            time.sleep(0.5)
            element.send_keys(text)
            return True
        except Exception as e:
            try:
                driver.execute_script("arguments[0].value = '';", element)
                driver.execute_script("arguments[0].value = arguments[1];", element, text)
                return True
            except Exception as e:
                print(f"Failed to send keys to {description}: {e}")
                return False
    
    def handle_tfa_banner_page(self, driver):
        current_url = driver.current_url
        if "tfa-banner" in current_url or "announcement" in current_url:
            continue_selectors = [
                ("xpath", "//button[contains(text(), 'Continue')]"),
                ("xpath", "//button[contains(text(), 'Skip')]"),
                ("xpath", "//button[contains(text(), 'Later')]"),
                ("xpath", "//button[contains(text(), 'Not now')]"),
                ("xpath", "//a[contains(text(), 'Continue')]"),
                ("xpath", "//a[contains(text(), 'Skip')]"),
                ("xpath", "//a[contains(text(), 'Later')]"),
                ("xpath", "//input[@value='Continue']"),
                ("xpath", "//input[@value='Skip']"),
                ("id", "continue"),
                ("id", "skip"),
                ("id", "later"),
                ("css", ".continue-btn"),
                ("css", ".skip-btn"),
                ("css", "button.primary"),
                ("css", "a.primary"),
                ("xpath", "//button[contains(@class, 'primary')]"),
                ("xpath", "//a[contains(@class, 'primary')]")
            ]
            element, _, _ = self.wait_and_find_element(driver, continue_selectors, 20)
            if element:
                if self.safe_click(driver, element, "continue/skip button"):
                    time.sleep(3)
                    return True
            else:
                parsed_url = urlparse(current_url)
                query_params = parse_qs(parsed_url.query)
                service_url = query_params.get('serviceurl', [None])[0]
                if service_url:
                    from urllib.parse import unquote
                    decoded_service_url = unquote(service_url)
                    driver.get(decoded_service_url)
                    time.sleep(10)
                    return True
        return False
    
    def is_consent_page(self, driver):
        try:
            consent_indicators = [
                "//div[contains(text(), 'would like to access')]",
                "//div[contains(text(), 'access the following information')]",
                "//div[@id='Approve_Reject']",
                "//input[@id='user-consent-check']",
                "//button[contains(text(), 'Accept')]",
                "//button[contains(text(), 'Reject')]",
                "//div[contains(@class, 'user-consent')]"
            ]
            
            for indicator in consent_indicators:
                try:
                    elements = driver.find_elements(By.XPATH, indicator)
                    if elements and len(elements) > 0:
                        return True
                except:
                    continue
                    
            return False
        except Exception as e:
            print(f"Error checking consent page: {e}")
            return False
    
    def handle_consent_page(self, driver):
        time.sleep(2)        
        consent_checkbox_selectors = [
            ("id", "user-consent-check"),
            ("xpath", "//input[@type='checkbox' and contains(@class, 'trust_check')]"),
            ("xpath", "//input[@type='checkbox' and contains(@name, 'user-consent')]"),
            ("xpath", "//input[@type='checkbox' and contains(@id, 'consent')]"),
            ("css", "input.trust_check"),
            ("xpath", "//input[@type='checkbox']")
        ]        
        checkbox_element, _, _ = self.wait_and_find_element_present(driver, consent_checkbox_selectors, 15)
        if checkbox_element:            
            try:
                if not checkbox_element.is_selected():
                    print("Checkbox not selected, attempting to check it...")                    
                    attempts = [
                        lambda: checkbox_element.click(),
                        lambda: driver.execute_script("arguments[0].click();", checkbox_element),
                        lambda: driver.execute_script("arguments[0].checked = true;", checkbox_element),
                        lambda: ActionChains(driver).move_to_element(checkbox_element).click().perform()
                    ]                    
                    for attempt in attempts:
                        try:
                            attempt()
                            time.sleep(1)
                            if checkbox_element.is_selected():
                                print("✓ Checkbox checked successfully")
                                break
                        except:
                            continue                    
                    if not checkbox_element.is_selected():
                        print("Forcing checkbox check with JavaScript...")
                        driver.execute_script("arguments[0].checked = true;", checkbox_element)
                        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", checkbox_element)
                    time.sleep(1)
                else:
                    print("✓ Checkbox already checked")
            except Exception as e:
                print(f"Error handling checkbox: {e}")
                try:
                    driver.execute_script("arguments[0].checked = true;", checkbox_element)
                    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", checkbox_element)
                except:
                    pass
        else:
            print("Could not find consent checkbox, but will try to proceed anyway")        
        print("Looking for Accept button...")        
        time.sleep(2)        
        accept_selectors = [
            ("xpath", "//button[contains(text(), 'Accept') and not(contains(@class, 'disable-button'))]"),
            ("xpath", "//button[contains(text(), 'Accept') and not(@disabled)]"),
            ("xpath", "//button[contains(text(), 'Accept')]"),
            ("xpath", "//button[contains(text(), 'Allow')]"),
            ("xpath", "//button[contains(text(), 'Authorize')]"),
            ("xpath", "//input[@value='Accept']"),
            ("xpath", "//input[@value='Allow']"),
            ("id", "accept"),
            ("id", "allow"),
            ("css", "button.accept"),
            ("css", "button.allow"),
            ("xpath", "//button[contains(@class, 'accept')]"),
            ("xpath", "//button[contains(@class, 'allow')]"),
            ("xpath", "//button[@class='btn submitApproveForm' and not(contains(@class, 'disable-button'))]")
        ]        
        accept_element = None
        max_attempts = 10        
        for attempt in range(max_attempts):
            accept_element, _, _ = self.wait_and_find_element(driver, accept_selectors, 3)
            if accept_element:
                try:
                    if accept_element.is_enabled():
                        print(f"✓ Found enabled Accept button (attempt {attempt + 1})")
                        break
                    else:
                        print(f"Accept button found but disabled (attempt {attempt + 1})")
                        accept_element = None                        
                        all_accept_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Accept')]")
                        for button in all_accept_buttons:
                            try:
                                driver.execute_script("arguments[0].click();", button)
                                print("✓ Clicked Accept button via JavaScript")
                                return True
                            except:
                                continue                        
                        time.sleep(1)
                except Exception as e:
                    print(f"Error checking button: {e}")
                    time.sleep(1)
            else:
                time.sleep(1)        
        if accept_element:
            print("Attempting to click Accept button...")
            click_success = False            
            click_methods = [
                lambda: accept_element.click(),
                lambda: driver.execute_script("arguments[0].click();", accept_element),
                lambda: ActionChains(driver).move_to_element(accept_element).click().perform(),
                lambda: driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click', { bubbles: true }));", accept_element)
            ]            
            for method in click_methods:
                try:
                    method()
                    click_success = True
                    print("✓ Accept button clicked successfully")
                    time.sleep(3)
                    break
                except Exception as e:
                    print(f"Click method failed: {e}")
                    continue            
            if not click_success:
                print("Trying to find and click any Accept button via JavaScript...")
                try:
                    buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Accept')]")
                    for button in buttons:
                        try:
                            driver.execute_script("arguments[0].click();", button)
                            print("✓ Clicked Accept button via JavaScript fallback")
                            time.sleep(3)
                            return True
                        except:
                            continue
                except:
                    pass
            return click_success
        else:
            print("Could not find Accept button, trying JavaScript to find and click...")
            try:
                script = """
                var buttons = document.querySelectorAll('button');
                for (var i = 0; i < buttons.length; i++) {
                    if (buttons[i].textContent.includes('Accept') || 
                        buttons[i].textContent.includes('Allow') || 
                        buttons[i].textContent.includes('Authorize')) {
                        buttons[i].click();
                        return true;
                    }
                }
                return false;
                """
                result = driver.execute_script(script)
                if result:
                    print("✓ Clicked Accept button via JavaScript search")
                    time.sleep(3)
                    return True
            except Exception as e:
                print(f"JavaScript fallback failed: {e}")
        return False
    
    def get_authorization_url(self):
        params = {
            'scope': 'ZohoCRM.modules.ALL,ZohoCRM.settings.ALL,ZohoCRM.users.ALL',
            'client_id': self.client_id,
            'response_type': 'code',
            'access_type': 'offline',
            'redirect_uri': self.redirect_uri
        }
        auth_url = f"{self.auth_url}?" + "&".join([f"{k}={v}" for k, v in params.items()])
        return auth_url
    
    def automate_oauth_flow(self, headless=False):
        driver = self.setup_driver(headless)
        if not driver:
            print("Failed to setup WebDriver")
            return False
        try:
            auth_url = self.get_authorization_url()
            print(f"Opening auth URL: {auth_url[:100]}...")
            driver.get(auth_url)
            time.sleep(3)            
            email_selectors = [
                ("id", "login_id"),
                ("name", "LOGIN_ID"),
                ("xpath", "//input[@type='email']"),
                ("xpath", "//input[@placeholder='Email ID']"),
                ("css", "input[type='email']"),
                ("xpath", "//input[contains(@class, 'email')]")
            ]            
            email_element, _, _ = self.wait_and_find_element(driver, email_selectors, 30)
            if not email_element:
                print("Email field not found")
                self.debug_page(driver)
                return False            
            print(f"Entering email: {self.email}")
            if not self.safe_send_keys(driver, email_element, self.email, "email field"):
                print("Failed to enter email")
                return False            
            next_selectors = [
                ("id", "nextbtn"),
                ("id", "signin_submit"),
                ("xpath", "//button[contains(text(), 'Next')]"),
                ("xpath", "//button[contains(text(), 'Continue')]"),
                ("xpath", "//input[@value='Next']"),
                ("xpath", "//input[@value='Continue']"),
                ("css", "button[type='submit']"),
                ("xpath", "//button[@type='submit']")
            ]
            next_element, _, _ = self.wait_and_find_element(driver, next_selectors, 20)
            if next_element:
                print("Clicking Next button...")
                if self.safe_click(driver, next_element, "next button"):
                    time.sleep(3)
            password_selectors = [
                ("id", "password"),
                ("name", "PASSWORD"),
                ("xpath", "//input[@type='password']"),
                ("css", "input[type='password']"),
                ("xpath", "//input[contains(@class, 'password')]")
            ]
            password_element, _, _ = self.wait_and_find_element(driver, password_selectors, 30)
            if not password_element:
                print("Password field not found")
                self.debug_page(driver)
                return False
            print("Entering password...")
            if not self.safe_send_keys(driver, password_element, self.password, "password field"):
                print("Failed to enter password")
                return False            
            signin_selectors = [
                ("id", "nextbtn"),
                ("id", "signin_submit"),
                ("xpath", "//button[contains(text(), 'Sign in')]"),
                ("xpath", "//button[contains(text(), 'Sign In')]"),
                ("xpath", "//button[contains(text(), 'Login')]"),
                ("xpath", "//input[@value='Sign in']"),
                ("xpath", "//input[@value='Sign In']"),
                ("css", "button[type='submit']"),
                ("xpath", "//button[@type='submit']")
            ]
            signin_element, _, _ = self.wait_and_find_element(driver, signin_selectors, 20)
            if not signin_element:
                print("Sign in button not found")
                self.debug_page(driver)
                return False
            print("Clicking Sign In button...")
            if not self.safe_click(driver, signin_element, "sign in button"):
                print("Failed to click sign in button")
                return False
            time.sleep(5)            
            if self.handle_tfa_banner_page(driver):
                print("Handled TFA banner page")
                time.sleep(3)            
            time.sleep(3)
            if self.is_consent_page(driver):
                print("Consent page detected, handling it...")
                if self.handle_consent_page(driver):
                    print("✓ Consent page handled successfully")
                else:
                    print("⚠ Could not handle consent page, but will try to continue")
                time.sleep(3)
            else:
                print("No consent page detected, checking for Accept button directly...")                
                accept_selectors = [
                    ("xpath", "//button[contains(text(), 'Accept')]"),
                    ("xpath", "//button[contains(text(), 'Allow')]"),
                    ("xpath", "//button[contains(text(), 'Authorize')]"),
                ]                
                accept_element, _, _ = self.wait_and_find_element(driver, accept_selectors, 10)
                if accept_element:
                    print("Found Accept button, clicking...")
                    if self.safe_click(driver, accept_element, "accept button"):
                        print("✓ Clicked Accept button")
                        time.sleep(3)            
            print("Waiting for redirect to Google with authorization code...")
            max_wait_time = 60
            start_time = time.time()
            while time.time() - start_time < max_wait_time:
                current_url = driver.current_url
                print(f"Current URL: {current_url[:100]}...")                
                if "google.com" in current_url and "code=" in current_url:
                    print("✓ Got redirect to Google with authorization code")
                    break                
                if "tfa-banner" in current_url or "announcement" in current_url:
                    self.handle_tfa_banner_page(driver)                
                if self.is_consent_page(driver):
                    print("Still on consent page, trying to handle again...")
                    self.handle_consent_page(driver)
                time.sleep(2)
            current_url = driver.current_url
            print(f"Final URL: {current_url[:150]}...")
            if "google.com" in current_url and "code=" in current_url:
                parsed_url = urlparse(current_url)
                code = parse_qs(parsed_url.query).get('code', [None])[0]
                if code:
                    print(f"✓ Extracted authorization code: {code[:20]}...")
                    success = self.get_access_token(code)
                    return success
                else:
                    print("❌ No authorization code found in URL")
                    return False
            else:
                print(f"❌ Did not get redirected to Google. Current URL: {current_url}")
                self.debug_page(driver)
                return False
                
        except Exception as e:
            print(f"❌ Error in automate_oauth_flow: {str(e)}")
            traceback.print_exc()
            self.debug_page(driver)
            return False
        finally:
            try:
                print("Closing WebDriver...")
                driver.quit()
            except:
                pass
    
    def debug_page(self, driver):
        try:    
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = f"debug_screenshot_{timestamp}.png"
            driver.save_screenshot(screenshot_path)
            print(f"Screenshot saved to: {screenshot_path}")            
            page_source = driver.page_source[:2000]
            print(f"\nPage source (first 2000 chars):\n{page_source}")
            elements_info = []
            for tag in ['input', 'button', 'a']:
                elements = driver.find_elements(By.TAG_NAME, tag)
                for element in elements[:20]:
                    try:
                        info = {
                            'tag': tag,
                            'id': element.get_attribute('id'),
                            'name': element.get_attribute('name'),
                            'type': element.get_attribute('type'),
                            'class': element.get_attribute('class'),
                            'text': element.text[:50] if element.text else '',
                            'displayed': element.is_displayed(),
                            'enabled': element.is_enabled()
                        }
                        elements_info.append(info)
                    except:
                        continue
            print(f"\nFound {len(elements_info)} elements")
            for i, info in enumerate(elements_info[:10]):  
                print(f"{i+1}. {info}")
        except Exception as e:
            print(f"Error in debug_page: {e}")
    
    def get_access_token(self, authorization_code):        
        print("Getting access token from authorization code...")
        data = {
            'grant_type': 'authorization_code',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'redirect_uri': self.redirect_uri,
            'code': authorization_code
        }
        
        try:
            response = requests.post(self.token_url, data=data)
            print(f"Token request status: {response.status_code}")
            
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                self.refresh_token = token_data['refresh_token']
                expires_in = token_data.get('expires_in', 3600)
                self.token_expires_at = datetime.now() + timedelta(seconds=expires_in)
                
                print(f"✓ Access token obtained successfully")
                print(f"  Token expires in: {expires_in} seconds")
                
                self.save_tokens()
                return True
            else:
                print(f"❌ Failed to get access token. Response: {response.text}")
                return False
        except Exception as e:
            print(f"❌ Exception getting access token: {e}")
            return False
    
    def refresh_access_token(self):
        if not self.refresh_token:
            print("No refresh token available")
            return False
        
        print("Refreshing access token...")
        data = {
            'grant_type': 'refresh_token',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'refresh_token': self.refresh_token
        }
        
        try:
            response = requests.post(self.token_url, data=data)
            print(f"Refresh request status: {response.status_code}")
            
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                expires_in = token_data.get('expires_in', 3600)
                self.token_expires_at = datetime.now() + timedelta(seconds=expires_in)
                
                print(f"✓ Access token refreshed successfully")
                self.save_tokens()
                return True
            else:
                print(f"❌ Failed to refresh token. Response: {response.text}")
                return False
        except Exception as e:
            print(f"❌ Exception refreshing token: {e}")
            return False
    
    def save_tokens(self):
        token_data = {
            'access_token': self.access_token,
            'refresh_token': self.refresh_token,
            'expires_at': self.token_expires_at.isoformat() if self.token_expires_at else None,
            'client_id': self.client_id,
            'saved_at': datetime.now().isoformat()
        }
        
        try:
            with open(self.token_file, 'w') as f:
                json.dump(token_data, f, indent=2)
            print(f"✓ Tokens saved to {self.token_file}")
            return True
        except Exception as e:
            print(f"❌ Error saving tokens: {e}")
            return False

    def load_tokens(self):
        if os.path.exists(self.token_file):
            try:
                print(f"Loading tokens from {self.token_file}...")
                with open(self.token_file, 'r') as f:
                    token_data = json.load(f)
                
                self.access_token = token_data.get('access_token')
                self.refresh_token = token_data.get('refresh_token')
                expires_at_str = token_data.get('expires_at')
                
                if expires_at_str:
                    self.token_expires_at = datetime.fromisoformat(expires_at_str)
                
                print(f"✓ Tokens loaded successfully")
                return True
            except Exception as e:
                print(f"❌ Error loading tokens: {e}")
                return False
        else:
            print(f"Token file {self.token_file} not found")
            return False
    
    def ensure_valid_token(self):
        try:
            print("Ensuring valid access token...")            
            if not self.access_token:
                self.load_tokens()              
            if self.token_expires_at and datetime.now() >= self.token_expires_at:
                print("Token expired, attempting refresh...")
                if not self.refresh_access_token():
                    print("Token refresh failed, starting OAuth flow...")
                    return self.automate_oauth_flow()
                return True                    
            if not self.access_token:
                print("No access token found, starting OAuth flow...")
                return self.automate_oauth_flow()
            
            print("✓ Access token is valid")
            return True
        except Exception as e:
            print(f"❌ Error in ensure_valid_token: {str(e)}") 
            return False

    def validate_record_for_zoho(self, formatted_record):
        """Validate record fields against Zoho constraints."""
        errors = []        
        name = formatted_record.get("Name", "")
        if len(str(name)) > 120:
            errors.append(f"Name field exceeds 120 characters: {len(str(name))} chars")
        email = formatted_record.get("Email", "")
        if email and "@" not in email:
            errors.append(f"Invalid email format: {email}")        
        return errors

    def format_record_for_zoho(self, record):
        formatted_record = {}
        field_mapping = {
            "Sales Person": "Lead_Owner", 
            "Email ID": "Email",
            "Mobile No.": "Mobile_Number",
            "Date of permit": "Date_of_Permit",
            "Applicant Name": "Lead_Name",
            "Nature of Development": "Nature_of_Developments",
            "Dwelling Unit Info": "Dwelling_Unit_Info",
            "Reference": "Reference",
            "Company_Name": "Company_Name",
            "Architect Name": "Architect",
            "Planning Permission No.": "Plan_Permission",
            "Applicant Address": "Applicant_Address",
            "Future_Projects": "Future_Project", 
            "Creation_Time": "Creation_Time",
            "Which_Brand_Looking_for": "Which_Brand_Looking_for",
            "How_Much_Square_Feet": "How_Much_Square_Feet",
            "Area Name": "Area_Name",  
            "Site Address": "Site_Address"
        }                
        try:
            dwelling_units = record.get("Dwelling Unit Info")
            if dwelling_units is not None and not pd.isna(dwelling_units):
                dwelling_str = str(dwelling_units).strip()
                if dwelling_str and dwelling_str != '' and dwelling_str.lower() != 'nan':
                    try:
                        numbers = re.findall(r'\d+', dwelling_str)
                        if numbers:
                            dwelling_value = int(numbers[0])
                            bathrooms = dwelling_value * 2
                            formatted_record["No_of_bathrooms"] = str(bathrooms)
                        else:
                            formatted_record["No_of_bathrooms"] = "0"
                    except (ValueError, TypeError) as e:
                        formatted_record["No_of_bathrooms"] = "0"
                else:
                    formatted_record["No_of_bathrooms"] = "0"
            else:
                formatted_record["No_of_bathrooms"] = "0"
        except Exception as e:
            traceback.print_exc()
            formatted_record["No_of_bathrooms"] = "0"        
        for excel_field, zoho_field in field_mapping.items():
            if excel_field not in record or record[excel_field] is None or pd.isna(record[excel_field]):
                formatted_record[zoho_field] = ""
                continue
            value = record[excel_field]
            if excel_field in ["Creation_Time", "Date_of_Permit"]:
                if isinstance(value, str):
                    try:
                        dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
                        formatted_record[zoho_field] = dt.strftime("%Y-%m-%dT%H:%M:%S+05:30")
                    except ValueError:
                        try:
                            dt = datetime.strptime(value, "%Y-%m-%d")
                            formatted_record[zoho_field] = dt.strftime("%Y-%m-%d")
                        except ValueError:
                            formatted_record[zoho_field] = str(value)
                elif hasattr(value, "strftime"):
                    formatted_record[zoho_field] = value.strftime("%Y-%m-%dT%H:%M:%S+05:30")
                else:
                    formatted_record[zoho_field] = str(value)            
            elif excel_field in ["Dwelling Unit Info", "How_Much_Square_Feet"]:
                try:
                    numbers = re.findall(r'\d+', str(value))
                    if numbers:
                        formatted_record[zoho_field] = numbers[0]
                    else:
                        formatted_record[zoho_field] = "0"
                except (ValueError, TypeError):
                    formatted_record[zoho_field] = "0"            
            elif excel_field == "Email ID":
                email_str = str(value).strip()
                if "@" in email_str and "." in email_str:
                    formatted_record[zoho_field] = email_str
                else:
                    formatted_record[zoho_field] = "" 
            elif excel_field == "Mobile No.":
                try:
                    mobile_str = str(value).strip()
                    mobile_clean = re.sub(r'[^\d+]', '', mobile_str)
                    formatted_record[zoho_field] = mobile_clean
                except:
                    formatted_record[zoho_field] = str(value).strip()
            else:
                formatted_record[zoho_field] = str(value).strip()
        
        name_value = ""
        if record.get("Applicant Name") and not pd.isna(record.get("Applicant Name")):
            name_value = str(record["Applicant Name"]).strip()
        elif record.get("Company_Name") and not pd.isna(record.get("Company_Name")):
            name_value = str(record["Company_Name"]).strip()         
        if name_value:
            formatted_record["Name"] = self.truncate_field(name_value, 120)
        else:
            formatted_record["Name"] = f"Record_{datetime.now().strftime('%Y%m%d%H%M%S')}"        
        long_fields = ["Applicant_Address", "Site_Address", "Nature_of_Developments", "Architect"]
        for field in long_fields:
            if field in formatted_record and formatted_record[field]:
                formatted_record[field] = self.truncate_field(formatted_record[field], 255)  # Assuming 255 char limit
        
        if "Sales Person" in record and record["Sales Person"] and not pd.isna(record["Sales Person"]):
            sales_person = str(record["Sales Person"]).strip()
            user_id = self.get_user_id_by_name(sales_person)
            if user_id:
                formatted_record["Lead_Owner"] = sales_person
            else:
                pass
        if "No_of_bathrooms" in formatted_record:
            pass
        else:
            pass        
        formatted_record['Lead_Source'] = "Digital Leads"
        return formatted_record

    def push_records_to_zoho(self, records, batch_size=100):
        try : 
            if not self.ensure_valid_token():
                print("Error: Unable to ensure valid access token.")
                return False
            if not records:
                return True 
            total_records = len(records)
            successful_records = 0
            failed_records = 0        
            for i in range(0, total_records, batch_size):
                batch = records[i:i + batch_size]
                formatted_batch = []
                for record in batch:
                    formatted_record = self.format_record_for_zoho(record)
                    if formatted_record:
                        # Validate the record before adding to batch
                        validation_errors = self.validate_record_for_zoho(formatted_record)
                        if validation_errors:
                            print(f"❌ Record validation failed: {validation_errors}")
                            failed_records += 1
                            continue
                        
                        formatted_batch.append(formatted_record)
                if not formatted_batch:
                    continue            
                url = f"{self.api_base_url}/{self.zoho_model_name}"
                headers = {'Authorization': f'Zoho-oauthtoken {self.access_token}','Content-Type': 'application/json'}
                payload = {'data': formatted_batch,'trigger': ['approval', 'workflow', 'blueprint']}
                try:
                    response = requests.post(url, json=payload, headers=headers)
                    if response.status_code == 201:
                        response_data = response.json()
                        batch_success = 0
                        batch_failed = 0
                        if 'data' in response_data:
                            for result in response_data['data']:
                                if result.get('status') == 'success':
                                    batch_success += 1
                                    print(f"✅ Record created successfully: {result.get('message', 'Success')}")
                                else:
                                    batch_failed += 1
                                    error_msg = result.get('message', 'Unknown error')
                                    error_details = result.get('details', 'No details')
                                    print(f"❌ Record failed: {error_msg}")
                                    if isinstance(error_details, dict) and 'api_name' in error_details:
                                        print(f"   Field: {error_details.get('api_name')}")
                                        if 'maximum_length' in error_details:
                                            print(f"   Max length: {error_details.get('maximum_length')}")
                        successful_records += batch_success
                        failed_records += batch_failed
                    else:
                        print(f"❌ HTTP Error {response.status_code}: {response.text}")
                        failed_records += len(formatted_batch)
                    if i + batch_size < total_records:
                        time.sleep(1)  
                except Exception as e:
                    print(f"Error pushing batch to Zoho CRM: {str(e)}")
                    failed_records += len(formatted_batch)        
            print(f"\n✅ Push completed: {successful_records} successful, {failed_records} failed out of {total_records} total")
            return successful_records > 0
        except Exception as e:
            print(f"Error in push_records_to_zoho: {str(e)}")
            return False

    def test_api_connection(self):
        print("Testing API connection...")
        if not self.ensure_valid_token():
            return False
        url = f"{self.api_base_url}/settings/modules"
        headers = {'Authorization': f'Zoho-oauthtoken {self.access_token}','Content-Type': 'application/json'}
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                modules = response.json()
                module_names = [module['api_name'] for module in modules.get('modules', [])]
                if self.zoho_model_name in module_names:
                    print(f"✅ API connection successful! Module '{self.zoho_model_name}' found.")
                    return True
                else:
                    print(f"⚠ Module '{self.zoho_model_name}' not found in available modules.")
                    return False
            else:
                print(f"❌ API connection failed. Status: {response.status_code}")
                return False
        except Exception as e:
            print(f"❌ Error during API connection test: {str(e)}")
            return False

    def get_user_id_by_name(self, sales_person_name):
        if not self.ensure_valid_token():
            return None        
        user_id_mapping = {
            "Abhishek R G": os.getenv("ZOHO_USER_ID_ABHISHEK"), 
            "Karthik": os.getenv("ZOHO_USER_ID_KARTHIK"),  
            "Jagan": os.getenv("ZOHO_USER_ID_JAGAN"),    
            "Dinakaran": os.getenv("ZOHO_USER_ID_DINAKARAN"),
            "Venkatesh": os.getenv("ZOHO_USER_ID_VENKATESH"),
            "Ameen Syed": os.getenv("ZOHO_USER_ID_AMEEN"),
            "Balachander": os.getenv("ZOHO_USER_ID_BALACHANDER"),
            "Vijaya Kumar": os.getenv("ZOHO_USER_ID_VIJAYA_KUMAR"),
        }        
        return user_id_mapping.get(sales_person_name)

    def split_applicant_name(self, applicant_name):
        if not applicant_name or pd.isna(applicant_name):
            return "Unknown", "Applicant"        
        applicant_name = str(applicant_name).strip()        
        if len(applicant_name) > 40:
            company_indicators = ['PVT', 'LTD', 'INC', 'LLC', 'CORP', 'ENTERPRISES', 'COMPANY', 'CO.', 'REP', 'REPRESENTATIVE']
            if any(indicator in applicant_name.upper() for indicator in company_indicators):
                first_part = applicant_name[:40]
                return first_part, "Company"
            else:
                first_name = applicant_name[:40]
                last_name = applicant_name[40:80] if len(applicant_name) > 40 else "Applicant"
                return first_name, last_name
        words = applicant_name.split()        
        if len(words) == 1:
            return words[0], "Applicant"        
        elif len(words) == 2:
            return words[0], words[1]        
        else:
            prefixes = ['MR', 'MRS', 'MS', 'DR', 'PROF']
            if words[0].upper() in prefixes:
                if len(words) >= 3:
                    first_name = words[1]
                    last_name = ' '.join(words[2:])
                else:
                    first_name = words[0]
                    last_name = ' '.join(words[1:])
            else:
                first_name = words[0]
                last_name = ' '.join(words[1:])            
            if len(first_name) > 40:
                first_name = first_name[:40]            
            if len(last_name) > 80:
                last_name = last_name[:77] + "..."            
            return first_name, last_name

    def split_sales_person_name(self, sales_person):
        if not sales_person or pd.isna(sales_person):
            return "Digital", "Lead"        
        sales_person = str(sales_person).strip()
        words = sales_person.split()        
        if len(words) == 1:
            return sales_person, "Digital Lead"
        elif len(words) >= 2:
            name_mapping = {
                "Abhishek R G": ("Abhishek", "R G"),
                "Ameen Syed": ("Ameen", "Syed"),
                "Balachander": ("Balachander", ""),
                "Dinakaran": ("Dinakaran", ""),
                "Jagan": ("Jagan", ""),
                "Karthik": ("Karthik", ""),
                "Venkatesh": ("Venkatesh", ""),
                "Vijaya Kumar": ("Vijaya", "Kumar")
            }            
            if sales_person in name_mapping:
                return name_mapping[sales_person]
            else:
                first_name = words[0]
                last_name = ' '.join(words[1:]) if len(words) > 1 else "Digital Lead"                
                if len(first_name) > 40:
                    first_name = first_name[:40]                
                return first_name, last_name
        else:
            return "Digital", "Lead"

    def create_lead_from_cmda_record(self, cmda_record):
        if not self.ensure_valid_token():
            return False        
        Architect_Name = f"{cmda_record.get('Architect Name', '')} {cmda_record.get('Architect Address', '')} {cmda_record.get('Architect Email', '')}"
        Architect_Name = Architect_Name.replace("nan", "").strip()
        Architect_Name = self.truncate_field(Architect_Name, 255)
        
        How_Much_Square_Feet = cmda_record.get("Dwelling Unit Info", "")
        numbers = re.findall(r'\d+', str(How_Much_Square_Feet))
        if numbers:
            How_Much_Square_Feet = numbers[0]
        else:
            How_Much_Square_Feet = "0"
        How_Much_Square_Feet = int(How_Much_Square_Feet) * 1000        
        sales_person = cmda_record.get("Sales Person", "")
        sales_person_clean = self.clean_value(sales_person)
        owner_id = None         
        if sales_person and self.clean_value(sales_person):
            sales_person_clean = sales_person.strip()
            owner_id = self.get_user_id_by_name(sales_person_clean)
            if owner_id:
                print(f"✅ Found user ID for {sales_person_clean}: {owner_id}")
            else:
                print(f"⚠️ No user ID mapping found for: {sales_person_clean}")        
        applicant_name = cmda_record.get("Applicant Name", "")
        first_name, last_name = self.split_applicant_name(applicant_name)
        
        # Truncate first and last names if needed
        first_name = self.truncate_field(first_name, 40)
        last_name = self.truncate_field(last_name, 80)
        
        lead_data = {
            "Planning_Permission_No": self.truncate_field(self.clean_value(cmda_record.get("Planning Permission No.", "")), 120),
            "Email": self.clean_value(cmda_record.get("Email ID", "")),
            "Phone": self.clean_value(cmda_record.get("Mobile No.", "")),
            "Company": self.truncate_field(cmda_record.get("Applicant Name", ""), 120),
            "First_Name": first_name,
            "Last_Name": last_name,
            "Nature_of_Development": self.truncate_field(self.clean_value(cmda_record.get("Nature of Development", "")), 255),
            "Area_Name": self.truncate_field(self.clean_value(cmda_record.get("Area Name", "")), 120),
            "Site_Address": self.truncate_field(self.clean_value(cmda_record.get("Site Address", "")), 255),
            "Reference": "Digital Lead Abhishek",
            "Architect_Name": Architect_Name,
            "Architect_Phone": self.truncate_field(cmda_record.get("Architect Mobile", "Not Provided"), 30),
            "Lead_Source": "Digital Leads",
            "How_Much_Square_Feet": str(How_Much_Square_Feet),
            "Billing_Area": self.truncate_field(self.clean_value(cmda_record.get("Applicant Address", "")), 255),
        }
        if owner_id:
            lead_data["Owner"] = owner_id
            print(f"🎯 Setting Owner field to: {owner_id}")        
        self.handle_numeric_fields(lead_data, cmda_record)
        self.handle_picklist_fields(lead_data, cmda_record)
        self.handle_date_fields(lead_data, cmda_record)
        lead_data = self.final_data_cleaning(lead_data)        
        print(f"📤 Sending Lead data to Zoho CRM:")
        print(f"   👤 First Name: {first_name}")
        print(f"   👤 Last Name: {last_name}")
        print(f"   🏢 Company: {lead_data.get('Company', 'Not set')}")
        print(f"   🎯 Owner ID: {lead_data.get('Owner', 'Not set')}")
        print(f"   👥 Sales Person: {sales_person_clean}")       
        url = f"{self.api_base_url}/Leads"
        headers = {
            'Authorization': f'Zoho-oauthtoken {self.access_token}',
            'Content-Type': 'application/json'
        }
        payload = {
            'data': [lead_data],
            'trigger': ['workflow']
        }
        try:
            response = requests.post(url, json=payload, headers=headers)            
            print(f"📊 Response Status: {response.status_code}")            
            if response.status_code == 201:
                result = response.json()
                print("✅ Lead created successfully in Zoho CRM!")                
                if 'data' in result:
                    for item in result['data']:
                        if item.get('status') == 'success':
                            lead_id = item.get('details', {}).get('id', 'Unknown')
                            created_by = item.get('details', {}).get('Created_By', {}).get('name', 'Unknown')                            
                            lead_details = self.get_lead_details(lead_id)
                            if lead_details:  
                                owner_name = lead_details.get('Owner', {}).get('name', 'Unknown')
                                print(f"🎉 Lead Created Successfully!")
                                print(f"   📝 Lead ID: {lead_id}")
                                print(f"   👤 Created By: {created_by}")
                                print(f"   🎯 Assigned To (Owner): {owner_name}")
                                print(f"   👤 First Name: {lead_details.get('First_Name', 'Not set')}")
                                print(f"   👤 Last Name: {lead_details.get('Last_Name', 'Not set')}")
                            else:
                                print(f"🎉 Lead Created Successfully! ID: {lead_id}")
                                print(f"👤 Created By: {created_by}")
                        else:
                            error_msg = item.get('message', 'Unknown error')
                            error_details = item.get('details', 'No details')
                            print(f"❌ Lead creation failed: {error_msg}")
                            print(f"🔍 Error Details: {error_details}")
                return True
            else:
                print(f"❌ Failed to create Lead. Status: {response.status_code}")
                print(f"🔍 Response: {response.text}")
                return False
                
        except Exception as e:
            print(f"❌ Error creating Lead in Zoho CRM: {e}")
            traceback.print_exc()
            return False

    def get_lead_details(self, lead_id):
        if not self.ensure_valid_token():
            return None
        url = f"{self.api_base_url}/Leads/{lead_id}"
        headers = {
            'Authorization': f'Zoho-oauthtoken {self.access_token}',
            'Content-Type': 'application/json'
        }        
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                return response.json().get('data', [{}])[0]
            else:
                print(f"⚠️ Could not fetch lead details: {response.status_code}")
                return None
        except Exception as e:
            print(f"⚠️ Error fetching lead details: {e}")
            return None

    def handle_sales_person_assignment(self, lead_data, cmda_record):
        sales_person = cmda_record.get("Sales Person", "")
        if sales_person and self.clean_value(sales_person):
            sales_person_clean = sales_person.strip()            
            user_id = self.get_user_id_by_name(sales_person_clean)
            if user_id:
                lead_data["Owner"] = user_id
                
    def clean_value(self, value):
        if value is None or pd.isna(value):
            return ""        
        value_str = str(value).strip()
        if value_str.lower() in ['', 'nan', 'none', 'null']:
            return ""        
        return value_str

    def handle_numeric_fields(self, lead_data, cmda_record):
        try:
            dwelling_units = cmda_record.get("Dwelling Unit Info", "0")
            if dwelling_units and self.clean_value(dwelling_units):
                numbers = re.findall(r'\d+', str(dwelling_units))
                if numbers:
                    dwelling_value = int(numbers[0])
                    lead_data["No_of_Bathrooms"] = dwelling_value * 2 
                else:
                    lead_data["No_of_Bathrooms"] = 0
            else:
                lead_data["No_of_Bathrooms"] = 0
        except (ValueError, TypeError) as e:
            lead_data["No_of_Bathrooms"] = 0
        
        try:
            dwelling_units = cmda_record.get("Dwelling Unit Info", "0")
            if dwelling_units and self.clean_value(dwelling_units):
                numbers = re.findall(r'\d+', str(dwelling_units))
                if numbers:
                    lead_data["No_of_Units"] = int(numbers[0])
                else:
                    lead_data["No_of_Units"] = 0
            else:
                lead_data["No_of_Units"] = 0
        except (ValueError, TypeError) as e:
            lead_data["No_of_Units"] = 0

    def handle_picklist_fields(self, lead_data, cmda_record):
        which_brand = cmda_record.get("Which_Brand_Looking_for", "")
        if which_brand and self.clean_value(which_brand):
            lead_data["Which_Brand_Looking_for"] = self.truncate_field(self.clean_value(which_brand), 120)        
        if not lead_data.get("Future_Projects"):
            lead_data["Future_Projects"] = "-None-"
        if not lead_data.get("Lead_Source"):
            lead_data["Lead_Source"] = "Digital Leads"

    def handle_date_fields(self, lead_data, cmda_record):
        date_of_permit = cmda_record.get("Date of permit", "")
        if date_of_permit and self.clean_value(date_of_permit):
            try:
                if isinstance(date_of_permit, str) and len(date_of_permit) == 10 and '-' in date_of_permit:
                    dt = datetime.strptime(date_of_permit, "%d-%m-%Y")
                    lead_data["Date_of_Permit"] = dt.strftime("%Y-%m-%d")
                else:
                    lead_data["Date_of_Permit"] = self.clean_value(date_of_permit)
            except ValueError:
                lead_data["Date_of_Permit"] = self.clean_value(date_of_permit)
        date_of_application = cmda_record.get("Date of Application", "")
        if date_of_application and self.clean_value(date_of_application):
            try:
                if isinstance(date_of_application, str) and '/' in date_of_application:
                    dt = datetime.strptime(date_of_application, "%d/%m/%Y")
                    lead_data["Date_of_Application"] = dt.strftime("%Y-%m-%d")
                else:
                    lead_data["Date_of_Application"] = self.clean_value(date_of_application)
            except ValueError:
                lead_data["Date_of_Application"] = self.clean_value(date_of_application)

    def final_data_cleaning(self, lead_data):
        cleaned_data = {}        
        for key, value in lead_data.items():
            if value is None:
                continue
            if isinstance(value, str) and value.strip() == "":
                continue
            if pd.isna(value):
                continue            
            cleaned_data[key] = value        
        return cleaned_data