import os
import time
import logging
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys


# Clear the log at the beginning of each run
with open('uploadlog.txt', 'w') as f:
    f.truncate(0)

# Setup logging
logging.basicConfig(
    filename='uploadlog.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

class PWPUploadAutomation:
    def __init__(self, excel_path, login_url, form_url, epr_files_path, final_result_excel):
        self.excel_path = excel_path
        self.login_url = login_url
        self.form_url = form_url
        self.epr_files_path = epr_files_path
        self.final_result_excel = final_result_excel
        self.cookies = []
        self.driver = None
        self.data = None
        

    def init_driver(self):
        self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    
    def wait_for_user_confirmation(self):
        root = tk.Tk()
        root.withdraw()
        # Make the messagebox topmost
        root.attributes('-topmost', True)
        root.lift()
        root.after(100, lambda: root.attributes('-topmost', True))
        messagebox.showinfo("Manual Login", "🔐 Please complete the login in the browser.\nClick OK once done.", parent=root)
        root.destroy()
        
    def manual_login_once(self):
        try:
            self.init_driver()
            self.driver.get(self.login_url)
            self.wait_for_user_confirmation()
            
            # Capture cookies in memory
            
            self.cookies = self.driver.get_cookies()            
            logging.info("Session cookies stored in memory.")
            
            logging.info(f"Cookie length: {len(self.cookies)}")
            
            # if len(self.driver.get_cookies()) !=0:
                # self.cookies = self.driver.get_cookies()            
                # logging.info("Session cookies stored in memory.")
            # else:
                # logging.info(f"No cookie stored in memory")
                # raise
                
        except Exception as e:
            logging.error(f"Manual login or cookie capture failed: {e}")
            raise

    def reuse_cookies_and_open_form(self):
        try:
            self.driver.get(self.form_url.split('/')[0] + "//" + self.form_url.split('/')[2])  # base domain
            for cookie in self.cookies:
                if 'login_token' in cookie:
                    cookie.pop('login_token')
                try:
                    self.driver.add_cookie(cookie)
                except Exception as e:
                    logging.warning(f"Skipping invalid cookie: {e}")
            self.driver.get(self.form_url)
            #time.sleep(2)
            logging.info("Navigated to form page with active session.")
        except Exception as e:
            logging.error(f"Failed to reuse session cookies: {e}")
            raise

    def load_data(self):
        try:
            self.data = pd.read_excel(self.excel_path, sheet_name='Sheet1', dtype=str)
            logging.info(f"{len(self.data)} rows loaded from Excel.")
        except Exception as e:
            logging.error(f"Excel load error: {e}")
            raise
    
    def log_success(self, row_index):
        with open("success_log.txt", "a") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - Row {row_index + 1} submitted successfully\n")

    
    def convert_to_yyyy_mm_dd(self, date_str):
        for fmt in ("%d %B %Y", "%d %b %Y", "%Y-%m-%d %H:%M:%S"):  # %B = full month, %b = abbreviated month
            try:
                date_obj = datetime.strptime(date_str, fmt)
                return date_obj.strftime("%Y-%m-%d")
            except ValueError:
                continue
        raise ValueError("Date format not recognized")
        
    
    
    def invoice_form_input_xpath_helper(self, element_xpath, value_to_fill):
        
        js_xpath_code = """
        var result = document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
        var elem = result.singleNodeValue;
        if (elem) {
            elem.value = '';
            elem.dispatchEvent(new Event('input', { bubbles: true }));
            elem.value = arguments[1];
            elem.dispatchEvent(new Event('input', { bubbles: true }));
        }
        """
        
        self.driver.execute_script(js_xpath_code, element_xpath, value_to_fill)
        
    
    def invoice_form_input_helper(self, element_id, value_to_fill):
        
        js_code = """
        var elem = document.getElementById(arguments[0]);
        if (elem) {
            elem.value = '';
            elem.dispatchEvent(new Event('input', { bubbles: true }));
            elem.value = arguments[1];
            elem.dispatchEvent(new Event('input', { bubbles: true }));
        }
        """
        
        self.driver.execute_script(js_code, element_id, value_to_fill)

    
    
    def is_null_string(self, s):
        return str(s).strip().lower() in ['none', 'null', 'nan', '']
        
    
    
    def check_modal_visibility(self):
        
        model_script = """
            (function () {
                const modal = document.querySelector('.modal.show');
                if (!modal) return 'hidden';

                const style = window.getComputedStyle(modal);
                if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') {
                    return 'hidden';
                }
                return 'visible';
            })();
        """
        
        return self.driver.execute_script(model_script)
    
    
    
    def get_invoice_upload_status(self):
        
        toast_message = """
        const toast = document.evaluate(
          '//*[@id="toast-container"]/div',
          document,
          null,
          XPathResult.FIRST_ORDERED_NODE_TYPE,
          null
        ).singleNodeValue;

        return toast ? toast.textContent.trim() : null;
        """
        
        message = self.driver.execute_script(toast_message)
        
        if not message:
            return "unknown"  # message is None or empty string
        
        message = message.lower()
    
        success_keywords = ["success", "successfully", "uploaded"]
        failure_keywords = ["error", "not"]

        if any(word in message for word in failure_keywords):
            return "failure"
        elif any(word in message for word in success_keywords):
            return "success"
        else:
            return "unknown"

        
        
    
    def close_popup(self):
        
        self.driver.execute_script("""
          const closeBtn = document.getElementById('closeInvoiceUploadPopup');
          if (closeBtn) closeBtn.click();
        """)
    
    
    def process_upload(self, row):
        
        
               
        try:
            epr_upload_Status = "No"
            
            if (self.is_null_string(str(row['EPR Invoice Uploaded'])) or \
            row['EPR Invoice Uploaded']=="No") and \
            str(row['EPR Invoice Generated'])=="Yes":
            
                file_path = os.path.abspath(f'{self.epr_files_path}/{str(row["EPR Number"])}.pdf')
                
                #fill search box
                search_field_xpath = f'//*[@id="simple-table-with-pagination"]/thead[1]/tr/th/div/div[1]/input'
                self.invoice_form_input_xpath_helper(search_field_xpath, str(row['EPR Number']))
                
                time.sleep(2)
                #click search button
                js_searchbutton = """
                var searchbuttonclass = "col-md-1 btn btn-primary";
                var element =document.getElementsByClassName(searchbuttonclass)[0]
                if(element) element.click();
                """
                self.driver.execute_script(js_searchbutton)
                
                
                time.sleep(8)
                # Open Pop-up                
                js_upload_pop = """
                var uploadiconclass = "fs-12 fa fa-exclamation-triangle color-red fs-15";
                var element =document.getElementsByClassName(uploadiconclass)[0]
                if(element) element.click();
                """
                self.driver.execute_script(js_upload_pop)
                
                time.sleep(3) # Testing
                
                # ##### Push file to file browser [START]
                    
                epr_file_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']")))

                # If it's hidden, make it visible
                self.driver.execute_script("arguments[0].style.display = 'block'; arguments[0].style.opacity = 1;", epr_file_input)

                # Set file path (must be absolute)
                #file_path = os.path.abspath(f'{self.epr_files_path}/{str(row["EPR Number"])}.pdf')
                epr_file_input.send_keys(file_path)
                
                self.driver.execute_script("""
                  var event = new Event('change', { bubbles: true });
                  arguments[0].dispatchEvent(event);
                """, epr_file_input)
                
                
                # ##### Push file to file browser [END]

                time.sleep(5)
                
                ##### Hit Uploadbutton [START]
                js_uploadbutton = """
                var uploadclass = "btn btn-primary float-end m-3";
                var element =document.getElementsByClassName(uploadclass)[0]
                if(element) element.click();
                """
                self.driver.execute_script(js_uploadbutton)
                ##### Hit Uploadbutton [END]
                
                
                ##### Get upload success / failure status [START]
                
                status = "unknown"
                for attempt in range(10):
                    status = self.get_invoice_upload_status()
                    if status in ("success", "failure"):
                        break  # Stop if a final status is determined
                    time.sleep(1)  # Wait 1 second before retrying
                    logging.info(f'Tried:{attempt}')
                
                
                upload_status = status
                
                if upload_status == "success":
                    epr_upload_Status = "Yes"
                else:
                    epr_upload_Status = "No"
                    
                    ## Get Popup visibility
                    if self.check_modal_visibility()== "visible":
                        self.close_popup()
                ##### Get upload success / failure status [END]
                
                
                # # Get uploading Status (Working fine)
                # status = self.driver.execute_script("""
                    # const success = document.querySelector('.fs-12.fa.fa-check-circle.color-active.fs-15');
                    # const failure = document.querySelector('.fs-12.fa.fa-exclamation-triangle.color-red.fs-15');

                    # if (success && window.getComputedStyle(success).display !== 'none') {
                        # return 'Yes';
                    # } else if (failure && window.getComputedStyle(failure).display !== 'none') {
                        # return 'No';
                    # } else {
                        # return 'none';
                    # }
                # """)
            
        except Exception as e:
            epr_upload_Status = "No"
            logging.error(f"EPR Upload error: {e}")
            #raise
        
        finally:
            
            if row["EPR Invoice Uploaded"]=="Yes":
                epr_upload_Status = "Yes"
            
            row["EPR Invoice Uploaded"] =epr_upload_Status
            
            
            columns_name = [
            'Production Date',
            'Quantity Sold (MT)',
            'PDF Name',
            'Name of the Entity',
            'Product Type',
            'Amount of Material Sold',
            'Percentage of Clinker',
            'Address',
            'State',
            'District',
            'GST No. of Seller',
            'Buyer GST',
            'HSN Code',
            'E- Invoice Number',
            'Bank Account No.',
            'IFSC Code',
            'Principal Amount',
            'Total Amount',
            'GST Amount',
            'Sales Date',
            'EPR Invoice Generated',
            'EPR Number',
            'EPR Invoice Uploaded',
            ]
            
            EPR_value = "" if str(row['EPR Number']).strip().lower() == "nan" else str(row['EPR Number'])
            
            data = [
                [
                row['Production Date'],
                str(row['Quantity Sold (MT)']),
                row['PDF Name'],
                row['Name of the Entity'],
                row['Product Type'],
                row['Amount of Material Sold'],
                row['Percentage of Clinker'],
                row['Address'],
                row['State'],
                row['District'],
                row['GST No. of Seller'],
                row['Buyer GST'],
                row['HSN Code'],
                row['E- Invoice Number'],
                row['Bank Account No.'],
                row['IFSC Code'],
                row['Principal Amount'],
                row['Total Amount'],
                row['GST Amount'],
                row['Sales Date'],
                row['EPR Invoice Generated'],
                EPR_value,
                row['EPR Invoice Uploaded'],
                ]
            ]

            

            
            df = pd.DataFrame(data, columns=columns_name)
            
            file = self.final_result_excel
            sheet = 'Sheet1'
            
            df_excel = pd.read_excel(file, sheet_name=sheet)
            non_blank_headers = [col for col in df_excel.columns if pd.notna(col) and str(col).strip() != ""]
            
            if non_blank_headers:
                with pd.ExcelWriter(file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    start_row = pd.read_excel(file, sheet_name=sheet).shape[0] + 1
                    df.to_excel(writer, sheet_name=sheet, index=False, header=False, startrow=start_row)
            else:
                
                with pd.ExcelWriter(file, engine='openpyxl', mode='w') as writer:
                    df.to_excel(writer, sheet_name=sheet, index=False)
    
    
    def submit(self):
        try:
            self.driver.find_element(By.XPATH, '//button[@type="submit"]').click()
        except Exception as e:
            logging.error(f"Submit click failed: {e}")
            raise

    def submit_form(self):
        for index, row in self.data.iterrows():
            try:
                logging.info(f"Processing Start Row: {index + 1}")
                self.process_upload(row)
                #self.submit()
                time.sleep(5)
                
                logging.info(f"Processing End Row: {index + 1}")
            except Exception as e:
                logging.error(f"Row {index + 1} failed: {e}")
                continue
        
    def run(self):
        try:
            self.manual_login_once()
            
            self.reuse_cookies_and_open_form()
            self.load_data()
            time.sleep(5)
            self.submit_form()
        except Exception as e:
            logging.critical(f"Automation aborted: {e}")
        finally:
            if self.driver:
                root = tk.Tk()
                root.withdraw()
                # Make the messagebox topmost
                root.attributes('-topmost', True)
                root.lift()
                root.after(100, lambda: root.attributes('-topmost', True))
                confirmation = messagebox.askokcancel("Manual Logout Confirmation", "🔐 Logout manually from the portal.\nClick OK if want to close the browser.", parent=root)
                if confirmation:
                    root.destroy()
                    self.driver.quit()
                    logging.info("Browser closed.")



if __name__ == "__main__":
    # URLs and file paths
    login_url = "https://eprplastic.cpcb.gov.in/#/plastic/home"      # Manual login page
    form_url = "https://eprplastic.cpcb.gov.in/#/epr/details/sales"        # Form page
    
    
    excel_file = "C:/Users/Sanjoy/Desktop/epr test/Execution Result25062025.xlsx"
    epr_pdf = "C:/Users/Sanjoy/Desktop/epr test/EPR PDF"
    result_excel = "C:/Users/Sanjoy/Desktop/epr test/final execution result.xlsx"

    bot = PWPUploadAutomation(excel_file, login_url, form_url, epr_pdf, result_excel)
    bot.run()
    
    
    # root = tk.Tk()
    # root.withdraw()
    # # Make the messagebox topmost
    # root.wm_attributes('-topmost', 1)
    # messagebox.showinfo("Manual Logout", "🔐 Logout manually from the portal.\nClick OK once done.")
    # root.destroy()
    # time.sleep(1)
