#
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
with open('log.txt', 'w') as f:
    f.truncate(0)


# Setup logging
logging.basicConfig(
    filename='log.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

class PWPInvoiceAutomation:
    
    def __init__(self, excel_path, login_url, form_url, original_pdf_folder, epr_pdf_folder, output_result_excel):
        self.excel_path = excel_path
        self.login_url = login_url
        self.form_url = form_url
        self.original_pdf_folder = original_pdf_folder
        self.epr_pdf_folder = epr_pdf_folder
        self.output_result_excel = output_result_excel
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
            
            df = pd.read_excel(self.excel_path, sheet_name='Sheet1', dtype=str)
            # Strip spaces from column names
            df.columns = df.columns.str.strip()

            df["Production Date"] = df["Production Date"].apply(self.convert_to_yyyy_mm_dd)
            df["Sales Date"] = df["Sales Date"].apply(self.convert_to_yyyy_mm_dd)

            # Strip spaces from all string values
            self.data = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)


            #self.data = pd.read_excel(self.excel_path, sheet_name='Sheet1')
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
        
    
    def WriteEPRTOPDF(self, eprno, original_pdf_path, output_pdf_path):
        from PyPDF2 import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        import io

        try:
            # Text to insert
            text = eprno

            # Font settings
            font_name = "Helvetica-Bold"     # Options: Helvetica, Times-Roman, Courier, etc.
            font_size = 12                   # Adjust as needed

            # File paths
            existing_pdf_path = f"{original_pdf_path}"
            output_pdf_path = f"{output_pdf_path}"

            # Create a PDF overlay with the text
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)

            # Set font and size
            can.setFont(font_name, font_size)

            # Draw the text at position (x, y) - adjust as needed
            can.drawString(100, 360, text)
            can.save()

            # Move buffer to the beginning
            packet.seek(0)

            # Read the overlay and existing PDF
            overlay_pdf = PdfReader(packet)
            existing_pdf = PdfReader(existing_pdf_path)
            output_pdf = PdfWriter()

            # Merge overlay onto the first page
            first_page = existing_pdf.pages[0]
            first_page.merge_page(overlay_pdf.pages[0])
            output_pdf.add_page(first_page)

            # Copy remaining pages, if any
            for page in existing_pdf.pages[1:]:
                output_pdf.add_page(page)

            # Write output PDF
            with open(output_pdf_path, "wb") as f:
                output_pdf.write(f)
            
            logging.info(f"PDF generated successfully: {output_pdf_path}")
        except Exception as e:
            logging.error(f"PDF creation error: {e}")
            raise
    
    
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


    
    def invoice_form_input_helper_onchange(self, element_id, value_to_fill):
        
        
        js_code="""
                var elem = document.getElementById(arguments[0]);
                elem.focus();
                elem.value = "";
                
                for (let char of arguments[1]) {
                    elem.value += char;
                    elem.dispatchEvent(new KeyboardEvent("keydown", { key: char, bubbles: true }));
                    elem.dispatchEvent(new KeyboardEvent("keypress", { key: char, bubbles: true }));
                    elem.dispatchEvent(new InputEvent("input", { data: char, inputType: "insertText", bubbles: true }));
                    elem.dispatchEvent(new KeyboardEvent("keyup", { key: char, bubbles: true }));
                }
                
                elem.dispatchEvent(new Event("change", { bubbles: true }));
                elem.dispatchEvent(new Event("blur", { bubbles: true }));
                """


        
        self.driver.execute_script(js_code, element_id, value_to_fill)   

    
    def invoice_form_fill_combo_helper(self, xpaths=[], option_to_select="", field_name=""):
        
        for xpath in xpaths:
            try:
                
                combo_input_xpath = xpath
                

                # js_script = """
                # const callback = arguments[arguments.length - 1];

                # function getElementByXPath(xpath) {
                    # return document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                # }

                # (async function selectComboOption(comboInputXPath, optionToSelect) {
                    # const comboBox = getElementByXPath(comboInputXPath);
                    # if (!comboBox) {
                        # callback("Combo box not found");
                        # return;
                    # }

                    # comboBox.click();
                    # comboBox.value = '';
                    # comboBox.dispatchEvent(new Event('input', { bubbles: true }));
                    # comboBox.value = optionToSelect;
                    # comboBox.dispatchEvent(new Event('input', { bubbles: true }));

                    # const normalizedOptionText = optionToSelect.trim().toLowerCase();
                    # const optionXPath = `//div[contains(@class, 'ng-option')]//span[
                        # translate(normalize-space(substring-after(text(), '')), 
                                  # 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') 
                        # = '${normalizedOptionText}']`;

                    # let attempts = 0;
                    # const maxAttempts = 10;
                    # const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

                    # while (attempts < maxAttempts) {
                        # const optionElem = getElementByXPath(optionXPath);
                        # if (optionElem) {
                            # optionElem.click();
                            # callback("Option selected");
                            # return;
                        # }
                        # await delay(250);
                        # attempts++;
                    # }

                    # callback("Option not found");
                # })(arguments[0], arguments[1]);
                # """
                
                js_script = """
                const callback = arguments[arguments.length - 1];

                function getElementByXPath(xpath) {
                    return document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                }

                (async function selectComboOption(comboInputXPath, optionToSelect) {
                    const comboBox = getElementByXPath(comboInputXPath);
                    if (!comboBox) {
                        callback("Combo box not found");
                        return;
                    }

                    // Step 1: Clear the selection if ng-clear-wrapper exists
                    const ngSelectContainer = comboBox.closest('ng-select');
                    if (ngSelectContainer) {
                        const clearIcon = ngSelectContainer.querySelector('.ng-clear-wrapper');
                        if (clearIcon) {
                            clearIcon.click();
                            await new Promise(resolve => setTimeout(resolve, 300));  // wait for UI to update
                        }
                    }

                    // Step 2: Proceed with selecting the new option
                    comboBox.click();
                    comboBox.value = '';
                    comboBox.dispatchEvent(new Event('input', { bubbles: true }));
                    comboBox.value = optionToSelect;
                    comboBox.dispatchEvent(new Event('input', { bubbles: true }));

                    const normalizedOptionText = optionToSelect.trim().toLowerCase();
                    const optionXPath = `//div[contains(@class, 'ng-option')]//span[
                        translate(normalize-space(substring-after(text(), '')), 
                                  'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') 
                        = '${normalizedOptionText}']`;

                    let attempts = 0;
                    const maxAttempts = 10;
                    const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

                    while (attempts < maxAttempts) {
                        const optionElem = getElementByXPath(optionXPath);
                        if (optionElem) {
                            optionElem.click();
                            callback("Option selected");
                            return;
                        }
                        await delay(250);
                        attempts++;
                    }

                    callback("Option not found");
                })(arguments[0], arguments[1]);
                """


                result = self.driver.execute_async_script(js_script, combo_input_xpath, option_to_select)
                

                break
            except:
                continue
        else:
            logging.error(f"Unable to detect {field_name}")
    
    def is_null_string(self, s):
        return str(s).strip().lower() in ['none', 'null', 'nan', '']
    
    
    def fill_generate_invoice_form(self, row):
        
        try:
            logging.info(f"Production Date: {str(row['Production Date'])}")
            logging.info(f"Qualifying Feed (MT): {str(row['Qualifying Feed (MT)'])}")
            logging.info(f"Quantity Sold (MT): {str(row['Quantity Sold (MT)'])}")
            logging.info(f"Name of the Entity: {str(row['Name of the Entity'])}")
            logging.info(f"Product Type: {str(row['Product Type'])}")
            logging.info(f"Amount of Material Sold: {str(row['Amount of Material Sold'])}")
            logging.info(f"Percentage of Clinker: {str(row['Percentage of Clinker'])}")
            logging.info(f"Address: {str(row['Address'])}")
            logging.info(f"State: {str(row['State'])}")
            logging.info(f"District: {str(row['District'])}")            
            logging.info(f"GST No. of Seller: {str(row['GST No. of Seller'])}")
            logging.info(f"Buyer GST: {str(row['Buyer GST'])}")
            logging.info(f"HSN Code: {str(row['HSN Code'])}")
            logging.info(f"E- Invoice Number: {str(row['E- Invoice Number'])}")
            logging.info(f"Bank Account No.: {str(row['Bank Account No.'])}")
            logging.info(f"IFSC Code: {str(row['IFSC Code'])}")
            logging.info(f"Principal Amount: {str(row['Principal Amount'])}")
            logging.info(f"GST Amount: {str(row['GST Amount'])}")
            logging.info(f"Sales Date: {str(row['Sales Date'])}")
            
            



            EPRnum = ""
            invoice_generation_Status = "No"
            
            if (self.is_null_string(str(row['EPR Invoice Generated'])) or row['EPR Invoice Generated']=="No") :
                
                # Getting number of row items
                checkbox_table_rows = len(self.driver.find_elements(by=By.XPATH, value='.//*[@id="ScrollableSimpleTableBody"]/tr'))
                
                # Maching and selecting the checkbox
                for table_row_idx in range(1, checkbox_table_rows+1):
                    
                    Production_Date = self.driver.find_element(by=By.XPATH, value=f'.//*[@id="ScrollableSimpleTableBody"]/tr[{table_row_idx}]/td[3]/span').text
                    Qualifying_Feed = self.driver.find_element(by=By.XPATH, value=f'.//*[@id="ScrollableSimpleTableBody"]/tr[{table_row_idx}]/td[4]/span').text
                    
                    
                    
                    if str(row['Production Date']) == str(Production_Date) and str(row['Qualifying Feed (MT)']) == str(Qualifying_Feed) :
                        
                        
                        checkbox = self.driver.find_element(by=By.XPATH, value=f'.//*[@id="ScrollableSimpleTableBody"]/tr[{table_row_idx}]/td[2]/input')
                        if checkbox.is_selected():
                            checkbox.click()
                        checkbox.click()
                        break # Selecting first matches and come out from the loop
                
                time.sleep(2)
                
                
                
                
                
                quantity_sold_xpath = '/html/body/app-root/app-epr/app-pwp-sales/div[1]/div/form/div[1]/div/div/div[2]/table/tbody/tr/td[9]/input'
                quantity_sold = self.driver.find_element(By.XPATH, quantity_sold_xpath)
                
                
                quantity_sold_js_code = f"""
                document.querySelector('input[name="qty_product_sold"]').value = "";
                """
                
                self.driver.execute_script(quantity_sold_js_code)
                
                quantity_sold.send_keys(row['Quantity Sold (MT)'])
                time.sleep(2)
                

                entity_name_xpath = f'/html/body/app-root/app-epr/app-pwp-sales/div[1]/div/form/div[2]/div/div/div[2]/div/div[1]/div/input'
                self.invoice_form_input_xpath_helper(entity_name_xpath, str(row['Name of the Entity']))

                
                product_type_xpaths = [
                    '//*[@id="productTypeId"]/div/div/div[2]/input'
                ]
                self.invoice_form_fill_combo_helper(product_type_xpaths, str(row['Product Type']), "Product Type")
                
                
                amt_material_sold_xpath = f'//*[@id="materialSold"]'
                self.invoice_form_input_xpath_helper(amt_material_sold_xpath, str(row['Amount of Material Sold']))

                if str(row['Product Type'])== "Cement":
                    #clickerPercentage_xpath = f'//*[@id="clickerPercentage"]'
                    self.invoice_form_input_helper_onchange("clickerPercentage", str(row['Percentage of Clinker']))
                
                entity_address_xpath = f'/html/body/app-root/app-epr/app-pwp-sales/div[1]/div/form/div[2]/div/div/div[2]/div/div[5]/div/input'
                self.invoice_form_input_xpath_helper(entity_address_xpath, str(row['Address']))
                
                
                state_xpaths = [
                    '/html/body/app-root/app-epr/app-pwp-sales/div[1]/div/form/div[2]/div/div/div[2]/div/div[6]/div/ng-select/div/div/div[2]/input'
                ]
                self.invoice_form_fill_combo_helper(state_xpaths, str(row['State']), "State")
                
                
                
                district_xpaths = [
                    '/html/body/app-root/app-epr/app-pwp-sales/div[1]/div/form/div[2]/div/div/div[2]/div/div[7]/div/ng-select/div/div/div[2]/input'
                ]
                self.invoice_form_fill_combo_helper(district_xpaths, str(row['District']), "District")
                    
                
                # Clear up the common fields and fill
                
                self.invoice_form_input_helper("sellerGst", str(row['GST No. of Seller']))
                self.invoice_form_input_helper("buyerGst", str(row['Buyer GST']))
                self.invoice_form_input_helper("hsnCode", str(row['HSN Code']))
                self.invoice_form_input_helper("invno", str(row['E- Invoice Number']))
                self.invoice_form_input_helper("account_number", str(row['Bank Account No.']))
                self.invoice_form_input_helper("ifsc_code", str(row['IFSC Code']))
                self.invoice_form_input_helper("amount", str(row['Principal Amount']))
                self.invoice_form_input_helper("gst", str(row['GST Amount']))
                
                
                #formatted_date = self.convert_to_yyyy_mm_dd(str(row['Sales Date']))
                formatted_date = str(row['Sales Date'])
                
                
                salesdate_js_code = f"""
                const input = document.querySelector('#salesDate');
                input.value = '{formatted_date}';
                input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                """
                
                self.driver.execute_script(salesdate_js_code)
                
                #time.sleep(20) # Temporary # to delete
                
                #Submit generate EPR button
                #self.driver.find_element(By.XPATH, '//button[@type="submit"]').click() 
                
                js_generate_epr = """
                var eprgenclass = "btn btn-primary me-2";
                var elementepr =document.getElementsByClassName(eprgenclass)[0]
                if(elementepr) elementepr.click();
                """
                self.driver.execute_script(js_generate_epr)
                time.sleep(5)
                        
                ####### ACTUAL START
                time.sleep(1)
                # Click confirm button
                self.driver.find_element(By.XPATH, '//*[@id="openViewConfirm"]/div/div/app-view-and-confirm/div[2]/button[2]').click() 
                
                time.sleep(2)

                

                EPRnum = ""
                for attempt in range(10):
                    #Execute JavaScript to get the EPR number of the element with ID 'invoiceNumberCopy'
                    EPRnum = self.driver.execute_script("return document.getElementById('invoiceNumberCopy').value;")
                    if len(str(EPRnum)) > 0:
                        break  # Stop if a final status is determined
                    time.sleep(1)  # Wait 1 second before retrying
                    logging.info(f'Tried:{attempt}')
                
                logging.info(f"EPR Generated: {EPRnum}")
                
                if len(str(EPRnum)) > 0:
                    original_pdf = f"{self.original_pdf_folder}/{str(row['PDF Name'])}.pdf"
                    output_pdf = f"{self.epr_pdf_folder}/{EPRnum}.pdf"
                    
                    self.WriteEPRTOPDF(EPRnum,original_pdf,output_pdf )
                    invoice_generation_Status = "Yes"
                    
                    
                    time.sleep(2)
                    #Click Reset button 
                    #self.driver.find_element(By.XPATH, '/html/body/app-root/app-epr/app-pwp-sales/div[1]/div/div[3]/div[2]/button').click()
                    
                    js_rest = """
                    var resetclass = "btn btn-primary mt-3";
                    var element =document.getElementsByClassName(resetclass)[0]
                    if(element) element.click();
                    """
                    self.driver.execute_script(js_rest)
                    
                else:
                    invoice_generation_Status = "No"
                    EPRnum = ""
                    logging.error(f"Issue with EPR Number capture")
                    
                ####### ACTUAL END
                
                # # #temporary - FOR TESTING ONLY
                # # root1 = tk.Tk()
                # # root1.withdraw()  
                # # root1.attributes("-topmost", True)
                # # result1 = messagebox.askokcancel("Confirmation", "Do you want to proceed?")

                # # if result1:
                    
                    # # time.sleep(1)
                    # # # Click confirm button
                    # # self.driver.find_element(By.XPATH, '//*[@id="openViewConfirm"]/div/div/app-view-and-confirm/div[2]/button[2]').click() 
                    
                    # # time.sleep(2)

                    # # # Execute JavaScript to get the EPR number of the element with ID 'invoiceNumberCopy'
                    # # EPRnum = self.driver.execute_script("return document.getElementById('invoiceNumberCopy').value;")
                    
                    # # logging.info(f"EPR Generated: {EPRnum}")
                    
                    # # if len(str(EPRnum)) > 0:
                        # # original_pdf = f"inpdffolder/{str(row['PDF Name'])}.pdf"
                        # # output_pdf = f"oppdffolder/{EPRnum}.pdf"
                        
                        # # self.WriteEPRTOPDF(EPRnum,original_pdf,output_pdf )
                        # # invoice_generation_Status = "Yes"
                        
                        
                        # # time.sleep(2)
                        # # #Click Reset button 
                        # # #self.driver.find_element(By.XPATH, '/html/body/app-root/app-epr/app-pwp-sales/div[1]/div/div[3]/div[2]/button').click()
                        
                        # # js_rest = """
                        # # var resetclass = "btn btn-primary mt-3";
                        # # var element =document.getElementsByClassName(resetclass)[0]
                        # # if(element) element.click();
                        # # """
                        # # self.driver.execute_script(js_rest)
                        
                        
                    # # else:
                        # # logging.error(f"Issue with EPR Number capture")

                # # else:
                    # # print("Cancel clicked")

                # # root1.destroy()  # Clean up
            
            
        except Exception as e:
            EPRnum = ""
            invoice_generation_Status = "No"
            logging.error(f"Form fill (fill_generate_invoice_form) error: {e}")
            #raise
        
        finally:
            
            if row["EPR Invoice Generated"]=="Yes":
                invoice_generation_Status = "Yes"
                
            row["EPR Invoice Generated"] = invoice_generation_Status
            row["EPR Number"] = EPRnum
            row["EPR Invoice Uploaded"] =""
            
            
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
                str(row['EPR Number']),
                row['EPR Invoice Uploaded'],
                ]
            ]

            

            df = pd.DataFrame(data, columns=columns_name)
            
            file = self.output_result_excel
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
                self.fill_generate_invoice_form(row)
                #self.submit()
                #time.sleep(2)
                #self.log_success(index)
                logging.info(f"Processing End Row: {index + 1}")
            except Exception as e:
                logging.error(f"Row {index + 1} failed: {e}")
                continue

    def run(self):
        try:
            self.manual_login_once()
            
            self.reuse_cookies_and_open_form()
            self.load_data()
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
    form_url = "https://eprplastic.cpcb.gov.in/#/epr/pwp-sales"        # Form page
    excel_file = r"C:\Users\User\Downloads\Automation\input.xlsx"
    orig_pdf_var = r"C:\Users\User\Downloads\Automation\JSW Cement PWP Sample Data\Invoices"
    epr_pdf_var = r"C:\Users\User\Downloads\Automation\epr_pdf_var"
    output_excel_var = r"C:\Users\User\Downloads\Automation\output.xlsx"


    bot = PWPInvoiceAutomation(excel_file, login_url, form_url, orig_pdf_var, epr_pdf_var, output_excel_var)
    bot.run()
    
    
    # # root = tk.Tk()
    # # root.withdraw()
    # # # Make the messagebox topmost
    # # root.wm_attributes('-topmost', 1)
    # # messagebox.showinfo("Manual Logout", "🔐 Logout manually from the portal.\nClick OK once done.")
    # # root.destroy()
    # # time.sleep(1)
