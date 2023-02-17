from main_download import DOWNLOAD_FARC
import datetime
import configparser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import xlwings as xw
import traceback
import chromedriver_autoinstaller
from selenium.webdriver.chrome.service import Service
import os
import sys
import logging
from time import sleep
from selenium.webdriver.chrome.options import Options
import pywinauto



class UPLOAD_FARC(DOWNLOAD_FARC):
    
    def qms_login(self):
        self.driver=webdriver.Chrome(self.driver_path,options=self.options)
        self.driver.maximize_window()
        self.driver.get('https://www-plmprd.cisco.com/Agile/')
        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="userInput"]'))).send_keys(self.config["email"])

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="login-button"]'))).click()

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="passwordInput"]'))).send_keys(self.config["password"])

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="login-button"]'))).click()

        count_render_2fa = 0

        while (self.driver.title != "Universal Prompt"):

            sleep(1)

            count_render_2fa+=1

            self.logger.info("Wait for Universal Prompt render:"+str(count_render_2fa))
            
        
        while True:
            
            try:

                WebDriverWait(self.driver, 60).until(ec.visibility_of_element_located((By.XPATH, '//button[@id="trust-browser-button"]'))).click()
                
                break
                
            except:
                
                self.logger.warning('Not found trust browser button')

        two_fa_url=self.driver.current_url

        count_duo_pass = 0

        while(two_fa_url==self.driver.current_url):

            sleep(1)

            count_duo_pass+=1

            self.logger.info("Wait for count_duo_pass:"+str(count_duo_pass))

            if count_duo_pass == 30:

                self.logger.warning("!!! LOGIN TIMEOUT !!!")

                self.driver.quit()

                sys.exit()

        self.logger.info("Login to QIS is success!!!")

        sleep(1)

        self.driver.get('https://www-plmprd.cisco.com/Agile/')

        self.main_page = self.driver.current_window_handle

        self.logger.debug("Main Page:"+self.main_page)

        handles = self.driver.window_handles

        for handle in handles:

            sleep(1)

            self.driver.switch_to.window(handle)

            if self.main_page != self.driver.current_window_handle:

                self.driver.close()

        self.driver.switch_to.window(self.main_page)

        self.driver.maximize_window()

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Collapse Left Navigation"]'))).click()

        return True
    
    def read_excel(self):
        
        #iterate untill last row to get data in each row
        self.farc_dict = dict()
        for row in range(2,self.lRow+1):
            if self.ws[f'D{row}'].value != None and self.ws[f'E{row}'].value == None:
                self.farc_dict.update({self.ws[f'A{row}'].value : {'sn': self.ws[f'B{row}'].value,
                                                                    'golf_id':self.ws[f'C{row}'].value,
                                                                    'farc_file':self.ws[f'D{row}'].value,
                                                                    'farc_link':self.ws[f'E{row}'].value,
                                                                    'farc_file_cell':f'D{row}',
                                                                    'farc_upload_status':f'E{row}',
                                                                    }})
                

    def upload(self,farc_case):

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).clear()

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(farc_case)
        
        WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
        
        #click on attachment btn
        WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-2].click()

        # check duplicate file
        rows_exact = int(WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//strong[@id="totalCount_ATTACHMENTS_FILELIST"]'))).text)
        rows = WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@class="GMDataRow"]')))
        rows_len = len(rows)
        self.logger.info("Exact rows >>> "+str(rows_exact)+" Rows number >>> "+str(rows_len))

        if int(rows_exact)*2 != rows_len:

            self.logger.critical(farc_case + ': Rows number does not match !!!')

            self.err.update({farc_case: 'Rows number does not match'})

        row_start = int(rows_len/2)
        
        for i in range(row_start, rows_len):

            row = rows[i]

            self.logger.debug(row)

            entries = row.find_elements(By.TAG_NAME,'td')

            file_name = entries[5].text.strip()
            
            if file_name in self.farc_dict[farc_case]['farc_file']:
                
                return False

            self.logger.debug("File Name >>> "+file_name)
            
            
            #upload file
            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_AddAttachment_10"]'))).click()
            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="browserFiles"]'))).click()
            sleep(1)
            
            while(True):
                try:
                    app = pywinauto.Application().connect(title="Open")
                    break
                except:
                    self.logger.warning("Wait for Browse Window")
                    sleep(1)
                    
            dlg = app.window(title="Open")
            sleep(1)
            dlg.Edit.set_text(f'"{self.farc_dict[farc_case]["farc_file"]}"')
            dlg.Edit.type_keys("{ENTER}")
            sleep(1)


        WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()
        sleep(1)
        
        #handle for click upload not response
        downloading =True
        while(downloading):

            try:
                close_upload_box = WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="lfuploadpalette_window_close"]'))).click()
                sleep(1)
                self.logger.debug(close_upload_box)
                downloading = False
                self.ws[self.farc_dict[farc_case]['farc_upload_status']].value = 'Done'

            except:

                self.logger.critical("Can not click \"upload\"")
                WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()
                sleep(1)
                
    def get_farc_link(self,farc_case):
        
        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).clear()

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(farc_case)
        
        WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
        
        #click on attachment btn
        WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-2].click()

        # check duplicate file
        rows_exact = int(WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//strong[@id="totalCount_ATTACHMENTS_FILELIST"]'))).text)
        rows = WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@class="GMDataRow"]')))
        rows_len = len(rows)
        self.logger.info("Exact rows >>> "+str(rows_exact)+" Rows number >>> "+str(rows_len))

        if int(rows_exact)*2 != rows_len:

            self.logger.critical(farc_case + ': Rows number does not match !!!')

            self.err.update({farc_case: 'Rows number does not match'})

        row_start = int(rows_len/2)
        
        for i in range(row_start, rows_len):

            row = rows[i]


            self.logger.debug(row)

            entries = row.find_elements(By.TAG_NAME,'td')

            file_name = entries[5].text.strip()
            
            if file_name in self.farc_dict[farc_case]['farc_file']:
                
                # click to select row not work
                # entries[0].click()
                # row.click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_More_10"]'))).click()
        
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[ text() = "Get Shortcut"]'))).click()
                
                #switch to get link window
                handles = self.driver.window_handles
                for handle in handles:
                    sleep(1)
                    self.driver.switch_to.window(handle)
                    if self.main_page != self.driver.current_window_handle:
                        self.ws[self.farc_dict[farc_case]['farc_upload_status']].value = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@id="clip"]'))).text
                        self.driver.close()
                self.driver.switch_to.window(self.main_page)
                self.logger.debug("Found Uploaded File >>> "+file_name)
                
                return True
            
        return False


    def check_farc_status(self,case):

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).clear()

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(case)

        self.logger.info('FARC Case >>> '+case)

        while True:

            try:

                WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

                break
            
            except:

                self.logger.warning("Wait for against click intercepted !!!")

                sleep(1)

        farc_status = WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//h2[@style="color:Blue;"]'))).text

        if farc_status not in ['RMA','Ship']:
            
            return False
        
        not_reset_audit = True

        while(not_reset_audit):

            try :   

                WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//em[@id="MSG_NextStatus_em"]'))).click()

                WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[ text() = "Prelim Analysis" ]'))).click()
                
                WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="ewfinish"]'))).click()
                    
                not_reset_audit = False
                
            except:

                reset_handle = False

                handles = self.driver.window_handles

                self.logger.debug("No. Of window handles1: "+str(len(handles))+", "+str(handles))

                for handle in handles:

                    self.driver.switch_to.window(handle)

                    window_title = self.driver.title

                    if window_title == 'Application Error':

                        self.logger.error("Application Error Window: "+window_title+" , " +handle)

                        reset_handle = True

                        self.driver.close()

                if reset_handle:
                    
                    self.driver.switch_to.window(self.main_page)

                    while True:

                        try:

                            WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

                            break
                        
                        except:

                            self.logger.warning("Wait for against click intercepted !!!")

                            sleep(1)


        #window handle
        found_change_status_window = False
        count_open_change_status_window = 0
        while(not found_change_status_window):

            reset_handle = False
            sleep(1)#9-Dec-2021  add delay
            handles = self.driver.window_handles
            size = len(handles)
            self.logger.debug("No. Of window handles2: "+str(size)+' >>>  '+str(handles))

            for handle in handles:
                
                self.driver.switch_to.window(handle)
                window_title = self.driver.title
                if window_title == 'Change Status':
                    self.logger.debug("Change Status Window: "+window_title+' >>> '+str(handles))
                    found_change_status_window = True
                    break
                elif window_title == 'Application Error':
                    self.logger.error("Application Error Window: "+window_title+' >>> '+str(handles))

                    sleep(1)#9-Dec-2021  add delay
                    reset_handle = True
                    sleep(1)#9-Dec-2021  add delay
                    self.driver.close()

            if reset_handle:
                
                self.driver.switch_to.window(self.main_page)
                
                WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_NextStatusspan"]'))).click()  #9-Dec-2021  visible to clickable
                
                WebDriverWait(self.driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="ewfinish"]'))).click()    #9-Dec-2021
                
            count_open_change_status_window+=3
            
            if count_open_change_status_window > 10:
            
                self.logger.critical("Can't open change_status Window: "+str(count_open_change_status_window)+" second")
                self.can_not_update_state[case] = "Can't Open Change Status Window"
                break

        sleep(5)              

        WebDriverWait(self.driver, 20).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="save"]'))).click()  


        count_close_change_status_window = 0
        while (len(self.driver.window_handles) > 1):
            sleep(1)
            count_close_change_status_window+=1
            self.logger.warning("Wait for Close Change Status Window: "+ str(count_close_change_status_window)+" second")
            if count_close_change_status_window > 10:
                self.logger.critical("Can't Close Change Status Window: "+str(count_close_change_status_window))
                self.can_not_update_state[case] = "Can't Close Change Status Window"
                sleep(1)  #9-Dec-2021  add delay
                self.driver.close()
                sleep(1)  #9-Dec-2021  add delay
                break       
        sleep(1)
        
        self.driver.switch_to.window(self.main_page)    


    def main(self):
        self.qms_login()
        self.read_excel()
        for farc in self.farc_dict:
            self.upload(farc)
            self.check_farc_status(farc)
            
if __name__ == '__main__':
    
    inst_tst = UPLOAD_FARC()
    inst_tst.main()
    