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

class DOWNLOAD_FARC:

    def __init__(self):

        self.PATH = os.path.abspath(os.path.dirname(__file__))
        self.USER = os.getlogin()
        self.DOWNLOAD_DIR = self.PATH + '\\src\\'
        self.CONFIG_PATH = r'C:\config\config.ini'
        self.TEMPLATE_PATH = self.PATH + r"\golf_farc_template.xlsm"
        
        # create logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

        # create console handler
        ch = logging.StreamHandler()

        #create file handler 
        date = str(datetime.datetime.now().strftime('%d-%b-%Y %H_%M_%S %p'))

        fh = logging.FileHandler(f'{self.PATH}\\debug\\{date}.log',encoding='utf-8')

        # create formatter
        formatter = logging.Formatter('%(asctime)s - %(funcName)s - %(lineno)d - %(levelname)s - %(message)s',datefmt='%d/%b/%Y %I:%M:%S %p')

        # add formatter to ch
        ch.setFormatter(formatter)

        #add formatter to fh
        fh.setFormatter(formatter)

        # add ch to logger
        self.logger.addHandler(ch)

        #add fh to logger
        self.logger.addHandler(fh)

        #config.init file
        self.my_config_parser = configparser.ConfigParser()

        self.my_config_parser.read(self.CONFIG_PATH)

        self.config = { 
        'golf_usr': self.my_config_parser.get('GOLF_FARC','usr'),
        'golf_pwd': self.my_config_parser.get('GOLF_FARC','pwd'),
        'email': self.my_config_parser.get('GOLF_FARC','email'),
        'password': self.my_config_parser.get('GOLF_FARC','password'),
        }
        
        self.logger.info(f'USER : {self.USER}')
        
        #init chrome driver
        self.driver_path = chromedriver_autoinstaller.install()
        self.logger.debug("Check chromedriver updating >>> "+self.driver_path)
        
        # Configure the browser options
        self.options = Options()
        self.options.add_experimental_option("prefs", {
        "download.default_directory": self.DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        })
        
        #init excel
        self.wb = xw.Book(self.TEMPLATE_PATH)
        self.ws = self.wb.sheets('Sheet1')
        self.lRow = self.ws.range("A1").end("down").row

        
    def golf_login(self):
            self.driver=webdriver.Chrome(self.driver_path,options=self.options)
            self.driver.maximize_window()
            self.driver.get('https://golf.fabrinet.co.th/')
            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="text"]'))).send_keys(self.config['golf_usr'])
            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="password"]'))).send_keys(self.config['golf_pwd'])
            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="submit"]'))).click()      

            # get the current window handle
            self.main_window = self.driver.current_window_handle

    def read_excel(self):
        
        #iterate untill last row to get data in each row
        self.farc_dict = dict()
        for row in range(2,self.lRow+1):
            if self.ws[f'D{row}'].value == None:
                self.farc_dict.update({self.ws[f'A{row}'].value : {'sn': self.ws[f'B{row}'].value,
                                                                   'golf_id':self.ws[f'C{row}'].value,
                                                                   'farc_file':self.ws[f'D{row}'].value,
                                                                   'farc_link':self.ws[f'E{row}'].value,
                                                                   'farc_file_cell':f'D{row}',
                                                                   'farc_link_cell':f'E{row}',
                                                                   }})
                                                            

    def download(self):
        
        for farc in self.farc_dict:
            
            if self.farc_dict[farc]['farc_file'] == None:
                
                self.driver.get(f"https://golf.fabrinet.co.th/normaluser/WorkFlow.asp?rnt={self.farc_dict[farc]['golf_id']}")
                
                frame = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//frame[@name="down"]')))
                
                self.driver.switch_to.frame(frame)
                
                #find last view button 
                enable_view_element = WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//input[@type="button"][@value="View"][@enabled]')))
                enable_view_element[-1].click()

                #check whether file was downloaded
                bf_download = set(os.listdir(self.DOWNLOAD_DIR))
                while ( set(os.listdir(self.DOWNLOAD_DIR)) == bf_download or  list(set(os.listdir(self.DOWNLOAD_DIR)) - bf_download)[0].split('.')[-1] == 'crdownload' ):
                    sleep(1)
                
                #rename dowloaded files
                prev_name = self.DOWNLOAD_DIR+ list(set(os.listdir(self.DOWNLOAD_DIR)) - bf_download)[0]
                new_name = self.DOWNLOAD_DIR + f"{self.farc_dict[farc]['sn']} CISCO FARC{len(enable_view_element)-1}" +'.'+prev_name.split('.')[-1]
                os.rename(prev_name,new_name)
                
                # switch to the new window to close
                for handle in self.driver.window_handles:
                    if handle != self.main_window:
                        self.driver.switch_to.window(handle)
                        break
                self.driver.close()
                self.driver.switch_to.window(self.main_window)
                
                #update data in excel
                self.ws[self.farc_dict[farc]['farc_file_cell']].value = new_name


    def main(self):
        self.golf_login()
        self.read_excel()
        self.download()
        
if __name__ == '__main__':
    inst_tst = DOWNLOAD_FARC()
    inst_tst.main()





