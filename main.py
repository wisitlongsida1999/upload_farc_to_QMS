import logging
import datetime
import configparser
import os
import traceback

class initialize:

    def __init__(self):

        self.PATH = os.getcwd()
        self.USER = os.getlogin()

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

        self.my_config_parser.read(f'{self.path}\\config\\config.ini')

        self.config = { 

        'email': self.my_config_parser.get('config','email'),
        'password': self.my_config_parser.get('config','password'),


        }
        
        self.logger.info(f'USER : {self.USER}')



if __name__ == '__main__':

    try:

        test = initialize()


    finally:

        test.logger.critical("Traceback Error: "+traceback.format_exc())



