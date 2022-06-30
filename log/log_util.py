import logging
import sys


class LogUtil:
    def log_info(self, name):
        logger = logging.getLogger(name)
        logger.setLevel(logging.INFO)
        try:
            raise Exception
        except:
            f = sys.exc_info()[2].tb_frame.f_back
        formater = logging.Formatter('%(asctime)s %(name)s %(levelname)s '+str(f.f_code.co_filename)+"—>"+str(f.f_code.co_name)+"—>"+str(f.f_lineno)+"行," +'%(message)s')
        if not logger.handlers:
            sh = logging.StreamHandler()
            sh.setFormatter(formater)
            logger.addHandler(sh)
        return logger
    def log_free(self):
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        try:
            raise Exception
        except:
            f = sys.exc_info()[2].tb_frame.f_back
        formater = logging.Formatter('%(asctime)s  %(levelname)s '+str(f.f_code.co_filename)+"—>"+str(f.f_code.co_name)+"—>"+str(f.f_lineno)+"行," +'%(message)s')
        if not logger.handlers:
            sh = logging.StreamHandler()
            sh.setFormatter(formater)
            logger.addHandler(sh)
        return logger

