import logging
import time

ENCODING = 'utf-8'

# 日志文件
LOG_FILENAME = 'log/log_%s.csv' % time.strftime('%Y%m%d')

# 日志格式
LOG_FORMAT = '%(asctime)s,%(name)s,%(levelname)s,%(message)s'

# 日志日期格式
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

file_handler = logging.FileHandler(LOG_FILENAME, encoding=ENCODING)
stream_handler = logging.StreamHandler()
logging.basicConfig(format=LOG_FORMAT, datefmt=LOG_DATE_FORMAT, level=logging.WARNING,
                    handlers=[file_handler, stream_handler])


def write_log(line, level=logging.WARNING, newline=False):
    """
    写日志并输出到控制台
    :param line:
    :param level:
    :param newline:
    :return:
    """
    if newline:
        line = '\n' + str(line)

    if level == logging.CRITICAL:
        logging.critical(line)
    elif level == logging.ERROR:
        logging.error(line)
    elif level == logging.WARNING:
        logging.warning(line)
    elif level == logging.INFO:
        logging.info(line)
    elif level == logging.DEBUG:
        logging.debug(line)
