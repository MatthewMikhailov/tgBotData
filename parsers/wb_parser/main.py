import time
import pandas as pd
import traceback
import logging
import os
from datetime import datetime
from multiprocessing import Process, Manager, active_children
from random import uniform
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
#from pyvirtualdisplay import Display
import urllib.request
from bs4 import BeautifulSoup as bs
#from load_file_interface import send_way_to_file


logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(name)s - %(message)s", level=logging.INFO, datefmt='%H:%M:%S'
)
# set higher logging level for httpx to avoid all GET and POST requests being logged

logger = logging.getLogger(__name__)
#logger.info('Ready')


def rnd():
    return uniform(0.5, 1.5)


def load_page(page_index, prompt):
    global art_dict, curr_page_load, main_counter
    logger.info(f'Парсим страничку {page_index + 1} по запросу <{prompt}>')
    if curr_page_load == 0:
        for i in range(30):
            driver.execute_script("window.scrollBy(0,500)", "")
            time.sleep(0.15)
    elif curr_page_load == 100:
        for i in range(24):
            driver.execute_script("window.scrollBy(0,500)", "")
            time.sleep(0.15)
    else:
        scroll_num = round(curr_page_load / 4.5)
        for i in range(max(scroll_num, 24)):
            driver.execute_script("window.scrollBy(0,500)", "")
            time.sleep(0.15)
    product_card_list = driver.find_element(By.CLASS_NAME, 'product-card-list')
    full_html = product_card_list.get_attribute("innerHTML")
    soup = bs(full_html, features="html.parser")
    all_art = soup.findAll('article')
    curr_page_load = len(all_art)
    #print(f"Обьектов на странице {page_index + 1} по запросу <{prompt}> - {curr_page_load}")
    for i in range(curr_page_load):
        main_counter += 1
        id = str(all_art[i]).split('>')[0].split('data-nm-id=')[1].split()[0][1:-1]
        #print(i, str(all_art[i]).split('>')[0].split('data-nm-id=')[1].split()[0], id)
        if 'product-card--adv' in str(all_art[i]):
            if id in art_dict.keys():
                art_dict[id].append(str(main_counter) + ' Реклама')
            else:
                art_dict[id] = [str(main_counter) + ' Реклама']
        else:
            if id in art_dict.keys():
                art_dict[id].append(str(main_counter))
            else:
                art_dict[id] = [str(main_counter)]


def get_data(prompt, data, chat_id, driverpath):
    global driver, art_dict, curr_page_load, main_counter
    # Linyx only
    #display = Display(visible=0, size=(1050, 892))  
    #display.start()
    #s = Service(executable_path='parsers/drivers/chrome_driver/chromedriver.exe')
    try:
        s = Service(executable_path=driverpath)
        #C:\Users\matth\.wdm\drivers\chromedriver\win64\116.0.5845.96
    except AttributeError:
        s = Service(ChromeDriverManager().install())
        
    options = webdriver.ChromeOptions()
    options.binary_location = '/usr/bin/google-chrome'
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('--no-sandbox')
    options.add_argument('--window-size=1050,892')
    options.add_argument('--headless')
    options.add_argument('--blink-settings=imagesEnabled=false')
    options.add_argument('--disable-background-networking')
    options.add_argument('--disable-default-apps')
    options.add_argument('--disable-sync')
    options.add_argument('--disable-translate')
    options.add_argument('--hide-scrollbars')
    options.add_argument('--metrics-recording-only')
    options.add_argument('--no-first-run')
    options.add_argument('--safebrowsing-disable-auto-update')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-client-side-phishing-detection")
    options.add_argument("--disable-crash-reporter")
    options.add_argument("--disable-oopr-debug-crash-dump")
    options.add_argument("--no-crash-upload")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-low-res-tiling")
    options.add_argument('--no-default-browser-check')
    options.add_argument('--deny-permission-prompts')
    options.add_argument('--disable-notifications')
    driver = webdriver.Chrome(options=options, service=s)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': '''
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Object
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Proxy
          '''
    })
    #driver.maximize_window()
    curr_page_load = 0
    #prompt_atr_dict = {}
    try:
        driver.get("https://www.wildberries.ru/")
        time.sleep(rnd())
        #time.sleep(30)
        #print(f'({chat_id})Парсим запрос {prompt}')
        #print(driver.get_window_size())
        art_dict = {}
        #driver.get_screenshot_as_file('headless1.png')
        input_line = driver.find_element(By.ID, 'searchInput')
        input_line.clear()
        input_line.send_keys(prompt)
        time.sleep(rnd())
        input_line.send_keys(Keys.ENTER)
        time.sleep(rnd())
        main_counter = 1
        for i in range(20):
            try_counter = 0
            while True:
                try:
                    load_page(i, prompt)
                    #if try_counter > 0:
                        #print(f'({chat_id})Успешно')
                    break
                except Exception as ex:
                    try_counter += 1
                    if try_counter <= 5:
                        #print(f'({chat_id})Ошибка при попытке парсинга. Запрос: {prompt} Страница: {i + 1}')
                        #print(ex)
                        #print(f'({chat_id})Повторяю попытку {try_counter}/5')
                        driver.get(f'https://www.wildberries.ru/catalog/0/search.aspx?page={i + 1}&sort=popular&search={prompt}')
                    else:
                        logger.warning(f'({chat_id})Не удалось получить данные. Запрос: {prompt} Страница: {i + 1}')
                        logger.warning(ex, exc_info=True)
                        break
            try:
                #print('Переход по ссылке')
                driver.get(f'https://www.wildberries.ru/catalog/0/search.aspx?page={i + 2}&sort=popular&search={prompt}')
                time.sleep(1)
                f = False
                try:
                    driver.find_element(By.XPATH, "/html[@class='adaptive']/body[@class='ru']/div[@class='wrapper']/main[@id='body-layout']/div[@id='mainContainer']/div[@id='app']/div[2]/div[@class='catalog-page catalog-page--non-search']/div[@class='catalog-page__main']/div[@class='catalog-page__not-found not-found-search']")
                    f = True
                except Exception:
                    pass
                if f:
                    raise Exception
            except Exception as ex:
                logger.info(f'({chat_id})По запросу {prompt} обработано страниц {i + 1}')
                logger.warning(ex, exc_info=True)
                #print(f'({chat_id})По запросу {prompt} обработано страниц {i + 1}')
                break
        data[prompt] = art_dict
        #print(art_dict)
        time.sleep(rnd())
    except Exception as ex:
        print(ex)
    finally:
        #driver.close()
        driver.quit()
        logger.info(f'({chat_id})Парсинг по запросу "{prompt}" завершен, обработано карточек товара: {main_counter}')
        #logger.info(f'({chat_id})Process {Process.name} finished')


def main(filename, chat_id, max_alowed_process, driverpath):
    dt1 = datetime.now()
    #filename = send_way_to_file()
    df = pd.read_excel(filename, index_col=False, keep_default_na=False)
    prompts = sorted(list(set(list(df['Запрос']))))
    if '' in prompts:
        prompts.remove('')
    len_df = df.shape[0]
    new_column_1 = ['' for generate in range(len_df)]
    new_column_2 = ['' for generate in range(len_df)]
    #print(prompts)
    #print(f"Ожидаемое время парсинга ~{str(2 * len(prompts))} мин.")
    data = {}
    with Manager() as manager:
        m_dict = manager.dict()
        for task_id in range(len(prompts)):
            while len(active_children()) - 1 >= max_alowed_process:
                #print('Ожидание парсинга, текущих процессов: ', len(active_children()) - 1)
                #print('Процессов в очереди: ', len(prompts) - task_id)
                time.sleep(10)
            #logger.info(f'({chat_id})Process {len(active_children())}/{max_alowed_process} started')
            Process(target=get_data, args=(prompts[task_id], m_dict, chat_id, driverpath), name=str(len(active_children()) - 1)).start()
        while len(active_children()) > 1:
            #print('Ожидание парсинга, текущих процессов: ', len(active_children()) - 1)
            time.sleep(5)
        for key in m_dict.keys():
            data[key] = m_dict[key]
    for i in range(len_df):
        try:
            art = str(df['Код ВБ'][i])
            prompt = df['Запрос'][i]
            parsed_index = data[prompt].keys()
            if art in parsed_index:
                dt = data[prompt][art]
                new_column_1[i] = str(dt[0])
                if len(dt) > 1:
                    new_column_2[i] = str(dt[1])
            #else:
                #print(f'Товар с артикулом {art} не найден')
        except Exception:
            continue
    #print(f'({chat_id})Запись в Excel')
    df['Индекс_1'] = new_column_1
    df['Индекс_2'] = new_column_2
    while True:
        try:
            #print(filename)
            filepath = '/content/tgBotData/'+ filename.split('/')[-3] + '/' + filename.split('/')[-2] + '/parsed_' + filename.split('/')[-1]
            #print(filepath)
            df.to_excel(filepath, index=False)
            break
        except Exception:
            print('Ошибка записи данных в Excel файл')
            break
    logger.info(f'({chat_id})Парсинг завершен {str(datetime.now() - dt1)}')


def return_time_wb(filename, max_alowed_process):
    df = pd.read_excel(filename, index_col=False, keep_default_na=False)
    prompts = len(set(list(df['Запрос'])))
    #print(prompts, round(prompts / 5))
    if prompts > max_alowed_process and prompts % max_alowed_process > 0:
        return (((prompts // 5 + 1) * 2) + 2) * 60
    elif prompts > max_alowed_process and prompts % max_alowed_process == 0:
        return (((prompts // 5) * 2) + 2) * 60
    elif prompts <= max_alowed_process:
        return 4 * 60


def run_parser_wb(filename, chat_id, max_alowed_process, driverpath):
    try:
        logger.info(f'({chat_id})Started')
        main(filename, chat_id, max_alowed_process, driverpath)
    except Exception:
        try:
            os.mkdir('../error_logs')
        except FileExistsError:
            pass
        log_date = '_'.join(str(datetime.today()).split('.')[0].split('-')[-1].split(' ')).replace(':', '-')
        file_name = f'../error_logs/wb_{log_date}.txt'
        err_file = open(file=file_name, mode='w')
        traceback.print_exc(file=err_file)
        err_file.close()


if __name__ == '__main__':
    run_parser_wb('ТестOzon.xlsx', 1275943662, 1)