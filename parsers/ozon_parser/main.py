import sys
import os
import traceback
import time
import pandas as pd
import logging
from datetime import datetime
from multiprocessing import Process, Manager, active_children
from random import uniform
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import urllib.request
import undetected_chromedriver as uc
from bs4 import BeautifulSoup as bs
# from load_file_interface import send_way_to_file


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
    # for i in range(1):
    #    driver.execute_script("window.scrollBy(0,500)", "")
    #    time.sleep(0.15)
    # time.sleep(1)
    #WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'iq9')))
    product_card_list = driver.find_element(By.ID, 'paginatorContent')
    full_html = product_card_list.get_attribute("innerHTML")
    soup = bs(full_html, features="html.parser")
    #print(soup)
    all_art = soup.findAll(name='a', attrs={'class': 'k2i tile-hover-target'})
    curr_page_load = len(all_art)
    if curr_page_load <= 2:
        all_art = soup.findAll(name='a', attrs={'class': 'im2 tile-hover-target'})
    #print(all_art)

    #print(f"Обьектов на странице {page_index + 1} по запросу <{prompt}> - {curr_page_load}")
    art_count = 0
    for card in all_art:
        main_counter += 1
        link = str(card).split('href="')[1].split('"')[0]
        art = link.split('/')[2].split('-')[-1]
        ad = True if '?advert=' in link else False
        if ad:
            ad_link = link.split('?advert=')[1].split('keywords')[0]
            avt = ad_link.split(';')[-4:-1]
            ad_link = ad_link.split(';')[0]
            # print(art, avt, ad_link)
            if art in art_dict.keys():
                # print(f'Повторение артикула {art}')
                art_count += 1
                # art_dict[art].append(str(main_counter) + ' Страница: ' + str(page_index + 1) + ' Реклама ' + ad_link)
                art_dict[art].append(str(main_counter) + ' Реклама')
            else:
                # art_dict[art] = [str(main_counter) + ' Страница: ' + str(page_index + 1) + ' Реклама ' + ad_link]
                art_dict[art] = [str(main_counter) + ' Реклама']
        else:
            if art in art_dict.keys():
                # print(f'Повторение артикула {art}')
                art_count += 1
                # art_dict[art].append(str(main_counter) + ' Страница: ' + str(page_index))
                art_dict[art].append(str(main_counter))
            else:
                # art_dict[art] = [str(main_counter) + ' Страница: ' + str(page_index)]
                art_dict[art] = [str(main_counter)]
    # print(art_count)


def get_data(prompt, data, chat_id, driverpath):
    global driver, art_dict, curr_page_load, main_counter
    try:
        s = Service(executable_path=driverpath)
        #s = Service(ChromeDriverManager().install())
        #C:\Users\matth\.wdm\drivers\chromedriver\win64\116.0.5845.96
    except ValueError:
        latest_chromedriver_version_url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
        latest_chromedriver_version = urllib.request.urlopen(latest_chromedriver_version_url).read().decode('utf-8')
        s = Service(ChromeDriverManager(driver_version=latest_chromedriver_version).install())
    options = uc.ChromeOptions()
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
    #options.add_argument('--window-size=1050,892')
    prefs = {"profile.default_content_settings.geolocation": 2}
    options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(options=options, driver_executable_path=s.path)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': '''
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Object
                delete window.cdc_adoQpoasnfa76pfcZLmcfl_Proxy
          '''
    })
    driver.maximize_window()
    curr_page_load = 0
    # prompt_atr_dict = {}
    try:
        driver.get("https://www.ozon.ru/")
        time.sleep(rnd())
        #driver.get_screenshot_as_file('headless1.png')
        try:
            driver.find_element(By.XPATH, "/html/body/div[@id='__ozon']/div[@id='layoutPage']/div[@class='b0']/header[@class='vc1 v1c']/div[@id='stickyHeader']/div[@class='vc4']/div[@class='d4-a s2']/div[@class='s3']/div[@class='r4']/div[@class='r6 r7']/div[@class='q8 q9 s']/div[@class='r0 r3 tsBodyM a2-a']/button[@class='a2-a4']/span[@class='a2-b1 a2-c5']/span[@class='a2-e7']").click()
        except NoSuchElementException:
            driver.find_element(By.XPATH, "/html/body/div[@id='__ozon']/div[@id='layoutPage']/div[@class='b0']/div[@class='jd']/div[@class='dj0']/div[@class='dj1']/div[@class='q4']/div[@class='d4-a q2']/div[@class='p5']/div[@class='p7']/div[@class='o9 p q0']/div[@class='p1 p4 tsBody400Small a2-a']/button[@class='a2-a4']/span[@class='a2-b1 a2-d6']/span[@class='a2-e7']").click()
        time.sleep(2)
        try:
            driver.find_element(By.XPATH, "/html/body/div[@class='vue-portal-target']/div[@class='d2-a d2-a9']/div[@class='d2-a2']/div[@class='d2-a3 d2-a4']/div[@class='d2-a5']/div[@class='g2']/div[@class='b0 g4']/div/div[@class='u9']/div[@class='v6']/div[@class='v8']/div[@class='p8d dp9']/div[@class='d0q q0d']/div[@class='q3d']/span[@class='dq5']").click()
        except NoSuchElementException:
            driver.find_element(By.XPATH, "/html/body/div[@class='vue-portal-target']/div[@class='d2-a d2-a9']/div[@class='d2-a2']/div[@class='d2-a3 d2-a4']/div[@class='d2-a5']/div[@class='g2']/div[@class='b0 g4']/div/div[@class='u9']/div[@class='v6']/div[@class='v8']/div[@class='p8d dp9']/div[@class='d0q q0d']/div[@class='qd1']/div[@class='d2q tsBody500Medium']").click()
        time.sleep(2)
        try:
            driver.find_element(By.XPATH, "/html/body/div[@class='vue-portal-target']/div[@class='d2-a d2-a9']/div[@class='d2-a2']/div[@class='d2-a3 d2-a4']/div[@class='d2-a5']/div[@class='g2']/div[@class='b0 g4']/div[@class='w4']/div[@class='x']/div[@class='x0']/div[@class='x1']/div[@class='p8d'][1]/div[@class='d0q d1q']").click()
        except NoSuchElementException:
            driver.find_element(By.XPATH, "/html/body/div[@class='vue-portal-target']/div[@class='d2-a d2-a9']/div[@class='d2-a2']/div[@class='d2-a3 d2-a4']/div[@class='d2-a5']/div[@class='g2']/div[@class='b0 g4']/div[@class='w4']/div[@class='x']/div[@class='x0']/div[@class='x1']/div[@class='p8d'][1]/div[@class='d0q d1q']").click()
        time.sleep(rnd())
        # time.sleep(30)
        # print(f'({chat_id})Парсим запрос {prompt}')
        # print(driver.get_window_size())
        art_dict = {}
        input_line = driver.find_element(By.NAME, 'text')
        input_line.clear()
        input_line.send_keys(prompt)
        time.sleep(rnd())
        input_line.send_keys(Keys.ENTER)
        time.sleep(rnd())
        main_counter = 1
        prompt_category = driver.current_url.split('/')[4]
        for i in range(56):
            try_counter = 0
            while True:
                try:
                    # print(driver.current_url)
                    load_page(i, prompt)
                    #if try_counter > 0:
                        #print(f'({chat_id})Успешно')
                    break
                except Exception as ex:
                    try_counter += 1
                    if try_counter <= 5:
                        #print(f'({chat_id})Ошибка при попытке парсинга. Запрос: {prompt} Страница: {i + 1}')
                        traceback.print_exc()
                        #print(f'({chat_id})Повторяю попытку {try_counter}/5')
                        driver.get(
                            f'https://www.ozon.ru/category/{prompt_category}/?category_was_predicted=true&deny_category_prediction=true&from_global=true&page={i + 1}&text={prompt}')
                    else:
                        logger.warning(f'({chat_id})Не удалось получить данные. Запрос: {prompt} Страница: {i + 1}')
                        logger.warning(ex, exc_info=True)
                        break
            try:
                # print('Переход по ссылке')
                # driver.get(f'https://www.ozon.ru/category/{prompt_category}/?category_was_predicted=true&deny_category_prediction=true&from_global=true&page={i + 2}&text={prompt}')
                try:
                    driver.find_element(By.XPATH,
                                    "/html/body/div[@id='__ozon']/div[@id='layoutPage']/div[@class='b0']/div[@class='container b4']/div[@class='e0'][2]/div[@class='c7'][2]/div[@class='er4']/div[@class='er5']/div[@class='u5w']/div[@class='wu5']/div[@class='w1u a2-a']/a[@class='a2-a4']/div[@class='a2-b1 a2-c']").click()
                except NoSuchElementException:
                    driver.find_element(By.XPATH, "/html/body/div[@id='__ozon']/div[@id='layoutPage']/div[@class='b0']/div[@class='container b4']/div[@class='e0'][2]/div[@class='c7'][2]/div[@class='r5e']/div[@class='r6e']/div[@class='zu']/div[@class='uz0']/div[@class='y6u a2-a']/a[@class='a2-a4']/div[@class='a2-b1 a2-c']").click()
                time.sleep(1)
            except Exception as ex:
                #traceback.print_exc()
                logger.warning(f'({chat_id})По запросу {prompt} обработано страниц {i + 1}')
                logger.warning(ex, exc_info=True)
                break
        data[prompt] = art_dict
        # print(art_dict)
        time.sleep(rnd())
    except Exception as ex:
        print(ex)
    finally:
        # driver.close()
        driver.quit()
        logger.info(f'({chat_id})Парсинг по запросу "{prompt}" завершен, обработано карточек товара: {main_counter}')


def main(filename, chat_id, max_alowed_process, driverpath):
    dt1 = datetime.now()
    # filename = send_way_to_file()
    df = pd.read_excel(filename, index_col=False, keep_default_na=False)
    prompts = sorted(list(set(list(df['Запрос']))))
    if '' in prompts:
        prompts.remove('')
    len_df = df.shape[0]
    new_column_1 = ['' for generate in range(len_df)]
    new_column_2 = ['' for generate in range(len_df)]
    # print(prompts)
    # print(f"Ожидаемое время парсинга ~{str(2 * len(prompts))} мин.")
    data = {}
    with Manager() as manager:
        m_dict = manager.dict()
        for task_id in range(len(prompts)):
            while len(active_children()) - 1 >= max_alowed_process:
                # print('Ожидание парсинга, текущих процессов: ', len(active_children()) - 1)
                # print('Процессов в очереди: ', len(prompts) - task_id)
                time.sleep(10)
            Process(target=get_data, args=(prompts[task_id], m_dict, chat_id, driverpath)).start()
        while len(active_children()) > 1:
            # print('Ожидание парсинга, текущих процессов: ', len(active_children()) - 1)
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
                #print(dt)
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
            df.to_excel('/'.join(filename.split('/')[:2]) + '/parsed_' + filename.split('/')[2], index=False)
            #df.to_excel('output.xlsx', index=False)
            break
        except Exception:
            print('Ошибка записи данных в Excel файл, закройте файл')
            input('Нажмите Enter чтобы повторить попытку')
            continue
    logger.info(f'({chat_id})Парсинг завершен {str(datetime.now() - dt1)}')


def return_time_ozon(filename, max_alowed_process):
    df = pd.read_excel(filename, index_col=False, keep_default_na=False)
    prompts = len(set(list(df['Запрос'])))
    # print(prompts, round(prompts / 5))
    if prompts > max_alowed_process and prompts % max_alowed_process > 0:
        return (((prompts // 5 + 1) * 2) + 2) * 60
    elif prompts > max_alowed_process and prompts % max_alowed_process == 0:
        return (((prompts // 5) * 2) + 2) * 60
    elif prompts <= max_alowed_process:
        return 4 * 60


def run_parser_ozon(filename, chat_id, max_alowed_process, driverpath):
    try:
        logger.info(f'({chat_id})Started')
        main(filename, chat_id, max_alowed_process, driverpath)
    except Exception:
        try:
            os.mkdir('../../error_logs')
        except FileExistsError:
            pass
        log_date = '_'.join(str(datetime.today()).split('.')[0].split('-')[-1].split(' ')).replace(':', '-')
        file_name = f'../error_logs/ozon_{log_date}.txt'
        err_file = open(file=file_name, mode='w')
        traceback.print_exc(file=err_file)
        err_file.close()


if __name__ == '__main__':
    run_parser_ozon('ТестOzon.xlsx', 1275943662, 1)
