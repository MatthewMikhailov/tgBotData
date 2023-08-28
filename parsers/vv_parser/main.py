import sys
import time
import pandas as pd
from random import uniform
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from tkinter import Tk
from tkinter import filedialog


def send_way_to_file():
    root = Tk()
    ftypes = [('Excel файл', '*.xls')]
    dlg = filedialog.Open(root, filetypes=ftypes)
    fl = dlg.show()
    root.destroy()
    print(fl)
    return fl


def rnd():
    return uniform(0.8, 1.5)


def read_input_excel():
    global df, len_df, brand_data, df_keys
    brand_data = {'Eleven': 'Eleven', 'Спейс': ('ArtSpace', 'OfficeSpace'), 'BERLINGO_': 'Berlingo',
                  'Faber-Castell': 'Faber-Castell',
                  'JOVI': 'JOVI', 'Оригами': 'Origami', 'OfficeClean': 'OfficeClean', 'Koh-I-Noor': 'Koh-I-Noor',
                  'MunHwa': 'MunHwa', 'Schneider': 'Schneider', 'Гамма_художка/хобби': 'Гамма',
                  'Гамма_детство/школа': 'Гамма', 'Мульти-Пульти': 'Мульти-Пульти', 'СТАММ': 'СТАММ', 'Luxor': 'Luxor',
                  'Greenwich Line®': 'Greenwich Line', 'MESHU': 'MESHU', 'БиДжи': 'BG', 'Crown': 'Crown',
                  'Forst': 'Först',
                  'Helmi': 'Helmi', 'Vega': 'Vega', 'ТРИ СОВЫ': 'ТРИ СОВЫ', 'Deffenfer': 'Deffender'}
    df = pd.read_excel(send_way_to_file(), index_col=False, keep_default_na=False)
    #df = pd.read_excel('C:/Users/matth/PycharmProjects/main_tg_bot/vv_parser/test4.xls', index_col=False, keep_default_na=False)
    len_df = df.shape[0]
    df_keys = (df.keys())
    if not ('Кратность' in df_keys):
        print('В УАМ не найден столбец <Кратность>!')
        input('Для закрытия программы нажмите Enter')
        sys.exit()


def fill_card(curr_id):
    #driver.get('https://lkvv.ru/card/detail/7fb3fc2f-1630-4722-8668-a0c88418e542')
    # Создать карточку
    print('Создать карточку')
    driver.get('https://lkvv.ru/card')
    time.sleep(10)
    driver.find_element(By.ID, 'modalCreateCard').click()
    time.sleep(rnd())
    name = driver.find_element(By.ID, 'cardproductmodel-description')
    name.clear()
    name.send_keys(df['Наименование товара'][curr_id])
    #time.sleep(rnd())
    type = driver.find_element(By.ID, 'cardproductmodel-producttype')
    type.click()
    type.send_keys(Keys.ARROW_DOWN)
    type.send_keys(Keys.ENTER)
    time.sleep(rnd())
    create_button = driver.find_element(By.XPATH,
                                        "/html/body[@class='card body modal-open']/div[@id='pageBlock']/main[@id='mainBlock']/div[@id='modal']/div[@class='modal-dialog ']/div[@class='modal-content']/div[@class='modal-body']/form[@id='add-card']/div[@class='form-group text-center']/button[@class='btn btn-primary']")
    create_button.click()
    time.sleep(5)
    # Технолог
    print('Технолог')
    driver.find_element(By.ID, "select2-cardproductmodel-technologist-container").click()
    texnol_input = driver.find_element(By.CLASS_NAME, 'select2-search__field')
    texnol_input.clear()
    texnol_input.send_keys('Воронова')
    time.sleep(15)
    texnol_input.send_keys(Keys.ENTER)
    #time.sleep(rnd())
    # Страна
    print('Страна')
    driver.find_element(By.ID, 'select2-cardproductmodel-country-container').click()
    country_input = driver.find_element(By.CLASS_NAME, 'select2-search__field')
    country_input.clear()
    country_input.send_keys(df['Страна происхождения'][curr_id])
    time.sleep(15)
    country_input.send_keys(Keys.ENTER)
    #time.sleep(rnd())
    # Вес нетто
    print('Вес нетто')
    weight_solo = driver.find_element(By.ID, 'cardproductmodel-nettopack')
    weight_solo.click()
    weight_solo.clear()
    weight_solo.send_keys(str(int(str(df['Вес единицы товара'][curr_id]).replace(',', '').replace('.', ''))))
    time.sleep(rnd())
    # Код ТН ВЭД
    print('Код ТН ВЭД')
    driver.find_element(By.ID, 'select2-cardproductmodel-codeokp-container').click()
    code_input = driver.find_element(By.CLASS_NAME, 'select2-search__field')
    code_input.clear()
    code_input.send_keys(str(df['Код ТН ВЭД'][curr_id]))
    time.sleep(15)
    code_input.send_keys(Keys.ENTER)
    time.sleep(rnd())
    # НДС
    print('НДС')
    if str(df['Ставка НДС'][curr_id]) == '20':
        nds = driver.find_element(By.ID, 'cardproductmodel-stavkacodeokp')
        nds.click()
        time.sleep(rnd())
        nds.send_keys(Keys.ARROW_DOWN)
        nds.send_keys(Keys.ENTER)
    #time.sleep(rnd())
    # Цена
    print('Цена')
    price_input = driver.find_element(By.ID, 'cardproductmodel-buyprice')
    price_input.click()
    price_input.clear()
    price_input.send_keys(str(df['Цена товара'][curr_id]))
    time.sleep(rnd())
    # Кратность
    print('Кратность')
    multiplicity_input = driver.find_element(By.ID, 'cardproductmodel-kvant')
    multiplicity_input.click()
    multiplicity_input.clear()
    multiplicity_input.send_keys(str(df['Кратность'][curr_id]))
    time.sleep(rnd())
    # Срок годности
    print('Срок годности')
    if df['Срок годности '][curr_id] == '':
        driver.find_element(By.XPATH, "/html[@class='type-card']/body[@class='card body']/div[@id='pageBlock']/main[@id='mainBlock']/div[@class='form-wrapper']/section[@id='cardFormSection']/form[@id='cardproductFormId']/div[@id='accordion-description']/div[@class='panel']/div[@id='collapse-description']/div[@class='panel-body']/div[@class='card-data project-data']/div[@class='row']/div[@class='col-sm-12 col-md-12 col-lg-12']/div[@class='form-group field-cardproductmodel-expirationunlim']/div[@class='row']/div[@class='col-sm-12 col-md-12 col-lg-12']/div[@class='form-row ']/div[@class='checkbox']/div[@class='row']/div[@class='col-sm-8 col-md-10 col-lg-11']/label/span[@class='checkbox__text']").click()
        time.sleep(rnd())
    else:
        expiration_date_type = driver.find_element(By.ID, 'cardproductmodel-expirationtype')
        expiration_date_type.click()
        expiration_date_type.send_keys(Keys.ARROW_DOWN)
        expiration_date_type.send_keys(Keys.ENTER)
        time.sleep(rnd())
        expiration_date = driver.find_element(By.ID, 'cardproductmodel-expiration')
        expiration_date.click()
        expiration_date.clear()
        expiration_date.send_keys(str(df['Срок годности '][curr_id]))
        time.sleep(rnd())
        second_expiration_date = driver.find_element(By.ID, 'cardproductmodel-expirationdate')
        second_expiration_date.click()
        second_expiration_date.send_keys(Keys.ARROW_DOWN)
        second_expiration_date.send_keys(Keys.ARROW_DOWN)
        second_expiration_date.send_keys(Keys.ENTER)
        time.sleep(rnd())
    # Штрихкод
    print('Штрихкод')
    barcode_type = driver.find_element(By.ID, 'cardproductmodel-eantype')
    barcode_type.click()
    barcode_type.send_keys(Keys.ARROW_DOWN)
    time.sleep(rnd())
    barcode_type.send_keys(Keys.ENTER)
    time.sleep(rnd())
    barcode_input = driver.find_element(By.ID, 'cardproductmodel-ean')
    barcode_input.click()
    barcode_input.clear()
    barcode_input.send_keys(str(df['Штрих-код единицы товара'][curr_id]))
    time.sleep(rnd())
    # Производитель(Бренд)
    print('Производитель(Бренд)')
    driver.find_element(By.ID, 'select2-cardproductmodel-manufacturer-container').click()
    brand_input = driver.find_element(By.CLASS_NAME, 'select2-search__field')
    brand_input.click()
    try:
        brand_input.send_keys(brand_data[df['Торговая Марка'][curr_id]])
        time.sleep(15)
    except KeyError:
        try:
            brand_input.send_keys(df['Торговая Марка'][curr_id])
            time.sleep(15)
        except Exception:
            pass
        pass
    brand_input.send_keys(Keys.ENTER)
    time.sleep(rnd())
    # Контактный телефон
    print('Контактный телефон')
    num_input = driver.find_element(By.ID, 'cardproductmodel-ordercontactphone')
    num_input.click()
    num_input.clear()
    num_input.send_keys('89106391539')
    time.sleep(rnd())
    # Контактное лицо
    print('Контактное лицо')
    name_input = driver.find_element(By.ID, 'cardproductmodel-ordercontact')
    name_input.click()
    name_input.clear()
    name_input.send_keys('Сбродов Сергей Александрович')
    time.sleep(rnd())
    # Этикетка
    print('Этикетка')
    termocheck_type = driver.find_element(By.ID, 'cardproductmodel-termocheck')
    termocheck_type.click()
    for repeat in range(5):
        termocheck_type.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
    termocheck_type.send_keys(Keys.ENTER)
    time.sleep(rnd())
    # Особенности продукта
    print('Особенности продукта')
    additional_data = driver.find_element(By.ID, 'cardproductmodel-fishkalong')
    additional_data.click()
    additional_data.clear()
    additional_data.send_keys(df['Потребительские свойства'][curr_id])
    time.sleep(rnd())
    # Ширина индивидуальной уп
    print('Ширина индивидуальной уп')
    width_solo = driver.find_element(By.ID, 'cardproductmodel-widthofpacking')
    width_solo.click()
    width_solo.clear()
    width_solo.send_keys(str("%g" % (float(df['Размер товара Ширина, см'][curr_id]) * 10)))
    time.sleep(rnd())
    # Высота индивидуальной уп
    print('Высота индивидуальной уп')
    height_solo = driver.find_element(By.ID, 'cardproductmodel-heightofpacking')
    height_solo.click()
    height_solo.clear()
    height_solo.send_keys(str("%g" % (float(df['Размер товара Высота, см'][curr_id]) * 10)))
    time.sleep(rnd())
    # Глубина индивидуальной уп
    print('Глубина индивидуальной уп')
    len_solo = driver.find_element(By.ID, 'cardproductmodel-depthofpacking')
    len_solo.click()
    len_solo.clear()
    len_solo.send_keys(str("%g" % (float(df['Размер товара Длина, см'][curr_id]) * 10)))
    time.sleep(rnd())
    # Определение размерности упаковки
    print('Определение размерности упаковки')
    multiplicity = str(df['Кратность'][curr_id])
    possible_pack = str(df['Фасовка товара (кол-во в возможных упаковках)'][curr_id]).split()
    print(possible_pack)
    if multiplicity == possible_pack[0]:
        current_pack = 'минимальной'
    elif multiplicity == possible_pack[2]:
        current_pack = 'средней'
    else:
        current_pack = 'максимальной'
    time.sleep(rnd())
    # Вес групп уп
    print('Вес групп уп')
    weight_group = driver.find_element(By.ID, 'cardproductmodel-dimensionswt')
    weight_group.click()
    weight_group.clear()
    weight_group.send_keys(df[f'Вес {current_pack} упаковки'][curr_id])
    time.sleep(rnd())
    # Кратность 2
    print('Кратность 2')
    multiplicity_group = driver.find_element(By.ID, 'cardproductmodel-numberofitemsinbox')
    multiplicity_group.click()
    multiplicity_group.clear()
    multiplicity_group.send_keys(str(df['Кратность'][curr_id]))
    time.sleep(rnd())
    # Длина групп уп
    print('Длина групп уп')
    len_group = driver.find_element(By.ID, 'cardproductmodel-dimensionslength')
    len_group.click()
    len_group.clear()
    len_group.send_keys(str(df[f'Длина {current_pack} упаковки'][curr_id]))
    time.sleep(rnd())
    # Ширина групп уп
    print('Ширина групп уп')
    width_group = driver.find_element(By.ID, 'cardproductmodel-dimensionswidth')
    width_group.click()
    width_group.clear()
    width_group.send_keys(str(df[f'Ширина {current_pack} упаковки'][curr_id]))
    time.sleep(rnd())
    # Высота групп уп
    print('Высота групп уп')
    height_group = driver.find_element(By.ID, 'cardproductmodel-dimensionsheight')
    height_group.click()
    height_group.clear()
    height_group.send_keys(str(df[f'Высота {current_pack} упаковки'][curr_id]))
    time.sleep(rnd())
    # Штрихкод для уп
    print('Штрихкод для уп')
    barcode_group = driver.find_element(By.ID, 'cardproductmodel-barcodethebox1')
    barcode_group.click()
    barcode_group.clear()
    barcode_group.send_keys(str(df[f'Штрих-код {current_pack} упаковки'][curr_id]))
    time.sleep(rnd())
    barcode_group.send_keys(Keys.ENTER)
    time.sleep(rnd())
    # Количество в уп
    print('Количество в уп')
    group_count = driver.find_element(By.ID, 'cardproductmodel-quantitypack1')
    group_count.click()
    group_count.clear()
    group_count.send_keys(str(df['Кратность'][curr_id]))
    time.sleep(rnd())
    # Температура от
    print('Температура от')
    temperature_from = driver.find_element(By.ID, 'cardproductmodel-temperaturestoragefrom')
    temperature_from.click()
    temperature_from.clear()
    temperature_from.send_keys('0')
    time.sleep(rnd())
    # Температура до
    print('Температура до')
    temperature_to = driver.find_element(By.ID, 'cardproductmodel-temperaturestorageto')
    temperature_to.click()
    temperature_to.clear()
    temperature_to.send_keys('25')
    time.sleep(rnd())
    # Склад
    print('Склад')
    driver.find_element(By.XPATH, "/html[@class='type-card']/body[@class='card body']/div[@id='pageBlock']/main[@id='mainBlock']/div[@class='form-wrapper']/section[@id='cardFormSection']/form[@id='cardproductFormId']/div[@id='accordion-informationForPlacingAnOrder']/div[@class='panel']/div[@id='collapse-informationForPlacingAnOrder']/div[@class='panel-body']/div[@class='card-data project-data']/div[@class='row']/div[@class='col-sm-12 col-md-12 col-lg-12']/div[@class='inputs-group'][1]/div[@class='row logistic-object-buttons']/div[@class='col-sm-12']/button[@class='btn btn-primary logistic-object-plus']").click()
    time.sleep(15)
    error_counter = 0
    while True:
        try:
            driver.find_element(By.XPATH, "/html[@class='type-card']/body[@class='card body modal-open']/div[@id='pageBlock']/main[@id='mainBlock']/div[@id='addLogisticObjectModal']/div[@class='modal-dialog ']/div[@class='modal-content']/div[@class='modal-body']/div[@class='mt-20']/button[@class='btn btn-success btn-select-logistic-object form-control white']").click()
            break
        except Exception:
            error_counter += 1
            if error_counter == 5:
                print('Не получилось выбрать склад')
                break
            print(f'Не найдена кнопка выбора склада, повторяю попытку {error_counter}/5')
            time.sleep(5)

    # График заказов
    time.sleep(3)
    print('График заказов')
    order_shedule = driver.find_element(By.XPATH, "/html[@class='type-card']/body[@class='card body']/div[@id='pageBlock']/main[@id='mainBlock']/div[@class='form-wrapper']/section[@id='cardFormSection']/form[@id='cardproductFormId']/div[@id='accordion-informationForPlacingAnOrder']/div[@class='panel']/div[@id='collapse-informationForPlacingAnOrder']/div[@class='panel-body']/div[@class='card-data project-data']/div[@class='row']/div[@class='col-sm-12 col-md-12 col-lg-12']/div[@class='inputs-group'][2]/div[@class='row day-of-order-buttons']/div[@class='col-md-12 col-xs-12']/button[@class='btn btn-primary day-of-order-plus']")
    order_shedule.click()
    time.sleep(1)
    order_shedule.click()
    time.sleep(1)
    day_order1 = driver.find_element(By.ID, 'cardproductmodel-dayoforder1')
    day_order2 = driver.find_element(By.ID, 'cardproductmodel-dayoforder2')
    day_order3 = driver.find_element(By.ID, 'cardproductmodel-dayoforder3')
    day_order1.click()
    day_order1.send_keys(Keys.ARROW_DOWN)
    time.sleep(0.2)
    day_order1.send_keys(Keys.ENTER)
    time.sleep(rnd())
    day_order2.click()
    day_order2.send_keys(Keys.ARROW_DOWN)
    time.sleep(0.2)
    day_order2.send_keys(Keys.ARROW_DOWN)
    time.sleep(0.2)
    day_order2.send_keys(Keys.ENTER)
    time.sleep(rnd())
    day_order3.click()
    for i in range(4):
        day_order3.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
    day_order3.send_keys(Keys.ENTER)
    day_delivery1 = driver.find_element(By.ID, 'cardproductmodel-dayofdelivery1')
    day_delivery2 = driver.find_element(By.ID, 'cardproductmodel-dayofdelivery2')
    day_delivery3 = driver.find_element(By.ID, 'cardproductmodel-dayofdelivery3')
    day_delivery1.click()
    for i in range(4):
        day_delivery1.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
    day_delivery1.send_keys(Keys.ENTER)
    time.sleep(rnd())
    day_delivery2.click()
    for i in range(5):
        day_delivery2.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
    day_delivery2.send_keys(Keys.ENTER)
    time.sleep(rnd())
    day_delivery3.click()
    for i in range(2):
        day_delivery3.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.2)
    day_delivery3.send_keys(Keys.ENTER)
    time.sleep(rnd())
    # Время доставки
    print('Время доставки')
    delivery_time = driver.find_element(By.ID, 'cardproductmodel-deliveryperiodrc')
    delivery_time.click()
    delivery_time.clear()
    delivery_time.send_keys('2')
    time.sleep(rnd())
    # Кратность 3
    print('Кратность 3')
    multiplicity_input_3 = driver.find_element(By.ID, 'cardproductmodel-minordercount')
    multiplicity_input_3.click()
    multiplicity_input_3.clear()
    multiplicity_input_3.send_keys(str(df['Кратность'][curr_id]))
    time.sleep(rnd())
    # Сохранение карточки
    print('Сохранение карточки')
    save_button = driver.find_element(By.XPATH, "/html[@class='type-card']/body[@class='card body']/div[@id='pageBlock']/main[@id='mainBlock']/div[@class='form-wrapper']/section[@id='navigationSection']/div[@class='btn-panel-bottom']/button[@class='btn btn-lg btn-default js-card-save']")
    save_button.click()
    time.sleep(10)


def load_data_to_site():
    global driver, current_line
    s = Service(executable_path='chrome_driver/chromedriver.exe')
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
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
    try:
        driver.get('https://lkvv.ru/login')
        time.sleep(rnd())
        login = driver.find_element(By.ID, 'loginform-username')
        password = driver.find_element(By.ID, 'loginform-password')
        login.clear()
        login.send_keys('6227009062')
        time.sleep(rnd())
        password.clear()
        password.send_keys('961791')
        time.sleep(rnd())
        driver.find_element(By.XPATH, "/html/body/div[@class='loginpage']/div[@class='loginform fadeIn first']/div[@class='formcontent js-trobber-block']/form[@id='w0']/button[@class='b-button primary js-login fadeIn fourth']").click()
        time.sleep(5)
        for i in range(len_df):
            current_line = i
            fill_card(i)
            print(f'Карточка {i + 1} успешно сохранена')
            time.sleep(5)
    except Exception as ex:
        print('Ошибка:')
        print(ex)
        input("Для закрытия нажмите Enter")
    finally:
        time.sleep(10)
        driver.quit()


def main():
    read_input_excel()
    load_data_to_site()


if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        print('Ошибка:')
        print(ex)
        input("--------------------------")
