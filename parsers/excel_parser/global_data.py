import pandas as pd
import warnings
import difflib
import traceback
import logging
import os
import sys
from datetime import datetime


def similarity(line1, line2):
    return difflib.SequenceMatcher(None, line1, line2).ratio()


def get_dm_brand():
    for dm_brand in dm_brand_data_list:
        for key in brand_data.keys():
            if key != 'Спейс':
                if brand_data[key].lower() == dm_brand.lower():
                    dm_brand_dict[brand_data[key]] = dm_brand
            else:
                if brand_data[key][0].lower() == dm_brand.lower():
                    dm_brand_dict[brand_data[key][0]] = dm_brand
                if brand_data[key][1].lower() == dm_brand.lower():
                    dm_brand_dict[brand_data[key][1]] = dm_brand


warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None


brand_data = {'Eleven': 'Eleven', 'Спейс': ('ArtSpace', 'OfficeSpace'), 'BERLINGO_': 'Berlingo',
                  'Faber-Castell': 'Faber-Castell',
                  'JOVI': 'JOVI', 'Оригами': 'Origami', 'OfficeClean': 'OfficeClean', 'Koh-I-Noor': 'Koh-I-Noor',
                  'MunHwa': 'MunHwa', 'Schneider': 'Schneider', 'Гамма_художка/хобби': 'Гамма',
                  'Гамма_детство/школа': 'Гамма', 'Мульти-Пульти': 'Мульти-Пульти', 'СТАММ': 'СТАММ', 'Luxor': 'Luxor',
                  'Greenwich Line®': 'Greenwich Line', 'MESHU': 'MESHU', 'БиДжи': 'BG', 'Crown': 'Crown', 'Forst': 'Först',
              'Helmi': 'Helmi', 'Vega': 'Vega', 'ТРИ СОВЫ': 'ТРИ СОВЫ', 'Мульти-Пульти_Чебурашка': 'Мульти-Пульти'}

verified_data_mask_dict = {'Код товара РЕЛЬЕФ': 'code', 'Бренд': 'brand', 'Наименование': 'name', 'Серия': 'series',
                           'ТХ': 'tx', 'Ошибки': 'error'}
verified_data_mask = ['Код товара РЕЛЬЕФ', 'Бренд', 'Наименование', 'Серия', 'ТХ', 'Ошибки']

full_input_mask = ['Наименование товара', 'Артикул', 'Размер товара Длина, см', 'Размер товара Ширина, см',
                       'Размер товара Высота, см', 'Вес минимальной упаковки', 'Бренд', 'Штрих-код минимальной упаковки',
                       'Код ТН ВЭД', 'Штрих-код единицы товара', 'Торговая Марка', 'Код товара РЕЛЬЕФ', 'Кратность',
                       'Страна происхождения', 'Вес единицы товара', 'Потребительские свойства', ' Номер документа сертификации',
                   'Ставка НДС', 'Длина минимальной упаковки', 'Ширина минимальной упаковки', 'Высота минимальной упаковки',
                   'Срок годности ', 'Дата документа сертификации ( конец срока действия документа)', 'Код ОКПД2 ',
                   'Фасовка товара (кол-во в возможных упаковках)']

wi_name_mask = ['name', 'brand', 'series', ',', 'tx', ',', 'art']
wi_non_accepted_symbols = '*{}[]*/|\!@#$%^&'
dm_non_accepted_symbols = ',""(){}[]*/|\!@#$%^&'
kp_non_accepted_symbols = '*{}[]*/|\!@#$%^&'

full_wi_mask = ['Бренд*', 'Наименование*', 'Артикул*', 'Штрихкод*', 'Код ТН ВЭД', 'Закупочная цена, руб*',
                    'РРЦ, руб', 'МОЦ, руб', 'Единица измерения*', 'Шт. в упак. *', 'Ширина, мм*', 'Длина, мм*',
                    'Высота, мм*', 'Вес, кг*', 'Описание ошибки', 'ОШИБКИ']

wi_mask_names_dict = {'Наименование товара': 'Наименование*', 'Артикул': 'Артикул*',
                          'Размер товара Длина, см': 'Длина, мм*', 'Размер товара Ширина, см': 'Ширина, мм*',
                          'Размер товара Высота, см': 'Высота, мм*', 'Вес минимальной упаковки': 'Вес, кг*',
                          'Бренд': 'Бренд*', 'Штрих-код минимальной упаковки': 'Штрихкод*', 'Код ТН ВЭД': 'Код ТН ВЭД',
                      'Штрих-код единицы товара': 'Штрихкод*', 'Вес единицы товара': 'Вес, кг*'}

try:
    dm_data_df = pd.read_excel('/content/drive/MyDrive/tgBotData/parsers/excel_parser/market_samples/CONSUMER_GOODS_template.xlsx', index_col=False, sheet_name='Справочник', keep_default_na=False)
except FileNotFoundError:
    try:
        os.mkdir('market_samples')
    except FileExistsError:
        pass
    print('Программа не может найти шаблон ДМ "CONSUMER_GOODS_template.xlsx" в папке "market_samples"')
    print('Загрузите шаблон с сайта и переместите в папку "market_samples"')
    input('Выполнение программы остановлено, для закрытия нажмите Enter')
    sys.exit()


dm_country_data_list = list(dm_data_df['country'])
dm_country_data = dm_country_data_list[1:dm_country_data_list.index('')]
dm_brand_data_list = list(dm_data_df['brand'])[1:]
dm_brand_dict = {'ArtSpace': 'ArtSpace', 'Berlingo': 'Berlingo', 'BG': 'BG', 'Crown': 'Crown', 'Eleven': 'Eleven',
                  'Greenwich Line': 'Greenwich Line', 'JOVI': 'Jovi', 'Koh-I-Noor': 'Koh-i-noor', 'Luxor': 'Luxor',
                  'MESHU': 'Meshu', 'MunHwa': 'Munhwa', 'Origami': 'ORIGAMI', 'OfficeSpace': 'OfficeSpace',
                  'Schneider': 'Schneider', 'Гамма': 'Гамма', 'Мульти-Пульти': 'Мульти-Пульти',
                  'СТАММ': 'СТАММ', 'ТРИ СОВЫ': 'ТРИ СОВЫ'}


full_dm_mask = ['ARTICLE_NUMBER', 'TYPE_PRODUCT', 'BRAND', 'MODEL', 'PLEDGE_PRICE', 'BASE_PRICE', 'SALE_PERCENT',
                'VAT_VALUE', 'BAR_CODE', 'CATEGORY', 'HERO', 'COUNTRY', 'GENDER_CATEGORY', 'ZM_AGE', 'ZAGE_TO',
                'WEIGHT', 'LENGTH', 'WIDTH', 'HEIGHT', 'DESCRIPTION', 'CERTIFICATE', 'TNVED', 'SAP_CODE',
                'IMAGE_LINKS', 'MIRROR', 'ARCHIVE_STATUS', 'ID', 'Описание ошибки', 'ОШИБКИ']


dm_mask_first_line = ['Артикул', 'Тип товара', 'Бренд', 'Модель товара', 'Залоговая стоимость*', 'Цена,  руб.*',
                'Процент скидки, %', 'НДС, %', 'Штрих-код', 'Категория', 'Герой', 'Страна-производитель', 'Пол',
                'Возраст от', 'Возраст до', 'Вес в\xa0упаковке, г', 'Длина в\xa0упаковке, см',
                'Ширина в\xa0упаковке, см', 'Высота в\xa0упаковке, см', 'Описание*', 'Сертификат',
                'Код ТН ВЭД*', 'Код ДМ', 'Ссылки на фото', 'Схема FBS', 'Архив', 'Идентификатор товара',
                      'Описание ошибки', 'ОШИБКИ']

dm_mask_names_dict = {'Наименование товара': ('TYPE_PRODUCT', 'MODEL', 'BRAND'), 'Артикул': 'ARTICLE_NUMBER', 'Ставка НДС': 'VAT_VALUE',
                      'Штрих-код минимальной упаковки': 'BAR_CODE', 'Штрих-код единицы товара': 'BAR_CODE',
                      'Страна происхождения': 'COUNTRY', 'Вес единицы товара': 'WEIGHT', 'Вес минимальной упаковки': 'WEIGHT',
                      'Размер товара Длина, см': 'LENGTH', 'Размер товара Ширина, см': 'WIDTH', 'Размер товара Высота, см': 'HEIGHT',
                      'Потребительские свойства': 'DESCRIPTION', ' Номер документа сертификации': 'CERTIFICATE',
                      'Код ТН ВЭД': 'TNVED'}


try:
    kp = pd.read_excel('/content/drive/MyDrive/tgBotData/parsers/excel_parser/market_samples/Копия вв КЭ.xlsx', index_col=False, sheet_name='Карточка товара', keep_default_na=False)
except FileNotFoundError:
    try:
        os.mkdir('market_samples')
    except FileExistsError:
        pass
    print('Программа не может найти шаблон КЭ "Копия вв КЭ.xlsx" в папке "market_samples"')
    print('Загрузите шаблон с сайта и переместите в папку "market_samples"')
    input('Выполнение программы остановлено, для закрытия нажмите Enter')
    sys.exit()


kp_mask_names_dict = {'Артикул': 'Артикул поставщика (до 7 букв)', 'Наименование товара': ('Название товара', 'Модель',
                                                                                           'Краткое описание'),
                      'Бренд': 'Бренд', 'Страна происхождения': 'Страна производства', 'Ставка НДС': 'Исходящий НДС%',
                      'Потребительские свойства': 'Описание товара', 'Штрих-код единицы товара': 'Штрихкод производителя',
                      'Штрих-код минимальной упаковки': 'Штрихкод производителя', 'Вес единицы товара': 'Вес (г)',
                      'Вес минимальной упаковки': 'Вес (г), брутто упаковка', 'Размер товара Длина, см': 'Длина (мм)',
                      'Длина минимальной упаковки': 'длина уп (м)', 'Размер товара Ширина, см': 'Ширина (мм)',
                      'Ширина минимальной упаковки': 'глубина уп (м)', 'Размер товара Высота, см': 'Высота (мм)',
                      'Высота минимальной упаковки': 'высота уп (м)', 'Срок годности ': 'Срок годности, дн',
                      ' Номер документа сертификации': 'Сертификат (номер)',
                      'Дата документа сертификации ( конец срока действия документа)': 'Сертификат (дата окончания)',
                      'Код ОКПД2 ': 'ОКПД', 'Кратность': 'Кратность (шт)'}


full_kp_mask = ['Артикул поставщика (до 7 букв)', 'Название товара', 'ГруппировкаSKU(до7букв)', 'Селектор категории',
                'id категории', 'Бренд', 'Модель', 'Страна производства', 'Исходящий НДС%',
                'Коллекция (мужская, женская, детская, унисекс)\n\n*для категории "одежда"', 'Описание товара',
                'Краткое описание', 'Свойства товара', 'Состав', 'Инструкция по уходу и эксплуатации',
                'Размерная сетка', 'Ссылки на фото ( 28 часов)', 'Штрихкод производителя', 'Цвет', 'Размер',
                'Характеристики дополнительные (до 3 характеристик)', 'Вес (г)', 'Ширина (мм)', 'Высота (мм)',
                'Длина (мм)', 'Себестоимость(руб)', 'Продажная цена (руб)', 'Производитель', 'Срок годности, дн',
                'ОСГ,%', 'Темпера-ный режим хранения', 'Единица измерения', 'Вес (г), брутто упаковка',
                'Вложимость в упаковку', 'длина транспортной упаковки(м)', 'глубина транспортной упаковки(м)',
                'высота транспортной упаковки(м)', 'длина уп (м)', 'глубина уп (м)', 'высота уп (м)',
                'кол-во коробок в слое на паллете', 'кол-во слоев на паллете', 'Кратность (шт)', 'ОКПД',
                'Сертификат (номер)', 'Сертификат (дата окончания)', 'Входящий НДС , %', 'Описание ошибки', 'ОШИБКИ']

