from global_data import *
from parsed_data_checker import fill_df_to_verify, check_verified_data, get_verified_data
from get_way_to_file import send_way_to_file


def main():
    pd.options.mode.chained_assignment = None
    if check_data_to_verification_file():
        print('--------------------------------------------------')
        print('Программа не может начать работу так как в файле "data/data_to_vereficalion.xlsx" содержаться непроверенные данные')
        print('Если вы уже проверили эти данные запустите программу "Подтвердить проверку данных"')
        input('Выполнение программы остановлено, для закрытия нажмите Enter')
        return ''
    load_data_file()
    load_global_data()
    load_site_name()
    check_list_keys()
    if not multiplicity_flag:
        if no_multiplicity_flag_raise():
            return ''
    generate_df()
    print('--------------------------------------------------')
    print('Парсим ваши данные...')
    start_parsing()
    check_errors()
    if new_data_flag:
        new_data_flag_raise()
        return ''
    while True:
        try:
            curr_df.to_excel('Output/Parsed_data.xlsx', index=False)
            break
        except PermissionError:
            print('--------------------------------------------------')
            print('Нет возможности записать файл, закройте файл с названием Parsed_data.xlsx')
            input('Для повторной повытки нажмите Enter')
        except OSError:
            try:
                os.mkdir('Output')
            except FileExistsError:
                pass
    print('--------------------------------------------------')
    print('Создан файл Output/Parsed_data.xlsx')
    print('Подстраеваем шаблон магазина...')
    if site == 'ВИ':
        pars_to_WI()
    elif site == 'ДМ':
        pars_to_DM()
    elif site == 'КЭ':
        pars_to_KP()
    check_errors()
    while True:
        try:
            out_df.to_excel('Output/Market_mask.xlsx', index=False)
            break
        except PermissionError:
            print('--------------------------------------------------')
            print('Нет возможности записать файл, закройте файл с названием Market_mask.xlsx')
            input('Для повторной повытки нажмите Enter')
        except OSError:
            try:
                os.mkdir('Output')
            except FileExistsError:
                pass
    print('--------------------------------------------------')
    print('Создан файл Output/Market_mask.xlsx')
    input('Для закрытия нажмите Enter')


def no_multiplicity_flag_raise():
    print('--------------------------------------------------')
    print('Не найден столбец "Кратность" в УАМ')
    print('Программа не сможет заполнить поля "Штрихкод", "Вес", "Длина", "Ширина", "Высота", если отсутствует столбец "Кратность"')
    no_multiplicity = input('Хотите продолжить выполнение без заполнения этих полей? "Y"/"N" :')
    while not no_multiplicity.lower() in ["y", "n"]:
        print('--------------------------------------------------')
        no_multiplicity = input('Ошибка ввода, введите "Y"/"N" :')
    if no_multiplicity.lower() == 'n':
        print('--------------------------------------------------')
        input('Выполнение программы остановлено, для закрытия нажмите Enter')
        return True
    else:
        return False


def new_data_flag_raise():
    send_data_to_verification()
    print('--------------------------------------------------')
    print('Были обработаны новые данные, они сохранены в файл "data/data_to_verification.xlsx"')
    print('Необходима полная ручная проверка этих данных для исключения ошибок в последствии!!!')
    print('После проверки/редактирования файла сохраните его и запустите программу "Подтвердить проверку данных"')
    print('--------------------------------------------------')
    input('Выполнение программы остановлено, для закрытия нажмите Enter')


def check_errors():
    global wi_name_error_flag, wi_other_error_flag, dm_name_error_flag, dm_other_error_flag, kp_name_error_flag, kp_other_error_flag
    if wi_name_error_flag or dm_name_error_flag or kp_name_error_flag:
        dm_name_error_flag = False
        wi_name_error_flag = False
        print('--------------------------------------------------')
        print('Oшибка превидения данных к шаблону сайта, возможна неточность в Базе данных')
        print('Проверьте стобец "ОШИБКИ" в выходном файле для получения конкретной информации')
    if wi_other_error_flag or dm_other_error_flag or kp_other_error_flag:
        dm_other_error_flag = False
        wi_other_error_flag = False
        print('--------------------------------------------------')
        print('В УАМ не хватает данных для заполнения шаблона сайта')
        print('Проверьте стобец "ОШИБКИ" в выходном файле для получения конкретной информации')


def check_list_keys():
    global multiplicity_flag
    name = 'Кратность'
    if name.lower() in df_keys or name in df_keys:
        multiplicity_flag = True


def check_data_to_verification_file():
    try:
        check_v_file = pd.read_excel('data/data_to_verification.xlsx', index_col=None, sheet_name='Лист1')
        if check_v_file.size != 0:
            return True
    except FileNotFoundError:
        return False


def start_parsing():
    for df_line in range(len_df):
        pars_line(df_line)


def pars_line(curr_index):
    global new_data_flag
    rf_code = get_unique_rf_code(curr_index)
    if not check_verified_data(rf_code):
        new_data_flag = True
        fill_name_main(curr_index)
    else:
        verified_data = get_verified_data(rf_code)
        curr_df[full_input_mask[0]][curr_index] = verified_data
    main_parser(curr_index)


def send_data_to_verification():
    fill_df_to_verify(data_to_verification)


def get_brand(curr_index):
    return get_unique_rf_code


def get_unique_rf_code(curr_index):
    curr_df[full_input_mask[11]][curr_index] = df[full_input_mask[11]][curr_index]
    return df[full_input_mask[11]][curr_index]


def fill_art(curr_index):
    curr_df[full_input_mask[1]][curr_index] = df[full_input_mask[1]][curr_index]


def fill_solo_len(curr_index):
    curr_df[full_input_mask[2]][curr_index] = df[full_input_mask[2]][curr_index]


def fill_len(curr_index):
    curr_df[full_input_mask[18]][curr_index] = df[full_input_mask[18]][curr_index]


def fill_solo_width(curr_index):
    curr_df[full_input_mask[3]][curr_index] = df[full_input_mask[3]][curr_index]


def fill_width(curr_index):
    curr_df[full_input_mask[19]][curr_index] = df[full_input_mask[20]][curr_index]


def fill_solo_height(curr_index):
    curr_df[full_input_mask[4]][curr_index] = df[full_input_mask[4]][curr_index]


def fill_height(curr_index):
    curr_df[full_input_mask[20]][curr_index] = df[full_input_mask[20]][curr_index]


def fill_weight(curr_index):
    curr_df[full_input_mask[5]][curr_index] = df[full_input_mask[5]][curr_index]


def fill_solo_weight(curr_index):
    curr_df[full_input_mask[14]][curr_index] = df[full_input_mask[14]][curr_index]


def fill_barcode(curr_index):
    curr_df[full_input_mask[7]][curr_index] = df[full_input_mask[7]][curr_index]


def fill_solo_barcode(curr_index):
    curr_df[full_input_mask[9]][curr_index] = df[full_input_mask[9]][curr_index]


def fill_TH_code(curr_index):
    curr_df[full_input_mask[8]][curr_index] = df[full_input_mask[8]][curr_index]


def fill_trade_mark(curr_index):
    curr_df[full_input_mask[10]][curr_index] = df[full_input_mask[10]][curr_index]


def fill_brand(curr_index):
    try:
        curr_df[full_input_mask[6]][curr_index] = curr_df[full_input_mask[0]][curr_index]['brand']
    except TypeError:
        curr_df[full_input_mask[6]][curr_index] = ''


def fill_multiplicity(curr_index):
    curr_df[full_input_mask[12]][curr_index] = int(df[full_input_mask[24]][curr_index].split()[0])


def fill_country_name(curr_index):
    curr_df[full_input_mask[13]][curr_index] = df[full_input_mask[13]][curr_index]


def fill_description(curr_index):
    curr_df[full_input_mask[15]][curr_index] = df[full_input_mask[15]][curr_index]


def fill_certificate(curr_index):
    curr_df[full_input_mask[16]][curr_index] = df[full_input_mask[16]][curr_index]


def fill_VAL(curr_index):
    curr_df[full_input_mask[17]][curr_index] = df[full_input_mask[17]][curr_index]


def fill_certificate_date(curr_index):
    curr_df[full_input_mask[22]][curr_index] = df[full_input_mask[22]][curr_index]


def fill_date(curr_index):
    curr_df[full_input_mask[21]][curr_index] = df[full_input_mask[21]][curr_index]


def fill_name_main(curr_index):
    name_mask = ['name', 'brand', 'series', 'tx']
    possible_series_tx_error = False
    possible_name_error = False
    possible_error = False
    try:
        data_out = {'name': '', 'brand': brand_data[get_brand(curr_index)], ',': ',', 'series': '', 'tx': ''}
    except KeyError:
        data_out = {'name': '', 'brand': '', ',': ',', 'series': '', 'tx': ''}
        data_out['error'] = f'Торговая марка с таким названием еще не добавлена в программу ' \
                            f'"{df[full_input_mask[10]][curr_index]}"'
        data_out['code'] = get_unique_rf_code(curr_index)
        data_out['art'] = str(df[full_input_mask[1]][curr_index])
        data_to_verification.append(data_out)
        return ''
    data = df[full_input_mask[0]][curr_index].replace(',', '').split()
    try:
        brand_index = data.index(data_out['brand'])
    except ValueError:
        if data_out['brand'] == '':
            curr_df[full_input_mask[0]][curr_index] = 'Ошибка'
            possible_name_error = True
            return ''
        # ('ArtSpace', 'OfficeSpace')
        elif not isinstance(data_out['brand'], str):
            ff = False
            for x in range(len(data_out['brand'])):
                for y in range(len(data)):
                    if similarity(data_out['brand'][x], data[y]) > 0.8:
                        data_out['brand'] = data_out['brand'][x]
                        brand_index = y
                        ff = True
                        break
                if ff:
                    break
        #elif ' ' in data_out['brand'] and len(data_out['brand'].split()) >= 2:
        #    for j in range(len(data)):
        #        if similarity(' '.join(data[j:j + len(data_out['brand']) + 1]), data_out['brand']) > 0.85:
        #            brand_index = j
        # Опечатки
        else:
            possible_name_error = True
            for j in range(len(data)):
                if similarity(data[j], data_out['brand']) > 0.75 or \
                        similarity(data[j], data_out['brand'].capitalize()) > 0.75 or\
                        similarity(data[j], data_out['brand'].lower()) > 0.75 or \
                        similarity(data[j], data_out['brand'].upper()) > 0.75:
                    brand_index = j
                    break
    try:
        data_out[name_mask[0]] = ' '.join(data[:brand_index])
    except UnboundLocalError:
        data_out['art'] = ' '
        data_out['code'] = get_unique_rf_code(curr_index)
        data_out['error'] = 'Ошибка бренда'
        data_to_verification.append(data_out)
        return None
    if data_out['brand'] == 'Greenwich Line':
        data = data[brand_index + 2:]
    else:
        data = data[brand_index + 1:]
    f = True
    for i in range(len(data)):
        try:
            # series 1 word
            if data[i][0] == '"' and (data[i][-1] == '"' or data[i][-2] == '"'):
                if i == 0:
                    data_out[name_mask[2]] = data[i]
                else:
                    data_out[name_mask[2]] = ' '.join(data[:i + 1])
                data_out[name_mask[3]] = ' '.join(data[i + 1:])
                f = False
                break
            # series 1+ words
            elif data[i][-1] == '"' or data[i][-2] == '"':
                data_out[name_mask[2]] = ' '.join(data[:i + 1])
                data_out[name_mask[3]] = ' '.join(data[i + 1:])
                f = False
                break
        except IndexError:
            continue
    if f:
        data_out[name_mask[2]] = ''
        data_out[name_mask[3]] = ' '.join(data)
    if len(data_out['series']) > 20 and len(data_out['tx']) <= 1:
        curr_data = data_out['series'].split()
        start_series_index = -1
        end_series_index = -1
        for i in range(len(curr_data)):
            try:
                if (curr_data[i][0] == '"' or curr_data[i][1] == '"') and start_series_index == -1:
                    start_series_index = i
                elif (curr_data[i][-1] == '"' or curr_data[i][-2] == '"') and end_series_index == -1:
                    end_series_index = i
                elif start_series_index != -1 and end_series_index != -1:
                    break
            except IndexError:
                pass
        data_out['series'] = ' '.join(curr_data[start_series_index:end_series_index + 1])
        data_out['tx'] = ' '.join(curr_data[:start_series_index] + curr_data[end_series_index + 1:])
        possible_series_tx_error = True
    for data_out_key in data_out.keys():
        if len(data_out[data_out_key]) < 1 and data_out_key != 'series':
            possible_error = True
    data_out['art'] = str(df[full_input_mask[1]][curr_index])
    curr_df[full_input_mask[0]][curr_index] = data_out
    data_out['code'] = get_unique_rf_code(curr_index)
    if possible_error:
        data_out['error'] = 'Возможная ошибка'
    elif possible_name_error:
        data_out['error'] = 'Возможная ошибка названия'
    elif possible_series_tx_error:
        data_out['error'] = 'Возможная ошибка серия/тх'
    else:
        data_out['error'] = 'Лучше проверить!'
    data_to_verification.append(data_out)


def fill_okpd2(curr_index):
    curr_df[full_input_mask[23]][curr_index] = df[full_input_mask[23]][curr_index]


def fill_possible_multiplicity(curr_index):
    curr_df[full_input_mask[24]][curr_index] = df[full_input_mask[24]][curr_index]


def main_parser(curr_index):
    fill_art(curr_index)
    fill_solo_len(curr_index)
    fill_len(curr_index)
    fill_solo_width(curr_index)
    fill_width(curr_index)
    fill_solo_height(curr_index)
    fill_height(curr_index)
    fill_weight(curr_index)
    fill_solo_weight(curr_index)
    fill_barcode(curr_index)
    fill_TH_code(curr_index)
    fill_solo_barcode(curr_index)
    fill_trade_mark(curr_index)
    fill_brand(curr_index)
    get_unique_rf_code(curr_index)
    fill_possible_multiplicity(curr_index)
    if multiplicity_flag:
        fill_multiplicity(curr_index)
    fill_country_name(curr_index)
    fill_description(curr_index)
    fill_certificate(curr_index)
    fill_VAL(curr_index)
    fill_date(curr_index)
    fill_certificate_date(curr_index)
    fill_okpd2(curr_index)




def pars_to_WI():
    global wi_other_error_flag, wi_name_error_flag
    WI_name_mask = ['name', 'brand', 'series', ',', 'tx', ',', 'art']
    WI_non_accepted_symbols = '*{}[]*/|\!@#$%^&'
    for pars_WI in range(len_df):
        # Название
        try:
            name_WI = curr_df[full_input_mask[0]][pars_WI]
            name_out_WI = []
            for key_WI in name_WI.keys():
                if not name_WI[key_WI].isalnum():
                    line_WI = name_WI[key_WI].split()
                    for word_WI in range(len(line_WI)):
                        for symbol in WI_non_accepted_symbols:
                            if symbol in line_WI[word_WI]:
                                if symbol == line_WI[word_WI][0] or symbol == line_WI[word_WI][-1]:
                                    line_WI[word_WI] = line_WI[word_WI].replace(symbol, '')
                                else:
                                    line_WI[word_WI] = line_WI[word_WI].replace(symbol, ' ')
                    name_WI[key_WI] = ' '.join(line_WI)
            for mask_WI in WI_name_mask:
                if mask_WI == ',':
                    name_out_WI[-1] = name_out_WI[-1] + ','
                elif mask_WI == 'art':
                    if len(str(curr_df[full_input_mask[1]][pars_WI])) >= 1:
                        name_out_WI.append(str(curr_df[full_input_mask[1]][pars_WI]))
                    else:
                        raise AttributeError
                else:
                    if len(str(name_WI[mask_WI])) >= 1 or mask_WI == 'series':
                        name_out_WI.append(str(name_WI[mask_WI]))
                    elif len(str(name_WI[mask_WI])) == 0 and mask_WI == 'tx':
                        pass
                    else:
                        raise AttributeError
            out_df[wi_mask_names_dict[full_input_mask[0]]][pars_WI] = ' '.join(name_out_WI)
        except AttributeError:
            wi_name_error_flag = True
            out_df[full_wi_mask[-1]][pars_WI] = 'Требует ручной проверки'
            out_df[full_wi_mask[-2]][pars_WI] = f'Ошибка приведения данных к шаблону! Проверьте Базу данных! ' \
                                                f'Код: {curr_df["Код товара РЕЛЬЕФ"][pars_WI]}'
        # Бренд
        out_df[wi_mask_names_dict[full_input_mask[6]]][pars_WI] = curr_df[full_input_mask[6]][pars_WI]
        # Артикул
        if curr_df[full_input_mask[1]][pars_WI] == '':
            wi_other_error_flag = True
            out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[1]}"'
            out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
        out_df[wi_mask_names_dict[full_input_mask[1]]][pars_WI] = curr_df[full_input_mask[1]][pars_WI]
        # Длина
        if multiplicity_flag:
            # Длина еденицы
            if curr_df[full_input_mask[12]][pars_WI] == 1:
                if curr_df[full_input_mask[2]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[2]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                try:
                    out_df[wi_mask_names_dict[full_input_mask[2]]][pars_WI] = "%g" % (float(curr_df[full_input_mask[2]][pars_WI]) * 10)
                except ValueError:
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[2]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
            # Длина упаковки
            else:
                if curr_df[full_input_mask[18]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[18]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                try:
                    out_df[wi_mask_names_dict[full_input_mask[2]]][pars_WI] = "%g" % (float(curr_df[full_input_mask[18]][pars_WI]) * 10)
                except ValueError:
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[18]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
        # Ширина
        if multiplicity_flag:
            # Ширина еденицы
            if curr_df[full_input_mask[12]][pars_WI] == 1:
                if curr_df[full_input_mask[3]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[3]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                try:
                    out_df[wi_mask_names_dict[full_input_mask[3]]][pars_WI] = "%g" % (float(curr_df[full_input_mask[3]][pars_WI]) * 10)
                except ValueError:
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[3]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
            # Ширина упаковки
            else:
                if curr_df[full_input_mask[19]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[19]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                try:
                    out_df[wi_mask_names_dict[full_input_mask[3]]][pars_WI] = "%g" % (float(curr_df[full_input_mask[19]][pars_WI]) * 10)
                except ValueError:
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[19]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
        # Высота
        if multiplicity_flag:
            # Высота еденицы
            if curr_df[full_input_mask[12]][pars_WI] == 1:
                if curr_df[full_input_mask[4]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[4]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                try:
                    out_df[wi_mask_names_dict[full_input_mask[4]]][pars_WI] = "%g" % (float(curr_df[full_input_mask[4]][pars_WI]) * 10)
                except ValueError:
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[4]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
            # Высота упаковки
            else:
                if curr_df[full_input_mask[20]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[20]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                try:
                    out_df[wi_mask_names_dict[full_input_mask[4]]][pars_WI] = "%g" % (float(curr_df[full_input_mask[20]][pars_WI]) * 10)
                except ValueError:
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[20]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
        # Вес
        if multiplicity_flag:
            # Вес штуки
            if curr_df[full_input_mask[12]][pars_WI] == 1:
                if curr_df[full_input_mask[14]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[14]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                out_df[wi_mask_names_dict[full_input_mask[5]]][pars_WI] = curr_df[full_input_mask[14]][pars_WI]
            # Вес минимальной упаковки
            else:
                if curr_df[full_input_mask[5]][pars_WI] == '':
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[5]}"'
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                out_df[wi_mask_names_dict[full_input_mask[5]]][pars_WI] = curr_df[full_input_mask[5]][pars_WI]
        # Штрихкод
        if multiplicity_flag:
            # Штрихкод еденицы товарa
            if curr_df[full_input_mask[12]][pars_WI] == 1:
                if curr_df[full_input_mask[9]][pars_WI] == '' or not str(
                        curr_df[full_input_mask[9]][pars_WI]).isdigit():
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                    out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[9]}"'
                else:
                    out_df[wi_mask_names_dict[full_input_mask[7]]][pars_WI] = curr_df[full_input_mask[9]][pars_WI]
            # Штрихкод минимальной упаковки
            else:
                if curr_df[full_input_mask[7]][pars_WI] == '' or not str(
                        curr_df[full_input_mask[7]][pars_WI]).isdigit():
                    wi_other_error_flag = True
                    out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
                    if curr_df[full_input_mask[9]][pars_WI] == '' or str(
                            curr_df[full_input_mask[9]][pars_WI]).isdigit():
                        out_df[wi_mask_names_dict[full_input_mask[7]]][pars_WI] = curr_df[full_input_mask[9]][pars_WI]
                        out_df[full_wi_mask[-2]][
                            pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[7]}", ' \
                                          f'Значение заменено на значение {full_input_mask[9]}'
                    else:
                        out_df[full_wi_mask[-2]][
                            pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[7]}", ' \
                                          f'Значение {full_input_mask[9]} так же отсутсвует/некорректно'
                else:
                    out_df[wi_mask_names_dict[full_input_mask[7]]][pars_WI] = curr_df[full_input_mask[7]][pars_WI]
        # TH код
        if curr_df[full_input_mask[8]][pars_WI] == '':
            wi_other_error_flag = True
            out_df[full_wi_mask[-2]][pars_WI] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[8]}"'
            out_df[full_wi_mask[-1]][pars_WI] = 'Отсутствуют/некорректные данные'
        out_df[wi_mask_names_dict[full_input_mask[8]]][pars_WI] = curr_df[full_input_mask[8]][pars_WI]
        # Еденица измерения + шт в упаковке
        if multiplicity_flag:
            if curr_df[full_input_mask[12]][pars_WI] == 1:
                out_df['Единица измерения*'][pars_WI] = 'Штука'
                out_df['Шт. в упак. *'][pars_WI] = '1'
            else:
                out_df['Единица измерения*'][pars_WI] = 'Упаковка'
                out_df['Шт. в упак. *'][pars_WI] = curr_df[full_input_mask[12]][pars_WI]


def pars_to_DM():
    global dm_other_error_flag, dm_name_error_flag
    dm_non_accepted_symbols = ',""(){}[]*/|\!@#$%^&'
    for pars_DM in range(len_df):
        # Артикул
        if curr_df[full_input_mask[1]][pars_DM] == '':
            dm_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[1]}"'
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
        out_df[dm_mask_names_dict[full_input_mask[1]]][pars_DM + 1] = curr_df[full_input_mask[1]][pars_DM]
        # НДС
        if curr_df[full_input_mask[17]][pars_DM] == '':
            dm_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[17]}"'
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
        out_df[dm_mask_names_dict[full_input_mask[17]]][pars_DM + 1] = str(curr_df[full_input_mask[17]][pars_DM]) + '%'
        # Штрихкод
        if multiplicity_flag:
            # Штрихкод еденици товара
            if curr_df[full_input_mask[12]][pars_DM] == 1:
                if curr_df[full_input_mask[9]][pars_DM] == '' or not str(
                        curr_df[full_input_mask[9]][pars_DM]).isdigit():
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                    out_df[full_dm_mask[-2]][
                        pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[9]}"'
                else:
                    out_df[dm_mask_names_dict[full_input_mask[7]]][pars_DM + 1] = curr_df[full_input_mask[9]][pars_DM]
            # Штрихкод минимальной упаковки
            else:
                if curr_df[full_input_mask[7]][pars_DM] == '' or not str(
                        curr_df[full_input_mask[7]][pars_DM]).isdigit():
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                    if curr_df[full_input_mask[9]][pars_DM] == '' or str(
                            curr_df[full_input_mask[9]][pars_DM]).isdigit():
                        out_df[dm_mask_names_dict[full_input_mask[7]]][pars_DM + 1] = curr_df[full_input_mask[9]][pars_DM]
                        out_df[full_dm_mask[-2]][
                            pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[7]}", ' \
                                       f'Значение заменено на значение {full_input_mask[9]}'
                    else:
                        out_df[full_dm_mask[-2]][
                            pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[7]}", ' \
                                       f'Значение {full_input_mask[9]} так же отсутсвует/некорректно'
                else:
                    out_df[dm_mask_names_dict[full_input_mask[7]]][pars_DM + 1] = curr_df[full_input_mask[7]][pars_DM]
        # Вес
        if multiplicity_flag:
            # Вес штуки
            if curr_df[full_input_mask[12]][pars_DM] == 1:
                if curr_df[full_input_mask[14]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][
                        pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[14]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[5]]][pars_DM + 1] = "%g" % (float(curr_df[full_input_mask[14]][pars_DM]) * 1000)
            # Вес минимальной упаковки
            else:
                if curr_df[full_input_mask[5]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[5]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[5]]][pars_DM + 1] = "%g" % (float(curr_df[full_input_mask[5]][pars_DM]) * 1000)
        # Длина
        if multiplicity_flag:
            # Длина штуки
            if curr_df[full_input_mask[12]][pars_DM] == 1:
                if curr_df[full_input_mask[2]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[2]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[2]]][pars_DM + 1] = curr_df[full_input_mask[2]][pars_DM]
            # Длина упаковки
            else:
                if curr_df[full_input_mask[18]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[18]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[2]]][pars_DM + 1] = curr_df[full_input_mask[18]][pars_DM]
        # Ширина
        if multiplicity_flag:
            # Ширина штуки
            if curr_df[full_input_mask[12]][pars_DM] == 1:
                if curr_df[full_input_mask[3]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[3]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[3]]][pars_DM + 1] = curr_df[full_input_mask[3]][pars_DM]
            # Ширина упаковки
            else:
                if curr_df[full_input_mask[19]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[19]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[3]]][pars_DM + 1] = curr_df[full_input_mask[19]][pars_DM]
        # Высота
        if multiplicity_flag:
            # Высота штуки
            if curr_df[full_input_mask[12]][pars_DM] == 1:
                if curr_df[full_input_mask[4]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[4]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[4]]][pars_DM + 1] = curr_df[full_input_mask[4]][pars_DM]
            # Высота упаковки
            else:
                if curr_df[full_input_mask[20]][pars_DM] == '':
                    dm_other_error_flag = True
                    out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[20]}"'
                    out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
                out_df[dm_mask_names_dict[full_input_mask[4]]][pars_DM + 1] = curr_df[full_input_mask[20]][pars_DM]
        # Описание
        if curr_df[full_input_mask[15]][pars_DM] == '':
            dm_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[15]}"'
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
        out_df[dm_mask_names_dict[full_input_mask[15]]][pars_DM + 1] = curr_df[full_input_mask[15]][pars_DM]
        # Сертификат
        if curr_df[full_input_mask[16]][pars_DM] == '':
            dm_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[16]}"'
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
        out_df[dm_mask_names_dict[full_input_mask[16]]][pars_DM + 1] = curr_df[full_input_mask[16]][pars_DM]
        # TH ВЭД
        if curr_df[full_input_mask[8]][pars_DM] == '':
            dm_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[8]}"'
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
        out_df[dm_mask_names_dict[full_input_mask[8]]][pars_DM + 1] = curr_df[full_input_mask[8]][pars_DM]
        # Тип товара + модель
        try:
            name_dm = curr_df[full_input_mask[0]][pars_DM]
            for key_dm in name_dm.keys():
                if not name_dm[key_dm].isalnum():
                    line_dm = name_dm[key_dm].split()
                    for word_dm in range(len(line_dm)):
                        for symbol in dm_non_accepted_symbols:
                            if symbol in line_dm[word_dm]:
                                if symbol == line_dm[word_dm][0] or symbol == line_dm[word_dm][-1]:
                                    line_dm[word_dm] = line_dm[word_dm].replace(symbol, '')
                                else:
                                    line_dm[word_dm] = line_dm[word_dm].replace(symbol, ' ')
                    name_dm[key_dm] = ' '.join(line_dm)
            out_df[dm_mask_names_dict[full_input_mask[0]][0]][pars_DM + 1] = name_dm['name']
            out_df[dm_mask_names_dict[full_input_mask[0]][1]][pars_DM + 1] = name_dm['tx']
        except AttributeError:
            dm_name_error_flag = True
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Требует ручной проверки'
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'Ошибка приведения данных к шаблону! Проверьте Базу данных! ' \
                                                f'Код: {curr_df["Код товара РЕЛЬЕФ"][pars_DM]}'
        # Страна
        if curr_df[full_input_mask[13]][pars_DM] == '':
            dm_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[13]}"'
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Отсутствуют/некорректные данные'
        try:
            dm_country = dm_country_data[dm_country_data.index(curr_df[full_input_mask[13]][pars_DM].capitalize())]
            out_df[dm_mask_names_dict[full_input_mask[13]]][pars_DM + 1] = dm_country
        except ValueError:
            dm_country = ''
            try:
                if len(curr_df[full_input_mask[13]][pars_DM].split()) == 2:
                    dm_country = 'Южная Корея'
                else:
                    raise ValueError
                out_df[dm_mask_names_dict[full_input_mask[13]]][pars_DM + 1] = dm_country
            except ValueError:
                dm_other_error_flag = True
                out_df[full_dm_mask[-2]][pars_DM + 1] = f'Ошибка подставления "{full_input_mask[13]}" в шаблон сайта,' \
                                                        f' значение в УАМ: {curr_df[full_input_mask[13]][pars_DM]}'
                out_df[full_dm_mask[-1]][pars_DM + 1] = 'Данные не соответствуют шаблону'
        # Бренд
        try:
            out_df[dm_mask_names_dict[full_input_mask[0]][2]][pars_DM + 1] = \
                dm_brand_dict[curr_df[full_input_mask[6]][pars_DM]]
        except KeyError:
            dm_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_DM + 1] = f'Ошибка подставления "{full_input_mask[6]}" в шаблон сайта,' \
                                                    f' значение в УАМ: {curr_df[full_input_mask[6]][pars_DM]}'
            out_df[full_dm_mask[-1]][pars_DM + 1] = 'Данные не соответствуют шаблону'


def pars_to_KP():
    global kp_other_error_flag, kp_name_error_flag
    kp_non_accepted_symbols = ['*{}[]*/|\!@#$%^&']
    for pars_kp in range(len_df):
        # Артикул
        if curr_df[full_input_mask[1]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[1]}"'
            out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[1]]][pars_kp] = curr_df[full_input_mask[1]][pars_kp]
        # Бренд
        try:
            out_df[kp_mask_names_dict[full_input_mask[6]]][pars_kp] = \
                dm_brand_dict[curr_df[full_input_mask[6]][pars_kp]]
        except KeyError:
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'Ошибка подставления "{full_input_mask[6]}" в шаблон сайта,' \
                                                    f' значение в УАМ: {curr_df[full_input_mask[6]][pars_kp]}'
            out_df[full_kp_mask[-1]][pars_kp] = 'Данные не соответствуют шаблону'
        # Страна
        if curr_df[full_input_mask[13]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[13]}"'
            out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[13]]][pars_kp] = curr_df[full_input_mask[13]][pars_kp]
        # НДС
        if curr_df[full_input_mask[17]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[17]}"'
            out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[17]]][pars_kp] = str(curr_df[full_input_mask[17]][pars_kp]) + '%'
        # Описание товара
        if curr_df[full_input_mask[15]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[15]}"'
            out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[15]]][pars_kp] = curr_df[full_input_mask[15]][pars_kp]
        # Сертификат
        if curr_df[full_input_mask[16]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[16]}"'
            out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[16]]][pars_kp] = curr_df[full_input_mask[16]][pars_kp]
        # Сертификат дата
        if curr_df[full_input_mask[22]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[22]}"'
            out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[22]]][pars_kp] = curr_df[full_input_mask[22]][pars_kp]
        # Наименование + Модель
        try:
            name_kp = curr_df[full_input_mask[0]][pars_kp]
            for key_dm in name_kp.keys():
                if not name_kp[key_dm].isalnum():
                    line_dm = name_kp[key_dm].split()
                    for word_dm in range(len(line_dm)):
                        for symbol in kp_non_accepted_symbols:
                            if symbol in line_dm[word_dm]:
                                if symbol == line_dm[word_dm][0] or symbol == line_dm[word_dm][-1]:
                                    line_dm[word_dm] = line_dm[word_dm].replace(symbol, '')
                                else:
                                    line_dm[word_dm] = line_dm[word_dm].replace(symbol, ' ')
                    name_kp[key_dm] = ' '.join(line_dm)
            out_df[kp_mask_names_dict[full_input_mask[0]][0]][pars_kp] = name_kp['name']
            out_df[kp_mask_names_dict[full_input_mask[0]][1]][pars_kp] = name_kp['series']
            out_df[kp_mask_names_dict[full_input_mask[0]][2]][pars_kp] = name_kp['name'] + name_kp['tx']
        except AttributeError:
            kp_name_error_flag = True
            out_df[full_kp_mask[-1]][pars_kp] = 'Требует ручной проверки'
            out_df[full_kp_mask[-2]][pars_kp] = f'Ошибка приведения данных к шаблону! Проверьте Базу данных! ' \
                                                f'Код: {curr_df["Код товара РЕЛЬЕФ"][pars_kp]}'
        # Кратсноть
        if multiplicity_flag:
            if curr_df[full_input_mask[12]][pars_kp] == '':
                kp_other_error_flag = True
                out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[12]}"'
                out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
            out_df[kp_mask_names_dict[full_input_mask[12]]][pars_kp] = curr_df[full_input_mask[12]][pars_kp]
        # Еденица измерения
        if multiplicity_flag:
            if curr_df[full_input_mask[12]][pars_kp] == 1:
                out_df['Единица измерения'][pars_kp] = "Штука"
            else:
                out_df['Единица измерения'][pars_kp] = "Упаковка"
        # ОКПД
        if curr_df[full_input_mask[23]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_kp_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[23]}"'
            out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[23]]][pars_kp] = curr_df[full_input_mask[23]][pars_kp]
        # Штрихкод
        if multiplicity_flag:
            # Штрихкод еденици товара
            if curr_df[full_input_mask[12]][pars_kp] == 1:
                if curr_df[full_input_mask[9]][pars_kp] == '' or not str(
                        curr_df[full_input_mask[9]][pars_kp]).isdigit():
                    dm_other_error_flag = True
                    out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
                    out_df[full_kp_mask[-2]][
                        pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[9]}"'
                else:
                    out_df[kp_mask_names_dict[full_input_mask[7]]][pars_kp] = curr_df[full_input_mask[9]][pars_kp]
            # Штрихкод минимальной упаковки
            else:
                if curr_df[full_input_mask[7]][pars_kp] == '' or not str(
                        curr_df[full_input_mask[7]][pars_kp]).isdigit():
                    kp_other_error_flag = True
                    out_df[full_kp_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
                    if curr_df[full_input_mask[9]][pars_kp] == '' or str(
                            curr_df[full_input_mask[9]][pars_kp]).isdigit():
                        out_df[kp_mask_names_dict[full_input_mask[7]]][pars_kp] = curr_df[full_input_mask[9]][pars_kp]
                        out_df[full_kp_mask[-2]][
                            pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[7]}", ' \
                                       f'Значение заменено на значение {full_input_mask[9]}'
                    else:
                        out_df[full_kp_mask[-2]][
                            pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[7]}", ' \
                                       f'Значение {full_input_mask[9]} так же отсутсвует/некорректно'
                else:
                    out_df[kp_mask_names_dict[full_input_mask[7]]][pars_kp] = curr_df[full_input_mask[7]][pars_kp]
        # Вес
        if curr_df[full_input_mask[14]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][
                pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[14]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        else:
            out_df[kp_mask_names_dict[full_input_mask[14]]][pars_kp] = "%g" % (
                    float(curr_df[full_input_mask[14]][pars_kp]) * 1000)
        # Вес минимальной упаковки
        if curr_df[full_input_mask[5]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[5]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        else:
            out_df[kp_mask_names_dict[full_input_mask[5]]][pars_kp] = "%g" % (
                    float(curr_df[full_input_mask[5]][pars_kp]) * 1000)
        # Высота
        # Высота штуки
        if curr_df[full_input_mask[4]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[4]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[4]]][pars_kp] = curr_df[full_input_mask[4]][pars_kp]
        # Высота упаковки
        if curr_df[full_input_mask[20]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[20]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        else:
            out_df[kp_mask_names_dict[full_input_mask[20]]][pars_kp] = float(curr_df[full_input_mask[20]][pars_kp]) * 0.01
        # Длина
        # Длина штуки
        if curr_df[full_input_mask[2]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[2]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[2]]][pars_kp] = curr_df[full_input_mask[2]][pars_kp]
        # Длина упаковки
        if curr_df[full_input_mask[18]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[18]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        else:
            out_df[kp_mask_names_dict[full_input_mask[18]]][pars_kp] = float(curr_df[full_input_mask[18]][pars_kp]) * 0.01
        # Ширина
        # Ширина штуки
        if curr_df[full_input_mask[3]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[3]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        out_df[kp_mask_names_dict[full_input_mask[3]]][pars_kp] = curr_df[full_input_mask[3]][pars_kp]
        # Ширина упаковки
        if curr_df[full_input_mask[19]][pars_kp] == '':
            kp_other_error_flag = True
            out_df[full_dm_mask[-2]][pars_kp] = f'В УАМ отсутствуют/некорректные данные "{full_input_mask[19]}"'
            out_df[full_dm_mask[-1]][pars_kp] = 'Отсутствуют/некорректные данные'
        else:
            out_df[kp_mask_names_dict[full_input_mask[19]]][pars_kp] = float(curr_df[full_input_mask[19]][pars_kp]) * 0.01


def similarity(line1, line2):
    return difflib.SequenceMatcher(None, line1, line2).ratio()


def load_global_data():
    global df_keys, df_key, len_df
    global brand_data, site_data
    global full_wi_mask, wi_mask_names_dict
    global full_input_mask, kp_name_error_flag
    global wi_name_error_flag, wi_name_error_data
    global wi_other_error_flag, wi_other_error_data
    global data_to_verification, new_data_flag
    global multiplicity_flag, kp_other_error_flag
    global dm_other_error_flag, dm_name_error_flag
    multiplicity_flag = False
    new_data_flag = False
    site_data = ['ВИ', 'ДМ', 'КЭ']
    data_to_verification = []
    len_df = df.shape[0]
    df_keys = list(df.keys())
    df_key = df_keys[0]
    wi_name_error_flag, kp_name_error_flag = False, False
    wi_other_error_flag, kp_other_error_flag = False, False
    dm_other_error_flag, dm_name_error_flag = False, False
    return None


def generate_df():
    global curr_df, out_df
    curr_df_generator = {}
    out_df_generator = {}
    for curr_df_key in full_input_mask:
        if curr_df_key == 'Кратность':
            if multiplicity_flag:
                curr_df_generator[curr_df_key] = ['' for generate in range(len_df)]
        else:
            curr_df_generator[curr_df_key] = ['' for generate in range(len_df)]
    if site == 'ВИ':
        for out_df_key in full_wi_mask:
            out_df_generator[out_df_key] = ['' for generate in range(len_df)]
        out_df = pd.DataFrame(out_df_generator)
    elif site == 'ДМ':
        for out_df_key in full_dm_mask:
            out_df_generator[out_df_key] = ['' for generate in range(len_df + 1)]
        df1 = pd.DataFrame(dict(zip(full_dm_mask, dm_mask_first_line)), index=[0])
        df2 = pd.DataFrame(out_df_generator)
        out_df = pd.concat([df1, df2], ignore_index=True)
    elif site == 'КЭ':
        for out_df_key in full_kp_mask:
            out_df_generator[out_df_key] = ['' for generate in range(len_df + 1)]
        out_df = pd.DataFrame(out_df_generator)
    curr_df = pd.DataFrame(curr_df_generator)


def load_data_file():
    global df
    try:
        df = pd.read_excel(f'{send_way_to_file()}', index_col=None, sheet_name='Лист1',
                               keep_default_na=False)
    except FileNotFoundError:
        print('--------------------------------------------------')
        print('Произошла ошибка виджета открытия файла')
        print('Пожалуйста переместите файл в папку с программой и введите его название вручную')
        while True:
            try:
                print('--------------------------------------------------')
                df = pd.read_excel(f'{input("Введите название исходного файла:")}.xls', index_col=None, sheet_name='Лист1',
                                   keep_default_na=False)
                return None
            except FileNotFoundError:
                print('--------------------------------------------------')
                print("Файл не найден!")
                print(f'Текущая папка:{os.getcwd()}')
                print('Текущий список файлов в папке с программой:')
                print(os.listdir(path=""))
                continue


def load_site_name():
    global site
    while True:
        site = input('Введите название сайта:')
        #site = 'КЭ'
        if site in site_data:
            return None
        else:
            print('--------------------------------------------------')
            print('Неверно указано название сайта!')
            print('Шаблоны названий сайтов:')
            print(site_data)


if __name__ == '__main__':
    try:
        main()
    except Exception:
        try:
            os.mkdir('error_logs')
        except FileExistsError:
            pass
        log_date = '_'.join(str(datetime.datetime.today()).split('.')[0].split(' ')).replace(':', '-')
        file_name = f'error_logs/log_{log_date}.txt'
        err_file = open(file=file_name, mode='w')
        traceback.print_exc(file=err_file)
        err_file.close()
        print('--------------------------------------------------')
        print('Критическая ошибка программы!')
        print(f'Полные сведения об ошибке были сохранены в папку "error_logs" в файл "log_{log_date}.txt"')
        print('Чтобы устранить ошибку необходимо отправить файл с описанием ошибки разработчику программы')
        input("Нажмите Enter чтобы закрыть")
