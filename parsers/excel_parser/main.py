from parsers.excel_parser.global_data import *

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(name)s - %(message)s", level=logging.INFO, datefmt='%H:%M:%S'
)

logger = logging.getLogger(__name__)


def get_current_mask(site):
    if site == 'ВИ':
        return full_wi_mask, wi_mask_names_dict, wi_non_accepted_symbols
    elif site == 'ДМ':
        return full_dm_mask, dm_mask_names_dict, dm_non_accepted_symbols
    else:
        return full_kp_mask, kp_mask_names_dict, kp_non_accepted_symbols


def pars_to_site_sample(site_name):
    # print('ok')
    out_df = get_out_df(site_name)
    # print('ok')
    current_mask, current_mask_names, non_accepted_symbols = get_current_mask(site_name)
    wi_flag = True if site_name == 'ВИ' else False
    dm_flag = True if site_name == 'ДМ' else False
    kp_flag = True if site_name == 'КЭ' else False
    for line_index in range(len_df):
        # print('ok')
        name = pars_name(line_index)
        try:
            if name['error'] == 'Ошибка':
                print('raise AttributeError')
                raise AttributeError
            name_out = []
            for key in name.keys():
                if not name[key].isalnum():
                    line = name[key].split()
                    for word in range(len(line)):
                        for symbol in non_accepted_symbols:
                            if symbol in line[word]:
                                if symbol == line[word][0] or symbol == line[word][-1]:
                                    line[word] = line[word].replace(symbol, '')
                                else:
                                    line[word] = line[word].replace(symbol, ' ')
                    name[key] = ' '.join(line)

            if wi_flag:
                for mask in wi_name_mask:
                    if mask == ',':
                        name_out[-1] = name_out[-1] + ','
                    elif mask == 'art':
                        if len(str(df[full_input_mask[1]][line_index])) >= 1:
                            name_out.append(str(df[full_input_mask[1]][line_index]))
                        else:
                            print('raise AttributeError')
                            raise AttributeError
                    else:
                        if len(str(name[mask])) >= 1 or mask == 'series':
                            name_out.append(str(name[mask]))
                        elif len(str(name[mask])) == 0 and mask == 'tx':
                            pass
                        else:
                            print('raise AttributeError')
                            raise AttributeError
                out_df[wi_mask_names_dict[full_input_mask[0]]][line_index] = ' '.join(name_out)
                out_df[wi_mask_names_dict[full_input_mask[6]]][line_index] = name['brand']
            elif dm_flag:
                out_df[dm_mask_names_dict[full_input_mask[0]][0]][line_index] = name['name']
                out_df[dm_mask_names_dict[full_input_mask[0]][1]][line_index] = name['tx']
                try:
                    out_df[dm_mask_names_dict[full_input_mask[0]][2]][line_index] = \
                        dm_brand_dict[name['brand']]
                except KeyError:
                    out_df[full_dm_mask[-2]][
                        line_index] = f'Ошибка подставления "{full_input_mask[6]}" в шаблон сайта,' \
                                          f' значение в УАМ: {name["brand"]}'
                    out_df[full_dm_mask[-1]][line_index] = 'Данные не соответствуют шаблону'
            elif kp_flag:
                out_df[kp_mask_names_dict[full_input_mask[0]][0]][line_index] = name['name']
                out_df[kp_mask_names_dict[full_input_mask[0]][1]][line_index] = name['series']
                out_df[kp_mask_names_dict[full_input_mask[0]][2]][line_index] = name['name'] + name['tx']
                try:
                    out_df[kp_mask_names_dict[full_input_mask[6]]][line_index] = name['brand']
                except KeyError:
                    out_df[full_kp_mask[-2]][line_index] = \
                        f'Ошибка подставления "{full_input_mask[6]}" в шаблон сайта, значение в УАМ: {name["brand"]}'
                    out_df[full_kp_mask[-1]][line_index] = 'Данные не соответствуют шаблону'
            if name['error'] != '':
                out_df[full_wi_mask[-1]][line_index] = 'Требует ручной проверки'
                out_df[full_wi_mask[-2]][line_index] = f'{name["error"]}'
        except AttributeError as ex:
            print(ex)
            out_df[full_wi_mask[-1]][line_index] = 'Требует ручной проверки'
            out_df[full_wi_mask[-2]][line_index] = f'Ошибка приведения данных к шаблону! {name}'

        if df[full_input_mask[24]][line_index].split()[0] == '1':
            # Штрихкод
            out_df[current_mask_names[full_input_mask[7]]][line_index] = df[full_input_mask[9]][line_index]
            try:
                # Вес
                out_df[current_mask_names[full_input_mask[5]]][line_index] = \
                    "%g" % (float(df[full_input_mask[14]][line_index]) * 1000) if dm_flag \
                    else df[full_input_mask[14]][line_index]
            except ValueError:
                pass
            try:
                # Высота
                out_df[current_mask_names[full_input_mask[4]]][line_index] = \
                    df[full_input_mask[4]][line_index] if dm_flag else \
                    "%g" % (float(df[full_input_mask[4]][line_index]) * 10)
                # Ширина
                out_df[current_mask_names[full_input_mask[3]]][line_index] = \
                    df[full_input_mask[3]][line_index] if dm_flag else \
                    "%g" % (float(df[full_input_mask[3]][line_index]) * 10)
                # Длина
                out_df[current_mask_names[full_input_mask[2]]][line_index] = \
                    df[full_input_mask[2]][line_index] if dm_flag else  \
                    "%g" % (float(df[full_input_mask[2]][line_index]) * 10)
            except ValueError:
                pass
            if wi_flag:
                out_df['Единица измерения*'][line_index] = 'Штука'
                out_df['Шт. в упак. *'][line_index] = '1'
        else:
            # Штрихкод
            out_df[current_mask_names[full_input_mask[7]]][line_index] = df[full_input_mask[7]][line_index]
            try:
                # Вес
                out_df[current_mask_names[full_input_mask[5]]][line_index] = \
                    "%g" % (float(df[full_input_mask[14]][line_index]) * 1000) if dm_flag \
                    else df[full_input_mask[14]][line_index]
            except ValueError:
                pass
            try:
                # Высота
                out_df[current_mask_names[full_input_mask[4]]][line_index] = \
                    df[full_input_mask[20]][line_index] if dm_flag else \
                    "%g" % (float(df[full_input_mask[20]][line_index]) * 10)
                # Ширина
                out_df[current_mask_names[full_input_mask[3]]][line_index] = \
                    df[full_input_mask[19]][line_index] if dm_flag else \
                    "%g" % (float(df[full_input_mask[19]][line_index]) * 10)
                # Длина
                out_df[current_mask_names[full_input_mask[2]]][line_index] = \
                    df[full_input_mask[18]][line_index] if dm_flag else \
                    "%g" % (float(df[full_input_mask[18]][line_index]) * 10)
            except ValueError:
                pass
            if wi_flag:
                out_df['Единица измерения*'][line_index] = 'Упаковка'
                out_df['Шт. в упак. *'][line_index] = df[full_input_mask[24]][line_index].split()[0]
        if dm_flag:
            # НДС
            out_df[dm_mask_names_dict[full_input_mask[17]]][line_index] = str(
                df[full_input_mask[17]][line_index]) + '%'
            # Страна
            try:
                dm_country = dm_country_data[dm_country_data.index(df[full_input_mask[13]][line_index].capitalize())]
                out_df[dm_mask_names_dict[full_input_mask[13]]][line_index] = dm_country
            except ValueError:
                try:
                    if len(df[full_input_mask[13]][line_index].split()) == 2:
                        dm_country = 'Южная Корея'
                    else:
                        raise ValueError
                    out_df[dm_mask_names_dict[full_input_mask[13]]][line_index] = dm_country if dm_country else ''
                except ValueError:
                    out_df[full_dm_mask[-2]][
                        line_index] = f'Ошибка подставления "{full_input_mask[13]}" в шаблон сайта,' \
                                       f' значение в УАМ: {df[full_input_mask[13]][line_index]}'
                    out_df[full_dm_mask[-1]][line_index] = 'Данные не соответствуют шаблону'
    # Артикул
    out_df[current_mask_names[full_input_mask[1]]] = df[full_input_mask[1]]
    # TH код
    out_df[current_mask_names[full_input_mask[8]]] = df[full_input_mask[8]]
    if not wi_flag:
        # Описание
        out_df[current_mask_names[full_input_mask[15]]] = df[full_input_mask[15]]
        # Сертификат
        out_df[current_mask_names[full_input_mask[16]]] = df[full_input_mask[16]]
    return out_df


def pars_name(curr_index):
    possible_series_tx_error = False
    possible_name_error = False
    possible_error = False
    data_out = {'name': '', 'brand': '', 'tx': '', 'series': '', 'error': '',
                'code': str(df[full_input_mask[11]][curr_index]), 'art': str(df[full_input_mask[1]][curr_index])}
    try:
        data_out['brand'] = brand_data[df[full_input_mask[10]][curr_index]]
    except KeyError:
        data_out['error'] = f'Торговая марка с таким названием еще не добавлена в программу ' \
                            f'"{df[full_input_mask[10]][curr_index]}"'
        return data_out
    data = df[full_input_mask[0]][curr_index].replace(',', '').split()
    brand_index = -1
    try:
        brand_index = data.index(data_out['brand'])
    except ValueError:
        # ('ArtSpace', 'OfficeSpace')
        if not isinstance(data_out['brand'], str):
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
        # Опечатки
        else:
            possible_name_error = True
            for j in range(len(data)):
                if similarity(data[j], data_out['brand']) > 0.75 or \
                        similarity(data[j], data_out['brand'].capitalize()) > 0.75 or \
                        similarity(data[j], data_out['brand'].lower()) > 0.75 or \
                        similarity(data[j], data_out['brand'].upper()) > 0.75:
                    brand_index = j
                    break
    try:
        data_out['name'] = ' '.join(data[:brand_index])
        if brand_index == -1:
            raise UnboundLocalError
    except UnboundLocalError:
        data_out['error'] = 'Ошибка'
        return data_out
    if ' ' in data_out['brand'] and len(data_out['brand'].split()) >= 2:
        data = data[brand_index + len(data_out['brand'].split()):]
    else:
        data = data[brand_index + 1:]
    f = True
    for i in range(len(data)):
        try:
            # series 1 word
            if data[i][0] == '"' and (data[i][-1] == '"' or data[i][-2] == '"'):
                if i == 0:
                    data_out['series'] = data[i]
                else:
                    data_out['series'] = ' '.join(data[:i + 1])
                data_out['tx'] = ' '.join(data[i + 1:])
                f = False
                break
            # series 1+ words
            elif data[i][-1] == '"' or data[i][-2] == '"':
                data_out['series'] = ' '.join(data[:i + 1])
                data_out['tx'] = ' '.join(data[i + 1:])
                f = False
                break
        except IndexError:
            continue
    if f:
        data_out['tx'] = ' '.join(data)
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
    for data_out_key in range(3):
        if len(data_out[list(data_out.keys())[data_out_key]]) < 1:
            possible_error = True
    if possible_error:
        data_out['error'] = 'Возможная ошибка'
    elif possible_name_error:
        data_out['error'] = 'Возможная ошибка названия'
    elif possible_series_tx_error:
        data_out['error'] = 'Возможная ошибка серия/тх'
    return data_out


def get_out_df(site):
    out_df_generator = {}
    if site == 'ВИ':
        for out_df_key in full_wi_mask:
            out_df_generator[out_df_key] = ['' for generate in range(len_df)]
        out_df = pd.DataFrame(out_df_generator)
    elif site == 'ДМ':
        for out_df_key in full_dm_mask:
            out_df_generator[out_df_key] = ['' for generate in range(len_df)]
        out_df = pd.DataFrame(out_df_generator)
    else:
        for out_df_key in full_kp_mask:
            out_df_generator[out_df_key] = ['' for generate in range(len_df)]
        out_df = pd.DataFrame(out_df_generator)
    return out_df


def main(filename, site_name, chat_id):
    global df, len_df
    df = pd.read_excel(f'{filename}', index_col=None, sheet_name='Лист1', keep_default_na=False)
    len_df = df.shape[0]
    out = pars_to_site_sample(site_name)
    out.to_excel('/'.join(filename.split('/')[:2]) + '/parsed_' + filename.split('/')[2].split('.')[0] + '.xlsx',
                 index=False)
    # out.to_excel('Output.xlsx', index=False)
    print('ok')


def run_excel_parser(filename, site_name, chat_id):
    try:
        logger.info(f'({chat_id})Started')
        main(filename, site_name, chat_id)
    except Exception:
        try:
            os.mkdir('../../error_logs')
        except FileExistsError:
            pass
        log_date = '_'.join(str(datetime.today()).split('.')[0].split('-')[-1].split(' ')).replace(':', '-')
        file_name = f'../error_logs/excel_{log_date}.txt'
        err_file = open(file=file_name, mode='w')
        traceback.print_exc(file=err_file)
        err_file.close()


if __name__ == '__main__':
    run_excel_parser('Excel_test.xls', 'ДМ', 1275943662)
