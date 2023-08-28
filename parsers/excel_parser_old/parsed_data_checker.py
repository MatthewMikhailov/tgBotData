from global_data import *


def main():
    global verified_data_keys, verified_data_dicts, verified_data
    try:
        verified_df = pd.read_excel('data/verified_data.xlsx', index_col=None, sheet_name='Лист1', keep_default_na=False)
        verified_data_keys = list(verified_df['Код товара РЕЛЬЕФ'])
        verified_data_dicts = []
        for i in range(verified_df.shape[0]):
            curr_dict = {}
            for j in verified_data_mask[1:-1]:
                curr_dict[verified_data_mask_dict[j]] = verified_df[j][i]
            verified_data_dicts.append(curr_dict)
        verified_data = dict(zip(verified_data_keys, verified_data_dicts))
    except FileNotFoundError:
        print('--------------------------------------------------')
        print('Программа не может начать работу так как отсутсвует база данных("data/verified_data.xlsx")')
        print('Если вы запускаете программу впервые это нормально!')
        c_v_d_f = input('Создать новую базу данных? "Y"/"N" :')
        while not c_v_d_f.lower() in ['y', 'n']:
            print('--------------------------------------------------')
            c_v_d_f = input('Ошибка ввода, введите "Y"/"N" :')
        if c_v_d_f.lower() == 'n':
            print('--------------------------------------------------')
            input('Выполнение программы остановлено, для закрытия нажмите Enter')
            return ''
        elif c_v_d_f.lower() == 'y':
            try:
                os.mkdir('data')
            except FileExistsError:
                pass
            create_verified_data_file()
            print('--------------------------------------------------')
            input('Файл успешно создан, для продолжения работы программы нажмите Enter')
            main()


def create_verified_data_file():
    generator = {}
    for i in verified_data_mask[:len(verified_data_mask) - 1]:
        generator[i] = ['', '']
    new_data_to_verify_df = pd.DataFrame(generator)
    new_data_to_verify_df.to_excel('data/verified_data.xlsx', index=False, sheet_name='Лист1')


def check_verified_data(rf_code):
    return rf_code in verified_data_keys


def get_verified_data(rf_code):
    return verified_data[rf_code]


def generate_df_to_verify(len_data):
    generator = {}
    for v_data_mask in verified_data_mask:
        generator[v_data_mask] = ['' for gen in range(len_data)]
    return pd.DataFrame(generator)


def fill_df_to_verify(data):
    df_to_verify = generate_df_to_verify(len(data))
    for line in range(len(data)):
        for column in verified_data_mask:
            df_to_verify[column][line] = data[line][verified_data_mask_dict[column]]
    df_to_verify.to_excel('data/data_to_verification.xlsx', index=False, sheet_name='Лист1')


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
