from global_data import *


def main():
    while True:
        try:
            data_to_verify_df = pd.read_excel('data/data_to_verification.xlsx', index_col=None, sheet_name='Лист1', keep_default_na=False)
            verified_df = pd.read_excel('data/verified_data.xlsx', index_col=None, sheet_name='Лист1')
            break
        except ValueError:
            print("Невозможно выполнить действие, закройте файлы 'data_to_verification.xlsx' и 'verified_data.xlsx'")
            input('Для повторной попытки нажмите Enter')
        except PermissionError:
            print("Невозможно выполнить действие, закройте файлы 'data_to_verification.xlsx' и 'verified_data.xlsx'")
            input('Для повторной попытки нажмите Enter')
    if data_to_verify_df.size == 0:
        input("Данные уже были добавлены в Базу данных, для закрытия нажмите Enter")
        return ''
    drop_lines = []
    for line in range(1, data_to_verify_df.shape[0]):
        if data_to_verify_df.iloc[line]['Код товара РЕЛЬЕФ'] == '':
            drop_lines.append(line)
    data_to_verify_df = data_to_verify_df.drop(drop_lines).drop('Ошибки', axis=1)

    new_verified_df = pd.concat([verified_df, data_to_verify_df])
    while True:
        try:
            new_verified_df.to_excel('data/verified_data.xlsx', index=False, sheet_name='Лист1')
            break
        except ValueError:
            print("Невозможно выполнить действие, закройте файлы 'data_to_verification.xlsx' и 'verified_data.xlsx'")
            input('Для повторной попытки нажмите Enter')
        except PermissionError:
            print("Невозможно выполнить действие, закройте файлы 'data_to_verification.xlsx' и 'verified_data.xlsx'")
            input('Для повторной попытки нажмите Enter')

    generator = {}
    for i in verified_data_mask:
        generator[i] = ['', '']
    new_data_to_verify_df = pd.DataFrame(generator)
    while True:
        try:
            new_data_to_verify_df.to_excel('data/data_to_verification.xlsx', index=False, sheet_name='Лист1')
            break
        except ValueError:
            print("Невозможно выполнить действие, закройте файлы 'data_to_verification.xlsx' и 'verified_data.xlsx'")
            input('Для повторной попытки нажмите Enter')
        except PermissionError:
            print("Невозможно выполнить действие, закройте файлы 'data_to_verification.xlsx' и 'verified_data.xlsx'")
            input('Для повторной попытки нажмите Enter')
    input("Данные успешно добавлены в Базу данных, для закрытия нажмите Enter")


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
