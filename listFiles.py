import io
import os
from googledrive import Create_Service
import pandas as pd
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import glob

CLIENT_SECRET_FILE = 'client_secrets.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']
import pandas as pd


# АЛМА АС FISH_Приложение №1 Спецификация Дулар,/Eurasia SCL Специфкация AIRBA FRESH.xlsx  ./merged_all/Мон При ТОО_ Приложение №1 Спецификация.xlsx
# /Мусихин_ Приложение №1 Спецификация ИМПОРТ.xlsx, Домашнее с любовью ТОО_ Ассортимент - Приложение №1.xlsx
# l/Aitas meat distribution ТОО_заморозка_Спецификция Ввод.xlsx

# /Альп Гулi- Приложение №1 Спецификация Флёр Альпин Бибиколь.xlsx
# PN KAZ_ Приложение №1 Спецификация.xlsx
# Рад дистрибьюция _ Приложение №1 Спецификация.xlsx



def search_by_barcode(df, df_utkonos):
    data = {}
    barcode = None


def excell_formatter(filename, row):
    data = {
                "file name": filename,
                "Поставщик": row.get('Поставщик'),
                "Штрих-код упаковки": row.get('Штрих-код упаковки') if row.get('Штрих-код упаковки') is not None else row.get('Штрих код'),
                "Штрих-код единицы": row.get('Штрих-код единицы') if row.get('Штрих-код единицы') is not None else row.get('штрих-код единицы'),
                "Наименование товара": row.get('Наименование товара ') if row.get('Наименование товара ') is not None else row.get('наименование товара'),
                "Наименование товара для маркетинга": row.get('Наименование товара для маркетинга') if row.get('Наименование товара для маркетинга') is not None else row.get('Наименование товара для маркетинга '),
                "Брэнд": row.get('Брэнд ') if row.get('Брэнд ') is not None else row.get('брэнд '),
                "Страна происхождение": row.get('Страна происхождение') if row.get('Страна происхождение') is not None else row.get('страна происхождение'),
                "Ед.изм.  шт/кг": row.get('Ед.изм.  шт/кг') if row.get('Ед.изм.  шт/кг') is not None else row.get('ед.изм.  шт/кг'),
                "Емкость, объем, вес": row.get('Емкость, объем, вес') if row.get(
                    'Емкость, объем, вес') is not None else row.get('емкость, объем, вес'),
                "Код ТН ВЭД": row.get('Код ТН ВЭД') if row.get('Код ТН ВЭД') is not None else row.get('код ТН ВЭД'),
                "Габариты штучного товара, см. Ширина": row.get('Габариты штучного товара, см. Ширина', None) if row.get(
                    'Габариты штучного товара, см. Ширина', None) is not None else row.get(
                    'Габариты штучного товара, см. ширина', None),
                "Габариты штучного товара, см. Высота": row.get('Габариты штучного товара, см. Высота') if row.get(
                    'Габариты штучного товара, см. Высота', None) is not None else row.get(
                    'Габариты штучного товара, см. высота', None),
                "Габариты штучного товара, см. Глубина": row.get('Габариты штучного товара, см. Глубина') if row.get(
                    'Габариты штучного товара, см. Глубина', None) is not None else row.get(
                    'Габариты штучного товара, см. глубина', None),
                "Минимальный заказ": row.get('Минимальный заказ') if row.get('Минимальный заказ') is not None else row.get(
                    'минимальный заказ'),
                "Описание продукции": row.get('Описание продукции') if row.get(
                    'Описание продукции') is not None else row.get('описание продукции'),
                "Состав": row['Состав'] if row.get('Состав') is not None else row.get('cостав'),
                "Белки ": row.get('Белки') if row.get('Белки') is not None else row.get('белки'),
                "Жиры": row.get('Жиры') if row.get('Жиры') is not None else row.get('жиры'),
                "Углеводы": row.get('Углеводы') if row.get('Углеводы') is not None else row.get('углеводы'),
                "Энергетическая ценность(КБЖУ)": row.get('Энергетическая ценность(КБЖУ)') if row.get(
                    'Энергетическая ценность(КБЖУ)') is not None else row.get('Энергетическая ценность'),
                "utkonos_id": "",
                "utkonos_url": ""
            }
    return data


def upload_file(file_path):
    # if not os.path.exists(file_path):
    #     print(f'{file_path} not found')
    #     return
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    try:
        folder_id = '1nUDd5MYcChlknoVRpz6F_9fejKlK79xK'
        media = MediaFileUpload(file_path,
                                mimetype='application/vnd.ms-excel')
        file_metadata = {
            'name': "Общая спецификация All with utk_id",
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [folder_id]
        }
        response = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        print(response)
    except Exception as e:
        print(e, 'this is error')
        return


def download_files():
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    # folder_id = '1UfWitOHJlODCU1hAhlckQvknKIway9l_'
    # folder_id = '14qaeozWWQ9F80LWszgE0ugRZourRbE-r'

    # folder_id = '1X_YfSbnIJy_YmvgHM1aGwNMmAqV61W_D'
    folder_id = '1pleFRO1OHVUCInnRHnPv0XBGt44KcKjM'

    query = f"parents = '{folder_id}'"
    response = service.files().list(q=query).execute()
    files = response.get('files')
    nextPageToken = response.get('nextPageToken')
    while nextPageToken is not None:
        response = service.files().list().execute()
        files.extend(response.get('files'))
        nextPageToken = response.get('nextPageToken')

    for file in files:
        request = service.files().get_media(fileId=file['id'], )
        print(request, ' this is request')
        print(file, ' this is file')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fd=fh, request=request)
        print(downloader, 'this is downloader')
        done = False

        while not done:

            status, done = downloader.next_chunk()
            print('Download progress {0}'.format(status.progress() * 100))
        fh.seek(0)

        with open(os.path.join('./files1', file['name']), 'wb') as f:
            f.write(fh.read())
            f.close()


def recreate_excell():
    path = os.path.join('./merged_all')
    filenames = glob.glob(path + "/*.xlsx")
    utkonos_path = os.path.join('./utkonosfile')
    utkonos_file = glob.glob(utkonos_path + "/*.xlsx")
    dfs = []
    df = {}
    df_utkonos = pd.read_excel(utkonos_file[0])
    count = 1
    for file in filenames:
        list_data = []
        count += 1
        filename = file.split('/')[2]
        df2 = pd.read_excel(file)
        # df2 = df.dropna(axis=0)
        # df2 = df.dropna().reset_index(drop=True)
        print(filename)
        for index, row in df2.iterrows():

            data = excell_formatter(filename, row)
            if row.get('Штрих-код единицы') is not None:
                barcode = row.get('Штрих-код единицы')
            elif row.get('Штрих-код упаковки') is not None:
                barcode = row.get('Штрих-код упаковки')
            else:
                barcode = row.get('Штрих-код')

            if barcode is not None:
                for index, row in df_utkonos.iterrows():
                    if barcode == row['Barcode']:
                        id = ''
                        if row.get('id') is not None:
                            id = row.get('id')

                        data['utkonos_id'] = id
                        data['utkonos_url'] = f"https://www.utkonos.ru/item/{id}"
            list_data.append(data)
        df = pd.DataFrame(list_data)
        dfs.append(df)
    dataframes = pd.concat(dfs)

    writer = pd.ExcelWriter("Merged_all_files_1.xlsx", engine='xlsxwriter')
    dataframes.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


def upload_run():

    path = str(os.path.join('.'))
    filenames = glob.glob(path + "/*.xlsx")
    print(filenames)
    upload_file(filenames[0])


if __name__ == '__main__':
    recreate_excell()
    # download_files()
    upload_run()
