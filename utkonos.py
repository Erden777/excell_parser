import io
import json
import os
from googledrive import Create_Service
import pandas as pd
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import glob


def main():
    utkonos_path = os.path.join('./utkonosfile')
    utkonos_file = glob.glob(utkonos_path + "/*.xlsx")
    df_utkonos = pd.read_excel(utkonos_file[0])
    data = {}
    for index, row in df_utkonos.iterrows():
        data[row['Barcode']] = row['id']
        with open("utk.json", "w") as outfile:
            json.dump(data, outfile)


if __name__ == '__main__':
    main()