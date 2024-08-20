import subprocess
import sys
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import pandas as pd
import os
import logging
from googleapiclient.errors import HttpError

# Ensure the correct timezone is set
os.environ['TZ'] = 'Asia/Jakarta'  # Set timezone to GMT+7
time.tzset()

# Configure logging
logging.basicConfig(level=logging.INFO)

# Install openpyxl if not already installed
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])

# JSON key directly embedded in the script
SERVICE_ACCOUNT_INFO = {
    "type": "service_account",
  "project_id": "bps3505-zonaintegritas",
  "private_key_id": "3a0671e7d928faadc10e3a2b969a7e2c39ed0906",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCZqObAP2k2252E\nUTbtH+Lxv37dmORM/zVLXq12dd9Xgo0hUcqoYQY1hKG0Tl2itDvYK3+FmsKEGq2n\nF5z7E++s2uYex2jqO3IrasaqaeD0fCzu3+uGdY3bIOKhJbqRHh/+qRZoBug3ezrY\nWNOGs7K09Th7ug7UDxQNdXcAby4PUz4OYqRnb8WN48L7A9Eb8xSPkQarO8GdmxVd\nkZe5z1c2SGUo4CA0RxHM5X+t+X+Jxt/Y/58ZdTeyzPRl4CyUOzsRutvyytW1wxqd\nLZ0Qjdl4KGp9jdYeXviuRtKY2e6e/m7RbLQ2j8IJdoVHOBQ5iCEyK54yqpReRdZt\npUIT3LmBAgMBAAECggEAMki1/oqhvTyAHwlOvql1JGRkuVKrv1Cy2Y/Jlx76sBH+\naj1wYsqhdAkLu8v7U1/Ex7hwWkHrTrzGQAx3qCh9geT+cmsSN7itY2zlR2YvogIy\n2Bb55b35ZpCr6U1F8PBZSwZ9WRyNiH5wotTqn8WVgSdQTRj1ekrW5pKel0tK2OOE\ncOiCQVrtbVzuHJJy0Js5f40OlVg69V4NPiyU5wmynWRSOFCGi4GZFnf/xf4ZUpjc\nZCFBNgOkWfkV4epFVkzZDTV35kvn1YxCylmEQ/+CwtUBxWL0R7yVOAZHKiU6+Ti1\nCG7BVEI9dfoRjqJ904zc1cYIPkDTbIV4HzJds7UxOwKBgQDKVUdaI2tT3OVvP0c2\nFNPHuAUSbvTkbCkIDKDGzHYb4b5MVzI3D4+62e7AjtguhYiPMQVbsdjoORtWqAti\nXK9/7/7cBKUCXMg9T93gZsKT1f0RTuxY8rgPaSKyJV50ntV6kzIvVQPXLHggJ2ma\nPcHIwA5oyN/dZ8EORNhrY4ka9wKBgQDCaqEm7D0iTZMWl/lGbQsK2MUA+Rr5p62e\n5jFsKAET2sSmp7BrSeWwS1kT+sQ+DwzFor9hrczOCFsVO9CoJJOD9bMhJZoX3uJt\nEOiAT6y9qJ4vGD2qnf02xIeG9tc9Bc3KZbrGi2eY83//Qpn0oFyHN8+bq99baROG\nYuPgrET5RwKBgGB2id8KlefUn7oLFBtPkKxeKmTga3bfriw9QQWmgwTF+mERDUq8\n64xszGwXbi+30CRcfa56uuv0FfmZglvxzmYTeJFS0YyvyXOZuTF8LHYpBk8TLpE1\nntUSDc2bDU5ST3rx5HI2eO9ELz09LRaxLMtV7Ui9xCUdiygPYJLKUJp/AoGAMtTM\nQ6/6n/BmZ77eZwJ1o6VfhMyct++WXnhTLbMb7QQC7IvlfXe5vSlGJgonqw4mSbou\njaxyYuAeaGPWP1Ao3ZSs/BqnulwFGX0VPQ8X3BKtISUWYniiTuJ9iNUbG5Jb5vJI\nLkcelAf+TFAujp4q8xOtjUcXw/+qIjXS3NhNxFsCgYEApsN4O3ycl4F3HHCPYBa5\nv6cuP4Q0MgeQs1ZSF+jt8cM08ESJQHgeat6H6zflERPrUD3ABfxxQ/MNYa1HLPUC\nsQjzx3JSnD5jDqXvQ7sB/3yrSrciDDlYOQ7PWiplYGHa6RkNu2kd2GoJLJMLcV8A\nVNYSNJncSPM58nhQJ7hqJ4w=\n-----END PRIVATE KEY-----\n",
  "client_email": "drive-api-service@bps3505-zonaintegritas.iam.gserviceaccount.com",
  "client_id": "114913066849671866860",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/drive-api-service%40bps3505-zonaintegritas.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

# Tentukan scope untuk mengakses metadata Google Drive dan pembuatan file
SCOPES = ['https://www.googleapis.com/auth/drive']

# Autentikasi dan bangun layanan API Google Drive
credentials = service_account.Credentials.from_service_account_info(
    SERVICE_ACCOUNT_INFO, scopes=SCOPES)
service = build('drive', 'v3', credentials=credentials)

# Fungsi untuk mengambil detail file secara rekursif dengan percobaan ulang jika gagal
def list_files(service, folder_id, path=''):
    files_info = []
    retry_count = 0
    max_retries = 1

    while retry_count < max_retries:
        try:
            results = service.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, owners, webViewLink, size)").execute()

            items = results.get('files', [])
            logging.info(f"Ditemukan {len(items)} item dalam folder {folder_id}")

            if not items:
                logging.info('Tidak ada file yang ditemukan.')
            else:
                for item in items:
                    if item['mimeType'] == 'application/vnd.google-apps.folder':
                        # Masuk ke dalam folder secara rekursif
                        files_info.extend(list_files(service, item['id'], path + item['name'] + '->'))
                    else:
                        # Mengumpulkan detail file
                        files_info.append({
                            'Location': path,
                            'File Name': item['name'],
                            'Last Modified': item['modifiedTime'],
                            'Owner': item['owners'][0]['emailAddress'] if 'owners' in item else 'Tidak ada info pemilik',
                            'File Link': item.get('webViewLink', 'Tidak ada link tersedia'),
                            'File Size (Bytes)': item.get('size', 'Ukuran tidak tersedia')
                        })
            return files_info
        except HttpError as error:
            logging.error(f"HttpError terjadi: {error}")
            if error.resp.status == 500:
                retry_count += 1
                logging.info(f"Mencoba ulang {retry_count}/{max_retries}...")
                time.sleep(2 ** retry_count)  # Penundaan eksponensial
            else:
                raise

# Fungsi untuk menjalankan pekerjaan
def job():
    logging.info("Menjalankan pekerjaan...")

    # Tentukan ID folder Google Drive untuk memulai
    folder_id = '1huIm8LcSgRBMrRx4lcwoH4rICVhjgy4p'
    folder_upload = '1sPsHuqK5Op6HxLs2uMI0GYD3neJAc-uD'

    # Ambil detail file dimulai dari folder yang ditentukan
    file_details = list_files(service, folder_id)

    # Log jumlah file yang ditemukan
    logging.info(f"Total file ditemukan: {len(file_details)}")

    if not file_details:
        logging.warning("Tidak ada file yang ditemukan dalam folder yang ditentukan.")
        return

    # Konversi detail file ke dalam DataFrame pandas
    df = pd.DataFrame(file_details)
    logging.info(f"DataFrame dibuat dengan {len(df)} baris")

    # Simpan DataFrame ke dalam file Excel secara lokal
    output_file = 'google_drive_files.xlsx'
    df.to_excel(output_file, index=False)

    logging.info(f"Detail file telah ditulis ke {output_file}")

    # Periksa apakah file sudah ada di folder tujuan
    existing_files = service.files().list(
        q=f"'{folder_upload}' in parents and name='cek-data.xlsx' and trashed=false",
        fields="files(id, name)"
    ).execute()

    if existing_files['files']:
        # Jika file sudah ada, perbarui (overwrite)
        file_id = existing_files['files'][0]['id']
        try:
            media = MediaFileUpload(output_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            updated_file = service.files().update(fileId=file_id, media_body=media).execute()
            logging.info(f"File yang ada diperbarui dengan ID: {file_id}")
        except Exception as e:
            logging.error(f"Gagal memperbarui file yang ada: {e}")
    else:
        # Jika file tidak ada, buat file baru
        file_metadata = {
            'name': 'cek-data.xlsx',
            'parents': [folder_upload]
        }
        media = MediaFileUpload(output_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        try:
            file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            logging.info(f"File telah diupload ke Google Drive dengan ID: {file.get('id')}")
        except Exception as e:
            logging.error(f"Gagal mengupload file: {e}")

# Panggil langsung fungsi job untuk mengeksekusinya sekali
job()
