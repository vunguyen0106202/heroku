import io
import os
import pickle
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import matplotlib.pyplot as plt
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import pandas as pd
import random
import time

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def authenticate_google_account():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def get_google_sheets_data(creds, spreadsheet_key, sheet_name):
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_key)
    worksheet = sh.worksheet(sheet_name)
    data = worksheet.get_all_records()
    return data, worksheet, sh

def get_or_create_sheet(sh, sheet_name):
    try:
        worksheet = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=sheet_name, rows="100", cols="20")
        print(f"Created new sheet '{sheet_name}'.")
    return worksheet

def create_bar_chart(ids, results,title):
    # Xác định màu sắc cho các id cụ thể
    colors = []
    used_colors = set()
    for id in ids:
        
        id = id.replace('Ð', 'Đ').strip()
        if id == 'Fail' or id == 'High':
            colors.append('red')
        elif id == 'Pass' or id == 'Đã hoàn thành':
            colors.append('green')
        elif id=='Chưa bắt đầu'or id=='chưa bắt đầu':
            colors.append('skyblue')    
        else:
            # Tạo màu ngẫu nhiên cho các id còn lại
            #colors.append(f'#{random.randint(0, 0xFFFFFF):06x}')
            while True:
                color = f'#{random.randint(0, 0xFFFFFF):06x}'
                if color not in ['#ff0000', '#00ff00'] and color not in used_colors:
                    colors.append(color)
                    used_colors.add(color)  # Thêm màu vào tập hợp màu đã sử dụng
                    break
    
    # Vẽ biểu đồ cột
    plt.bar(ids, results, color=colors)
    plt.title(title)
    for i in range(len(ids)):
        plt.text(i, results[i], str(results[i]), ha='center', va='bottom')
    # Lưu biểu đồ vào luồng hình ảnh
    bar_image_stream = io.BytesIO()
    plt.savefig(bar_image_stream, format='png')
    plt.close()
    bar_image_stream.seek(0)
    
    return bar_image_stream

def create_pie_chart(ids, results,title):
    colors = []
    used_colors = set()
    for id in ids:
        
        id = id.replace('Ð', 'Đ').strip()
        if id == 'Fail' or id == 'High':
            colors.append('red')    
        else:
            # Tạo màu ngẫu nhiên cho các id còn lại
            #colors.append(f'#{random.randint(0, 0xFFFFFF):06x}')
            while True:
                color = f'#{random.randint(0, 0xFFFFFF):06x}'
                if color not in ['#ff0000', '#00ff00'] and color not in used_colors:
                    colors.append(color)
                    used_colors.add(color)  # Thêm màu vào tập hợp màu đã sử dụng
                    break
    plt.pie(results, labels=ids, autopct='%1.1f%%', startangle=90,colors=colors)
    plt.axis('equal')
    plt.title(title)
    pie_image_stream = io.BytesIO()
    plt.savefig(pie_image_stream, format='png')
    plt.close()
    pie_image_stream.seek(0)
    return pie_image_stream

def create_line_chart(ids, results):
    plt.plot(ids, results, marker='o', linestyle='-', color='b')
    plt.title('')
    line_image_stream = io.BytesIO()
    plt.savefig(line_image_stream, format='png')
    plt.close()
    line_image_stream.seek(0)
    return line_image_stream

def upload_image_to_drive(creds, image_stream, filename):
    drive_service = build('drive', 'v3', credentials=creds)
    results = drive_service.files().list(q=f"name='{filename}' and mimeType='image/png'", fields="files(id)").execute()
    for file in results.get('files', []):
        drive_service.files().delete(fileId=file.get('id')).execute()
    file_metadata = {'name': filename}
    media = MediaIoBaseUpload(image_stream, mimetype='image/png')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    file_id = file.get('id')
    drive_service.permissions().create(fileId=file_id, body={'role': 'reader', 'type': 'anyone'}).execute()
    return f"https://drive.google.com/uc?id={file_id}"
def filter_names(data, owner_name,row_name):
    """Lọc các tên dựa trên tên chủ sở hữu."""
    filtered_names = [row[row_name] for row in data if row['Chủ sở hữu'].strip() == owner_name]
    return [name for name in filtered_names if name.strip()]

def create_dataframe(names1, names2,row_name):
    """Tạo DataFrame từ hai danh sách tên."""
    combined_names = names1 + names2
    return pd.DataFrame(combined_names, columns=[row_name])
def update_sheet_with_data(sheet, names, values, update_range,row_name):
    """Cập nhật dữ liệu vào sheet."""
    values_to_update = [[row_name, 'Số lần xuất hiện']] + list(zip(names, values))
    sheet.update(range_name=update_range, values=values_to_update)
def clear_data_in_range(sheet, start_row, end_row, start_col, end_col):
    """Xóa dữ liệu trong phạm vi chỉ định mà không xóa các hàng hoặc cột."""
    range_to_clear = f'{chr(start_col + 65)}{start_row}:{chr(end_col + 65)}{end_row}'
    sheet.batch_clear([range_to_clear])
    print(f"Đã xóa dữ liệu trong phạm vi: {range_to_clear}")    
def update_google_sheet_with_image_links(worksheet,name, bar_image_url, bar_image_url1,pie_image_url,row,row_char ):  
    worksheet.update_acell(f'{row}4', f'{name}')
    worksheet.update_acell(f'{row_char}5', f'=IMAGE("{bar_image_url}")')   
    worksheet.update_acell(f'{row}13', f'{name}')
    worksheet.update_acell(f'{row_char}14', f'=IMAGE("{ bar_image_url1}")')
    worksheet.update_acell(f'{row}23', f'{name}')
    worksheet.update_acell(f'{row_char}24', f'=IMAGE("{pie_image_url}")')
    format_text(worksheet,f'{row}4',16)
    format_text(worksheet,f'{row}13',16)
    format_text(worksheet,f'{row}23',16)
    if row=='A':
        worksheet.update_acell(f'{row}22', f'Biểu đồ đánh giá dựa trên category của công việc')
        worksheet.update_acell(f'{row}12', f'Biểu đồ đánh giá dựa trên trạng thái của công việc')
        worksheet.update_acell(f'{row}3', f'Biểu đồ đánh giá dựa trên mức độ ưu tiên của công việc')
        format_text(worksheet, f'{row}22',24)
        format_text(worksheet, f'{row}12',24)
        format_text(worksheet, f'{row}3',24)

def format_text(worksheet, cell_range,font):
    worksheet.format(cell_range, {
        "textFormat": {
            "bold": True,
            "fontSize": font
        }
    })

def data_table(worksheet,data1,data2,name,start_row,end_row,start_col, end_col,update_range,row_name):
    filtered_names = filter_names(data1, name,row_name=row_name)
    filtered_names1 = filter_names(data2,name,row_name=row_name)
    
    df = create_dataframe(filtered_names, filtered_names1,row_name=row_name)
    counts = df[row_name].value_counts()
    
    update_sheet_with_data(worksheet, counts.index.tolist(), counts.values.tolist(), update_range, row_name=row_name)
    clear_data_in_range(worksheet, start_row=len(counts) + start_row, end_row=end_row, start_col=start_col, end_col=end_col)
    if not filtered_names1 and not filtered_names:
        print("Không có dữ liệu để tạo biểu đồ.")
        return
    bar_image_stream = create_bar_chart(counts.index.tolist(), counts.values.tolist(),f'biểu đồ {row_name}')
    return bar_image_stream
def data_table1(worksheet,data1,data2,name,start_row,end_row,start_col, end_col,update_range,row_name):
    filtered_names = filter_names(data1, name,row_name=row_name)
    filtered_names1 = filter_names(data2,name,row_name=row_name)
    
    df = create_dataframe(filtered_names, filtered_names1,row_name=row_name)
    counts = df[row_name].value_counts()
    
    update_sheet_with_data(worksheet, counts.index.tolist(), counts.values.tolist(), update_range, row_name=row_name)
    clear_data_in_range(worksheet, start_row=len(counts) + start_row, end_row=end_row, start_col=start_col, end_col=end_col)
    if not filtered_names1 and not filtered_names:
        print("Không có dữ liệu để tạo biểu đồ.")
        return
    bar_image_stream = create_pie_chart(counts.index.tolist(), counts.values.tolist(),f'biểu đồ {row_name}')
    return bar_image_stream
def task1(name,start_col,end_col,row,row1):
    update_range=f'{row}5:{row1}'
    row_char=shift_letter(row, 2)
    creds = authenticate_google_account()
    sheet_name1 = "BackEnd"  
    sheet_name2='FrontEnd'
    sheet_key='1phtfYoUPC3Crjf_DrThb2p55M5d_wMUnOpR2_EyirVU'
    data1, _, sh = get_google_sheets_data(creds, sheet_key, sheet_name1)
    data2, _,sh=get_google_sheets_data(creds,sheet_key,sheet_name2)
    output_sheet_name = "Biểu đồ công việc"
    worksheet = get_or_create_sheet(sh, output_sheet_name)
    #
    bar_image_stream=data_table(worksheet,data1,data2,name,6,11,start_col=start_col,end_col=end_col,update_range=update_range,row_name='Mức độ ưu tiên')
    bar_image_url = upload_image_to_drive(creds, bar_image_stream, f'{name}_chart1.png')
    #
    update_range1=f'{row}14:{row1}'
    bar_image_stream1=data_table(worksheet,data1,data2,name,15,21,start_col=start_col,end_col=end_col,update_range=update_range1,row_name='Trạng thái')
    bar_image_url1= upload_image_to_drive(creds, bar_image_stream1, f'{name}_chart2.png')
    #
    update_range2=f'{row}24:{row1}'
    pie_image_stream2=data_table1(worksheet,data1,data2,name,25,50,start_col=start_col,end_col=end_col,update_range=update_range2,row_name='Category')
    pie_image_url2= upload_image_to_drive(creds, pie_image_stream2, f'{name}_chart3.png')  
    ###
    update_google_sheet_with_image_links(worksheet,name, bar_image_url,bar_image_url1,pie_image_url2,row,row_char)
    
    
    print(f"uploaded to your Google Drive and linked in sheet '{output_sheet_name}' successfully.")
    
def shift_letter(letter, shift):
    # Chuyển ký tự thành mã ASCII và cộng thêm giá trị shift
    new_code = ord(letter) + shift
    # Chuyển mã ASCII mới thành ký tự
    new_letter = chr(new_code)
    return new_letter
if __name__ == "__main__":
    name=['Đỗ Phương Nam','Nguyễn Đình Thắng','Nguyễn Văn Khánh','Phạm Thị Hà']
    t=0
    original_letter = 'A'
    for i in name:
       start_col=0+t
       end_col=1+t
       row=shift_letter(original_letter, start_col)
       row1= shift_letter(original_letter, end_col)
       task1(i,start_col, end_col,row,row1)
       time.sleep(10)
       t=t+3
       
       