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
    worksheet = sh.get_worksheet(sheet_name)
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
        elif id=='Phạm Thị Hà':
            colors.append('pink')
        elif id=='Chưa bắt đầu'or id=='chưa bắt đầu':
            colors.append('skyblue')   
        else:
            # Tạo màu ngẫu nhiên cho các id còn lại
            #colors.append(f'#{random.randint(0, 0xFFFFFF):06x}')
            while True:
                color = f'#{random.randint(0, 0xFFFFFF):06x}'
                if color not in ['#ff0000', '#00ff00' ,'#ffc0cb', '#87ceeb'] and color not in used_colors:
                    colors.append(color)
                    used_colors.add(color)  # Thêm màu vào tập hợp màu đã sử dụng
                    break
    
    # Vẽ biểu đồ cột
    plt.bar(ids, results, color=colors)
    plt.title(title)
    # Hiển thị giá trị trên mỗi cột
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
        elif id == 'Pass' or id == 'Đã hoàn thành':
            colors.append('green')
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
def filter_names1(data,row_name, owner_name='Fail'):
    """Lọc các tên dựa trên tên chủ sở hữu."""
    filtered_names = [row[row_name] for row in data if row['Result'].strip() == owner_name]
    return [name for name in filtered_names if name.strip()]
def filter_names(data, row_name):
    """Lọc các tên dựa trên tên chủ sở hữu."""
    filtered_names = [row[row_name] for row in data]
    return [name for name in filtered_names if name.strip()]
def create_dataframe(names1,row_name):
    """Tạo DataFrame từ hai danh sách tên."""
    combined_names = names1 
    return pd.DataFrame(combined_names, columns=[row_name])
def create_dataframe1(names1, names2,row_name):
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
def update_google_sheet_with_image_links(worksheet,name, bar_image_url,bar_image_url1,bar_image_url2,bar_image_url3,bar_image_url4,t ):
    worksheet.update_acell(f'A{t}', f'Biểu đồ {name}')
    format_text(worksheet, f'A{t}',24)

    worksheet.update_acell(f'A{t+1}', f'Biểu đồ đánh giá dựa trên Group Name ')
    worksheet.update_acell(f'C{t+2}', f'=IMAGE("{bar_image_url}")')
    worksheet.update_acell(f'D{t+1}', f'Biểu đồ đánh giá dựa trên Result ')
    worksheet.update_acell(f'F{t+2}', f'=IMAGE("{bar_image_url1}")')
    worksheet.update_acell(f'G{t+1}', f'Biểu đồ đánh giá dựa trên Tester')
    worksheet.update_acell(f'I{t+2}', f'=IMAGE("{bar_image_url2}")')
    worksheet.update_acell(f'K{t+1}', f'Biểu đồ đánh giá dựa trên những ngày test có test case fail ')
    worksheet.update_acell(f'M{t+2}', f'=IMAGE("{bar_image_url3}")')
    worksheet.update_acell(f'N{t+1}', f'Biểu đồ đánh giá dựa trên số test case fail được Dev sửa')
    worksheet.update_acell(f'P{t+2}', f'=IMAGE("{bar_image_url4}")')

    format_text(worksheet, f'A{t+1}',16)
    format_text(worksheet, f'D{t+1}',16)
    format_text(worksheet, f'G{t+1}',16)
    format_text(worksheet, f'K{t+1}',16)
    format_text(worksheet, f'N{t+1}',16)


def format_text(worksheet, cell_range,font):
    worksheet.format(cell_range, {
        "textFormat": {
            "bold": True,
            "fontSize": font
        }
    })

def data_table(worksheet,data1,start_row,end_row,start_col, end_col,update_range,row_name):
    filtered_names = filter_names(data1,row_name=row_name)
    
    df = create_dataframe(filtered_names, row_name=row_name)
    counts = df[row_name].value_counts()   
    update_sheet_with_data(worksheet, counts.index.tolist(), counts.values.tolist(), update_range, row_name=row_name)
    clear_data_in_range(worksheet, start_row=len(counts) + start_row, end_row=end_row, start_col=start_col, end_col=end_col)
    if  not filtered_names:
        #print("Không có dữ liệu để tạo biểu đồ.")
        return
    bar_image_stream = create_bar_chart(counts.index.tolist(), counts.values.tolist(),f'biểu đồ {row_name}')
    return bar_image_stream,df
def data_table1(worksheet,data1,start_row,end_row,start_col, end_col,update_range,row_name):
    filtered_names = filter_names1(data1,row_name=row_name)
    df = create_dataframe(filtered_names, row_name=row_name)
    counts = df[row_name].value_counts()
    update_sheet_with_data(worksheet, counts.index.tolist(), counts.values.tolist(), update_range, row_name=row_name)
    clear_data_in_range(worksheet, start_row=len(counts) + start_row, end_row=end_row, start_col=start_col, end_col=end_col)
    if  not filtered_names:
        #print("Không có dữ liệu để tạo biểu đồ.")
        return
    bar_image_stream = create_bar_chart(counts.index.tolist(), counts.values.tolist(),f'biểu đồ {row_name}')
    return bar_image_stream,df

def data_all(worksheet,data_df,start_row,end_row,start_col, end_col,update_range,row_name):
    if not isinstance(data_df, list) or len(data_df) == 0 or not all(isinstance(df, pd.DataFrame) for df in data_df):
        return None
    combined_df = pd.concat(data_df, ignore_index=True)
    
    # Lọc dữ liệu dựa trên row_name
    filtered_df = filter_names(combined_df.to_dict('records'), row_name=row_name)
    df = pd.DataFrame(filtered_df, columns=[row_name])
    counts = df[row_name].value_counts()
    bar_image_stream = create_bar_chart(counts.index.tolist(), counts.values.tolist(),f'biểu đồ {row_name}')
    update_sheet_with_data(worksheet, counts.index.tolist(), counts.values.tolist(), update_range, row_name=row_name)
    clear_data_in_range(worksheet, start_row=len(counts) + start_row, end_row=end_row, start_col=start_col, end_col=end_col)
    return bar_image_stream

def task1(start_col=0,end_col=1):
    
    creds = authenticate_google_account()
    t=18
    G_N=[]
    Result=[]
    Tester=[]
    TestDate=[]
    EFix=[]
    output_sheet_name = "Biểu đồ testcase"
    
    for i in range(6,24):
        sheet_name=i
        sheet_key='1phtfYoUPC3Crjf_DrThb2p55M5d_wMUnOpR2_EyirVU'
        data_1, _, sh = get_google_sheets_data(creds, sheet_key, sheet_name)
        worksheet = get_or_create_sheet(sh, output_sheet_name)
        worksheets = sh.worksheets()
        name = worksheets[sheet_name].title
        
        #
        update_range=f'A{t+1}:B'
        bar_image_stream,data=data_table(worksheet,data_1,t+2,t+7,start_col=start_col,end_col=end_col,update_range=update_range,row_name='Group Name')
        bar_image_url = upload_image_to_drive(creds, bar_image_stream, f'{name}_chart1.png')
        G_N.append(pd.DataFrame(data))
        #
        update_range=f'D{t+1}:E'
        bar_image_stream1,data1=data_table(worksheet,data_1,t+2,t+7,start_col=3,end_col=4,update_range=update_range,row_name='Result')
        bar_image_url1 = upload_image_to_drive(creds, bar_image_stream1, f'{name}_chart2.png')
        Result.append(pd.DataFrame(data1))
        #
        update_range=f'G{t+1}:H'
        bar_image_stream2,data2=data_table(worksheet,data_1,t+2,t+7,start_col=6,end_col=7,update_range=update_range,row_name='Tester')
        bar_image_url2 = upload_image_to_drive(creds, bar_image_stream2, f'{name}_chart3.png')
        Tester.append(pd.DataFrame(data2))

        #
        update_range=f'K{t+1}:L'
        result = data_table1(worksheet, data_1, t+2, t+7, start_col=10, end_col=11, update_range=update_range, row_name='Test date')

        if result is not None:
            bar_image_stream3, data3 = result
            if bar_image_stream3 is not None:
                bar_image_url3 = upload_image_to_drive(creds, bar_image_stream3, f'{name}_chart4.png')
                TestDate.append(pd.DataFrame(data3))
            else:
                bar_image_url3 = ''
        else:
            bar_image_url3 = ''
        # bar_image_stream3,data3=data_table1(worksheet,data_1,t+2,t+7,start_col=9,end_col=10,update_range=update_range,row_name='Test date')
        # if(bar_image_stream3 != None):
        #     bar_image_url3 = upload_image_to_drive(creds, bar_image_stream3, f'{name}_chart4.png')
        #     TestDate.append(data3)
        # else:
        #     bar_image_url3=''  
        #    
        update_range=f'N{t+1}:O'
        result = data_table(worksheet, data_1, t+2, t+7, start_col=13, end_col=14, update_range=update_range, row_name='Employee Fix')

        if result is not None:
            bar_image_stream4, data4 = result
            if bar_image_stream4 is not None:
                bar_image_url4 = upload_image_to_drive(creds, bar_image_stream4, f'{name}_chart5.png')
                EFix.append(pd.DataFrame(data4))
            else:
                bar_image_url4 = ''
        else:
            bar_image_url4 = ''

        # bar_image_stream4,data4=data_table1(worksheet,data_1,t+2,t+7,start_col=12,end_col=13,update_range=update_range,row_name='Employee Fix')
        # if(bar_image_stream4 != None):
        #     bar_image_url4 = upload_image_to_drive(creds, bar_image_stream4, f'{name}_chart5.png')
        #     EFix.append(data4)
        # else:
        #     bar_image_url4=''      



        update_google_sheet_with_image_links(worksheet,name, bar_image_url,bar_image_url1,bar_image_url2,bar_image_url3,bar_image_url4,t-1)
        t=t+9
        time.sleep(10)    
    t=2
    #print(Result)
    #biểu đồ tổng quát    
    name='tổng quát'
    update_range='A3:B' 
    bar_image_stream=data_all(worksheet,G_N,t+2,t+14,start_col=start_col,end_col=end_col,update_range=update_range,row_name='Group Name')
    bar_image_url = upload_image_to_drive(creds, bar_image_stream, f'{name}_chart1.png')

    # #
    update_range=f'D{t+1}:E'
    bar_image_stream1=data_all(worksheet,Result,t+2,t+14,start_col=3,end_col=4,update_range=update_range,row_name='Result')
    bar_image_url1 = upload_image_to_drive(creds, bar_image_stream1, f'{name}_chart2.png')
    #
    update_range=f'G{t+1}:H'
    bar_image_stream2=data_all(worksheet,Tester,t+2,t+14,start_col=6,end_col=7,update_range=update_range,row_name='Tester')
    bar_image_url2 = upload_image_to_drive(creds, bar_image_stream2, f'{name}_chart3.png')

    #
    update_range=f'K{t+1}:L'
    bar_image_stream3=data_all(worksheet,TestDate,t+2,t+14,start_col=10,end_col=11,update_range=update_range,row_name='Test date')
    if(bar_image_stream3 != None):
        bar_image_url3 = upload_image_to_drive(creds, bar_image_stream3, f'{name}_chart4.png')
    else:
        bar_image_url3=''  
    #    
    update_range=f'N{t+1}:O'
    bar_image_stream4=data_all(worksheet,EFix,t+2,t+14,start_col=13,end_col=14,update_range=update_range,row_name='Employee Fix')
    if(bar_image_stream4 != None):
        bar_image_url4 = upload_image_to_drive(creds, bar_image_stream4, f'{name}_chart5.png')
    else:
        bar_image_url4=''

    update_google_sheet_with_image_links(worksheet,name, bar_image_url,bar_image_url1,bar_image_url2,bar_image_url3,bar_image_url4,t-1)

    print(f"uploaded to your Google Drive and linked in sheet '{output_sheet_name}' successfully.")
def shift_letter(letter, shift):
    # Chuyển ký tự thành mã ASCII và cộng thêm giá trị shift
    new_code = ord(letter) + shift
    # Chuyển mã ASCII mới thành ký tự
    new_letter = chr(new_code)
    return new_letter
if __name__ == "__main__":
    task1()
       
