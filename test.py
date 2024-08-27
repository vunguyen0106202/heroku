import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Xác thực và kết nối với Google Sheets
creds = Credentials.from_service_account_file('api.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
gc = gspread.authorize(creds)
sh = gc.open_by_key('1dDFibYofUv0gAGRyz018FHCINravASEsYinBgwhWuh8')
sheet_name = 't'
worksheet = sh.worksheet(sheet_name)

# Đọc dữ liệu từ dòng 2 trở đi
data = worksheet.get_all_records(head=1)
names = [row['B'] for row in data]
values = [row['C'] for row in data]

# Tạo hoặc mở sheet mới
sheet_title = 'Biểu đồ'
existing_sheets = [sheet.title for sheet in sh.worksheets()]

if sheet_title in existing_sheets:
    new_sheet = sh.worksheet(sheet_title)
else:
    new_sheet = sh.add_worksheet(title=sheet_title, rows="100", cols="20")

# Cập nhật dữ liệu vào sheet
update_range = 'A1:B'
values_to_update = [['Tên', 'Số']] + list(zip(names, values))
new_sheet.update(range_name=update_range, values=values_to_update)

# Tạo kết nối với Google Sheets API
service = build('sheets', 'v4', credentials=creds)

def save_chart_id_to_sheet(sheet, chart_id):
    # Cập nhật ID biểu đồ vào ô C1
    sheet.update('C1', [[chart_id]])

def load_chart_id_from_sheet(sheet):
    try:
        # Đọc ID biểu đồ từ ô C1
        chart_id = sheet.acell('C1').value
        return chart_id
    except Exception as e:
        print(f"Đã xảy ra lỗi khi lấy ID biểu đồ từ sheet: {e}")
        return None

def create_chart(sheet_id):
    try:
        requests = [{
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Biểu đồ mẫu",
                        "basicChart": {
                            "chartType": "COLUMN",
                            "legendPosition": "BOTTOM_LEGEND",
                            "axis": [
                                {"position": "BOTTOM_AXIS", "title": "Tên"},
                                {"position": "LEFT_AXIS", "title": "Số"}
                            ],
                            "domains": [
                                {"domain": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(values) + 1, "startColumnIndex": 0, "endColumnIndex": 1}]}}}
                            ],
                            "series": [
                                {"series": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(values) + 1, "startColumnIndex": 1, "endColumnIndex": 2}]}}}
                            ]
                        }
                    },
                    "position": {"overlayPosition": {"anchorCell": {"sheetId": sheet_id, "rowIndex": 5, "columnIndex": 5}}}
                }
            }
        }]

        batch_update_request = {"requests": requests}
        response = service.spreadsheets().batchUpdate(spreadsheetId=sh.id, body=batch_update_request).execute()
        chart_id = response['replies'][0]['addChart']['chart']['chartId']
        save_chart_id_to_sheet(new_sheet, chart_id)  # Lưu ID biểu đồ vào sheet
        print(f"Biểu đồ đã được tạo với ID: {chart_id}")
        return chart_id
    except HttpError as err:
        print(f"Đã xảy ra lỗi khi tạo biểu đồ: {err}")
        return None

def update_chart(sheet_id, chart_id):
    try:
        requests = [{
            "updateChartSpec": {
                "chartId": chart_id,
                "spec": {
                    "title": "Biểu đồ mẫu",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Tên"},
                            {"position": "LEFT_AXIS", "title": "Số"}
                        ],
                        "domains": [
                            {"domain": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(values) + 1, "startColumnIndex": 0, "endColumnIndex": 1}]}}}
                        ],
                        "series": [
                            {"series": {"sourceRange": {"sources": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": len(values) + 1, "startColumnIndex": 1, "endColumnIndex": 2}]}}}
                        ]
                    }
                }
            }
        }]

        batch_update_request = {"requests": requests}
        service.spreadsheets().batchUpdate(spreadsheetId=sh.id, body=batch_update_request).execute()
        print(f"Biểu đồ với ID {chart_id} đã được cập nhật.")
    except HttpError as err:
        print(f"Đã xảy ra lỗi khi cập nhật biểu đồ: {err}")

# Xử lý biểu đồ
chart_id = load_chart_id_from_sheet(new_sheet)  # Tải ID biểu đồ từ ô

# Nếu đã có ID biểu đồ, cập nhật biểu đồ
if chart_id:
    update_chart(new_sheet.id, chart_id)
else:
    # Nếu không có ID biểu đồ, tạo mới
    chart_id = create_chart(new_sheet.id)
