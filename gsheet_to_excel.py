import win32com.client as win32
from Google import Create_Service

CLIENT_SECRET_FILE = 'C:\\Users\\ashag\\Downloads\\createsecretkey34.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

gsheet_id = '1NIEuYvxOUuySWPD6l3qg2EJd7X5krDpSPetUivVDaDY'
gsheet_name = 'sales data'

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)
gs = service.spreadsheets()
rows = gs.values().get(
    spreadsheetId=gsheet_id,
    range=gsheet_name,
).execute().get('values')

#xlApp = win32.Dispatch('Excel.Application')
xlApp=win32.gencache.EnsureDispatch('Excel.Application')
wb = xlApp.Workbooks('Sample.xlsx.xlsx')
wsData = wb.Worksheets("Data")

wsData.Cells.ClearContents()

rowNumber = 1
colCount = len(rows[0])

for row in rows:
    wsData.Range(wsData.Cells(rowNumber, 1), wsData.Cells(rowNumber, colCount)).Value = row
    rowNumber += 1