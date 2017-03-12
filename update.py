import sys
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials

def auth_gss_client(path, scopes):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scopes)

    return gspread.authorize(credentials)

def login2Google(json_path,scopes):
    try:
        _gss_client = auth_gss_client(json_path, scopes)
        return _gss_client
    except Exception:
        _gss_client = -1
        return (_gss_client)

def opensheet(wks, sheet_name):
    sheet = -1
    try:
        sheet = wks.worksheet(sheet_name)
    except Exception:
        return (sheet)
    return (sheet)

def read_cell(wks, row, col):
    val = -1
    try:
        val = wks.cell(row,col).value
    except Exception:
        return (val)
    return (val)

def write_cell(wks, row, col, data):
    val = -1
    try:
        wks.update_cell(row, col, data)
        val = 0
    except Exception:
        return (val)
    return (val)

def open4read_file(path):
    key = -1
    try:
        f = open(path)
        key = f.read()
        f.close()
    except Exception:
        return (key)
    return (key)

#------main loop--------------------------------------------
auth_json_path = 'auth.json'
scopes = ['https://spreadsheets.google.com/feeds']
init_row = 2
trtl_col_qpl_id = 2
trtl_col_pr_id = 7
dil_col_qpl_id = 1
dil_col_pr_id = 13
now = time.strftime("%c")
spreadsheet_key_path = open4read_file('key_path.txt')

if spreadsheet_key_path == -1:
    print('Open key file fail')
    sys.exit(1)

#login to Google, if fail then abort
gss_client = login2Google(auth_json_path, scopes)
if gss_client == -1:
    print('Login fail')
    sys.exit(1)
else:
    wks = gss_client.open_by_key(spreadsheet_key_path)
#open daily issue list on the 2K17 MSAF QPL Google doc
    sheet_daily_issue_list = opensheet(wks, "Daily Issue List")
    sheet_tpe_report_tracking_list = opensheet(wks, "TPE_Report_Tracking_list")

    if sheet_daily_issue_list == -1 or sheet_tpe_report_tracking_list == -1:
        print('Open sheet fail')
        sys.exit(1)
    else:
        i = init_row
        j = init_row
        dil_qpl_id = read_cell(sheet_daily_issue_list, i, dil_col_qpl_id)
        dil_pr_id = read_cell(sheet_daily_issue_list, i, dil_col_pr_id)
        while dil_qpl_id != '':
            try:
                dil_pr_id = dil_pr_id.split('\n')
            except Exception:
                dil_pr_id=''
                print(i, j, 'string split exception')
            for x in range(0,len(dil_pr_id),1):
                print(i, j, dil_qpl_id, dil_pr_id[x])
                write_cell(sheet_tpe_report_tracking_list, j, trtl_col_qpl_id, dil_qpl_id)
                write_cell(sheet_tpe_report_tracking_list, j, trtl_col_pr_id, dil_pr_id[x])
                j += 1
            i += 1
            dil_qpl_id = read_cell(sheet_daily_issue_list, i, dil_col_qpl_id)
            dil_pr_id = read_cell(sheet_daily_issue_list, i, dil_col_pr_id)




#            if i == 3:
#                print('force stop')
#                sys.exit(1)

end = time.strftime("%c")
print(now)
print(end)
print('--End--')
sys.exit(1)