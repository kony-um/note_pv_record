from selenium import webdriver
from time import sleep
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

"""
ダッシュボードにあるデータを取得して辞書型で返す
@param webdriver
@return dictionary
"""
def get_dashboard_data(driver):
    datas = {}
    elem_view = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div > div:nth-child(2) > div > div > div:nth-child(2) > div.stats-count > div:nth-child(1) > p > span.num")
    datas['view'] = elem_view.text

    elem_commwnt = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div > div:nth-child(2) > div > div > div:nth-child(2) > div.stats-count > div:nth-child(3) > p > span.num")
    datas['comment'] = elem_commwnt.text

    elem_like = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div > div:nth-child(2) > div > div > div:nth-child(2) > div.stats-count > div:nth-child(5) > p > span.num")
    datas['like'] = elem_like.text

    return datas

"""
週・月・年・全期間を探索してそれぞれの中のアクセス情報を取得して辞書型で返す
@param webdriver
@return dictionary
"""
def coverage_dashboard(driver):
    # 月間データの取得
    monthly = get_dashboard_data(driver)

    # 週間データの取得
    elem_week = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div > div:nth-child(2) > div > div > div:nth-child(2) > div.period-nav > ul > li:nth-child(1) > a")
    elem_week.click()

    sleep(2)

    weekly = get_dashboard_data(driver)

    # 年間データの取得
    elem_year = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div > div:nth-child(2) > div > div > div:nth-child(2) > div.period-nav > ul > li:nth-child(3) > a")
    elem_year.click()

    sleep(2)

    yearly = get_dashboard_data(driver)

    # 全期間データの取得
    elem_all = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div > div:nth-child(2) > div > div > div:nth-child(2) > div.period-nav > ul > li:nth-child(4) > a")
    elem_all.click()

    sleep(2)

    all_term = get_dashboard_data(driver)

    datas = {}
    datas['weekly'] = weekly
    datas['monthly'] = monthly
    datas['yearly'] = yearly
    datas['all_term'] = all_term

    return datas


"""
noteにログインしてダッシュボードのページに移動する
@param string
@param string
@param webdriver
"""
def login_note(id, password, driver):
    driver.get("https://note.mu/sitesettings/stats")

    elem_search_word = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div.register-container > form > div > div:nth-child(1) > input")
    elem_search_word.send_keys(id) # note IDを入力する

    elem_search_word = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div.register-container > form > div > div:nth-child(2) > input")
    elem_search_word.send_keys(password) # パスワードを書き込む
    elem_search_btn = driver.find_element_by_css_selector("body > div.clearfix.space--under-header > div.register-container > form > button > div:nth-child(2)")
    elem_search_btn.click()

    sleep(5) # ページ読み込み中なので処理を一時停止


"""
スプレッドシートを開く
@return spreadsheet
"""
def open_spreadsheet():
    scope = ['https://spreadsheets.google.com/feeds']
    doc_id = '' # 自分のものを入れる
    path = os.path.expanduser('') # 自分のものを入れる

    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile   = client.open_by_key(doc_id)
    return gfile

"""
更新する対象のシートを選択する
@param string
@param spreadsheet
@return worksheet
"""
def select_sheet(year_month, gfile): 
    # year_monthがあればそれを選択する。
    worksheets = gfile.worksheets()

    if (search_sheet(year_month, worksheets)):
    	return gfile.worksheet(year_month)
    else:
    	return add_sheet(year_month, gfile)

"""
スプレッドシートの中にその年月のシートがあるかを判定する
@param string
@param list
@return bool
"""
def search_sheet(year_month, worksheets):
	for sheet in worksheets:
		if sheet.title == year_month:
			return True

	return False

"""
その年月のシートがなかった場合には新規作成する
@param string
@param spreadsheet
@return worksheet
"""
def add_sheet(year_month, gfile):
    # sheetを作る
    sheet = gfile.add_worksheet(year_month, 40, 15)

    # どの期間なのかわかるようにする
    sheet.update_acell('C2', '週')
    sheet.update_acell('F2', '月')
    sheet.update_acell('I2', '年')
    sheet.update_acell('L2', '全期間')

    # T列の意味を書き込む()
    header = ['PV', 'コメント', 'スキ']
    header = header * 4

    cell_list = sheet.range('C3:N3')
    count = 0
    for cell in cell_list:
        cell.value = header[count]
        count+=1

    sheet.update_cells(cell_list)

    return sheet

"""
PV情報をワークシートに書き込む
@param dictionary
@param worksheet
"""
def write_dashboard_data(dashboard_data, sheet):
    now = datetime.datetime.now()
    today = int(now.strftime("%d"))
    row_num = str(today + 3)

    date = now.strftime("%m/%d")

    sheet.update_acell('B' + row_num, date)

    for k,v in dashboard_data.items():
        cols_range = get_cols_range(k)
        write_data(v, cols_range, sheet)

"""
期間に合わせて、書き込むセルを変える処理
@param string
@return list
"""
def get_cols_range(key):
    if key == 'weekly':
        return ['C', 'D', 'E']
    if key == 'monthly':
        return ['F', 'G', 'H']
    if key == 'yearly':
        return ['I', 'J', 'K']
    if key == 'all_term':
        return ['L', 'M', 'N']

"""
実際にデータを書き込む処理
@param dictionary
@param list
@param worksheet 
"""
def write_data(value, cols_range, sheet):
    now = datetime.datetime.now()
    today = int(now.strftime("%d"))
    row_num = str(today + 3)

    sheet.update_acell(cols_range[0] + row_num, value['view'])
    sheet.update_acell(cols_range[1] + row_num, value['comment'])
    sheet.update_acell(cols_range[2] + row_num, value['like'])

"""
メインの処理。
・noteにログインする
・データを取ってくる
・シートを選択する
・そこにアクセス情報を書き込む
"""
def main():
    ## Seleniumでnoteからデータを取得する処理
    driver = webdriver.Chrome()

    # ダッシュボードに飛ぶ前提でnoteにログインする
    id = '' # 自分のものを入れる
    password = '' # 自分のものを入れる
    login_note(id, password, driver)

    # ダッシュボードで各情報を取得していく
    dashboard_data = coverage_dashboard(driver)

    driver.quit()


    ## スプレッドシートの処理
    # スプレッドシートを取得する
    gfile = open_spreadsheet() 


    now = datetime.datetime.now()
    year_month = now.strftime("%Y%m")

    sheet = select_sheet(year_month, gfile)

    # 値を書き出す
    write_dashboard_data(dashboard_data, sheet)


if __name__  == '__main__':
    main()
