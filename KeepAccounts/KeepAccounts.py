from dateutil.relativedelta import relativedelta
import PySimpleGUI as sg
import os, json, datetime, subprocess

class DATEedit:
    def __init__(self):
        ''' プログラム実行日の日付を取得し、その一か月前と一か月後を計算 '''
        self.registration_date = datetime.date.today()
        self.next_month = self.registration_date + relativedelta(months = 1)
        self.last_month = self.registration_date - relativedelta(months = 1)
        
    def date_for_HTML(self, date):
        ''' 日付をHTMLタグで囲む '''
        return '<time datetime="{0}">{1}</time>'.format(date.replace('/', '-'), date)

class EntryClerk(DATEedit):
    
    def __init__(self, purchase_date, trade_name, commodity_price, points_used, payment_method):
        super().__init__()
        
        self.purchase_date = purchase_date
        self.trade_name = trade_name
        self.commodity_price = commodity_price
        self.points_used = points_used
        self.payment_method = payment_method
        self.amountPaid_by_CreditCard = 0 #クレジットカードでの支払い金額
        self.amountPaid_by_after = 0 # 後払いでの支払い金額
        self.MonthlyTotal = (0, 0) # 今月の支払い金額の合計(クレカ, 後払い)
        
        self.TOTALIZATION_FILE = "totalization.dat" # 集計記録ファイル
        
        # 集計記録ファイルが存在すれば中のデータを取り出し、ファイルが存在しなければ作成
        if os.path.exists(self.TOTALIZATION_FILE):
            with open(self.TOTALIZATION_FILE, "r") as open_file:
                self.totalization_data = json.load(open_file) 
        else:
            self.totalization_data = {"CreditCard": 0, "after": 0, "update_date": self.registration_date.strftime("%Y-%m-%d")}
    
        # 今更新する帳簿の対象期間を判定
        if self.registration_date.timetuple()[2] >= 21:
            self.start_RecodingPeriod = self.registration_date.strftime("%Y/%m/21")
            self.end_RecodingPeriod = self.next_month.strftime("%Y/%m/20")
        else:
            self.start_RecodingPeriod = self.last_month.strftime("%Y/%m/21")
            self.end_RecodingPeriod = self.registration_date.strftime("%Y/%m/20")

        self.RecodingPeriod_forTITLE = self.start_RecodingPeriod.replace("/", "") + "-" + self.end_RecodingPeriod.replace("/", "")
        self.RecodingPeriod_forHTML = self.date_for_HTML(self.start_RecodingPeriod) + "～" + self.date_for_HTML(self.end_RecodingPeriod)
        
        self.ACOUNTBOOK_FILE = "AccountBook_{0}.html".format(self.RecodingPeriod_forTITLE) # 今回記録する帳簿ファイル名
        self.result_message = "" # 帳簿作成後に表示するメッセージ
        self.make_AccountBook() # 帳簿を作成
        
    def clear(self):
        ''' 集計記録ファイルのリセットを行う '''
        self.totalization_data["CreditCard"] = 0
        self.totalization_data["after"] = 0
        self.totalization_data["update_date"] = self.registration_date.strftime("%Y-%m-%d")
        with open(self.TOTALIZATION_FILE, mode="w") as open_file:
            json.dump(self.totalization_data, open_file)
    
    def judge_and_Clear(self):
        ''' 集計記録ファイルをリセットすべきか判定し、リセットを実行'''
        # 最終更新日を取得
        final_update = self.totalization_data["update_date"]
        final_update = datetime.date.fromisoformat(final_update)
        # 帳簿の対象期間の開始日を取得
        start = self.start_RecodingPeriod.replace("/", "-")
        start = datetime.date.fromisoformat(start)
        # 最終更新日が今更新する帳簿の対象期間より前なら、集計データを初期化
        if final_update < start: self.clear()
    
    def calc_amountPaid(self):
        ''' 今回の支払い金額を計算 '''
        amountPaid = int(self.commodity_price) - int(self.points_used) # 支払い金額
        return amountPaid
    
    def amountPaid_by_Method(self):
        ''' 支払い方法ごとの今回の利用額を計算 '''
        if self.payment_method == 'クレジットカード':
            self.amountPaid_by_CreditCard = self.calc_amountPaid()
        if self.payment_method == 'メルペイスマート払い(後払い)':
            self.amountPaid_by_after = self.calc_amountPaid()
        return (self.amountPaid_by_CreditCard, self.amountPaid_by_after)
    
    def write_TotalFile(self):
        ''' 今月の利用額の合計を求め、集計記録を更新'''
        self.amountPaid_by_Method()
        self.judge_and_Clear()
        self.totalization_data["CreditCard"] += self.amountPaid_by_CreditCard
        self.totalization_data["after"] += self.amountPaid_by_after
        self.totalization_data["update_date"] = self.registration_date.strftime("%Y-%m-%d")
        with open(self.TOTALIZATION_FILE, mode="w") as open_file:
            json.dump(self.totalization_data, open_file)
        total_CreditCard = self.totalization_data["CreditCard"] 
        total_after = self.totalization_data["after"] 
        self.MonthlyTotal = (total_CreditCard, total_after)
        return self.MonthlyTotal
    
    def make_HTMLpart(self):
        ''' 帳簿のうち、今回の利用明細を記述する行のHTML形式テキストを作成 '''
        self.MonthlyTotal = self.write_TotalFile()
        new_data = """\
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td align="right">{2}</td>
                <td align="right">{3}</td>
                <td align="right">{4}</td>
                <td align="right">{5}</td>
                <td align="right">{6}</td>
                <td align="right">{7}</td>
            </tr>
""".format(self.date_for_HTML(self.purchase_date), self.trade_name, self.commodity_price, self.points_used, self.amountPaid_by_CreditCard, self.amountPaid_by_after, self.MonthlyTotal[0], self.MonthlyTotal[1])
        return new_data
    
    def make_HTMLfull(self):
        ''' 帳簿全体のHTML形式テキストを作成'''
        new_data = self.make_HTMLpart()
        full_data = """<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>教材購入帳簿</title>
</head>
<body>
    <table border ="1">
        <caption>{0}の購入状況</caption>
        <thead>
            <tr>
                <th>購入日</th>
                <th>商品題名</th>
                <th width="100">商品代金</th>
                <th width="100">ポイント<br>利用額</th>
                <th width="100">クレカでの<br>支払い代金</th>
                <th width="100">後払いでの<br>支払い代金</th>
                <th width="100">今期の<br>クレカでの<br>支払い総額</th>
                <th width="100">来月末までの<br>後払い総額</th>
            </tr>
        </thead>
        <tbody>
{1}
        </tbody>
    </table>
</body>
</html>""".format(self.RecodingPeriod_forHTML, new_data)
        return full_data
    
    def make_AccountBook(self):
        ''' 帳簿ファイルを作成 '''
        # 既存の帳簿ファイルが存在しなければ新規作成する
        if not os.path.exists(self.ACOUNTBOOK_FILE):
            with open(self.ACOUNTBOOK_FILE, mode="w", encoding="utf-8") as open_file:
                open_file.write(self.make_HTMLfull())
            result_message = "新しい帳簿を作成しました。\n"
        # 対象期間の帳簿ファイルが存在すれば、追記する
        else:
            with open(self.ACOUNTBOOK_FILE, mode="r", encoding="utf-8") as open_file:
                fileData_temp = open_file.readlines()
            fileData_temp.insert(-4, self.make_HTMLpart())
            with open(self.ACOUNTBOOK_FILE, mode="w", encoding="utf-8") as open_file:
                open_file.writelines(fileData_temp)
            self.result_message = '帳簿に追加しました。\n'
        self.result_message += "今月のクレジットカード使用総額は{0}円です。\n".format(self.MonthlyTotal[0])
        print(self.result_message)

sg.theme('Dark Blue 3')

inputForm_layout = [
    [sg.Text('教材購入帳簿\n')],
    [sg.Text('購入日', size=(15, 1)), sg.InputText("{0}".format(datetime.date.today().strftime("%Y/%m/%d")), key="purchase_date")],
    [sg.Text('商品題名', size=(15, 1)), sg.InputText('', key="trade_name")],
    [sg.Text('商品代金', size=(15, 1)), sg.InputText('', key="commodity_price")],
    [sg.Text('ポイント利用額', size=(15, 1)), sg.InputText('', key="points_used")],
    [sg.Text('支払い方法', size=(15, 1)), sg.InputCombo(('クレジットカード', 'メルペイスマート払い(後払い)'), key="payment_method", size=(30, 1))],
    [sg.Text('')],
    [sg.Submit(button_text='確定')]
]
input_form = sg.Window('購入情報を入力', inputForm_layout)

while True:
    event, values = input_form.read()
    if event is None:
        print('exit')
        break
    if event == '確定':        
        AddItem = EntryClerk(values["purchase_date"], values["trade_name"], values["commodity_price"], values["points_used"], values["payment_method"])
        show_message = AddItem.result_message
        input_form.close()

ResultPage_layout = [
    [sg.Text(show_message)],
    [sg.Submit(button_text='更新した帳簿を表示')],
]
Result_page = sg.Window('帳簿の更新', ResultPage_layout)

while True:
    event, values = Result_page.read()
    if event is None:
        print('exit')
        break
    if event == '更新した帳簿を表示':
        subprocess.Popen([r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', AddItem.ACOUNTBOOK_FILE])
        break
Result_page.close()