# @Zak https://github.com/oooooranger
# This program can get Balance Sheet, Profit Statement and Cash Flow of Chinese listed companies by finance.sina.
# You need to prepare an Excel document with companies' code.
import requests
import pandas as pd
import time


headers = {
    "user-agent": "Mozilla/5.0 (windows NT 10.0; WOW 64) "
                  "AppleWebKit/537.36 (KHTML, Like Gecko) Chrome/75.0.3770.100 Safari 537.36"
}


def get_stocklist(file_path_1):
    df = pd.read_excel(file_path_1)
    stocklist = list(df['code'])
    return stocklist


def get_balancesheet(stock_id_list):
    for i in stock_id_list:
        url = 'https://money.finance.sina.com.cn/corp/go.php/vDOWN_BalanceSheet/displaytype/4/stockid/%s/'\
              'ctrl/all.phtml' % str(i)
        print(url)
        time.sleep(1)
        try:
            r = requests.get(url, allow_redirects=True)
            with open(r"E:\data\fin_tab\BalanceSheet\%s.xls" % str(i), "wb") as f:
                f.write(r.content)
            print(str(i) + 'Balance Sheet succeed!')
        except Exception as e:
            print(e)
            print(str(i) + 'BalanceSheet'+'失败')


def get_profitstatement(stock_id_list):
    for i in stock_id_list:
        url = 'https://money.finance.sina.com.cn/corp/go.php/vDOWN_ProfitStatement/displaytype/4/stockid/%s/'\
              'ctrl/all.phtml' % str(i)
        print(url)
        time.sleep(1)
        try:
            r = requests.get(url, allow_redirects=True)
            with open(r"E:\data\fin_tab\ProfitStatement\%s.xls" % str(i), "wb") as f:
                f.write(r.content)
            print(str(i) + 'Profit Statement succeed!')
        except Exception as e:
            print(e)
            print(str(i) + 'Profit Statement'+'失败')


def get_cashflow(stock_id_list):
    for i in stock_id_list:
        url = 'https://money.finance.sina.com.cn/corp/go.php/vDOWN_CashFlow/displaytype/4/stockid/%s/'\
              'ctrl/all.phtml' % str(i)
        print(url)
        time.sleep(1)
        try:
            r = requests.get(url, allow_redirects=True)
            with open(r"E:\data\fin_tab\CashFlow\%s.xls" % str(i), "wb") as f:
                f.write(r.content)
            print(str(i) + 'Cash Flow succeed!')
        except Exception as e:
            print(e)
            print(str(i) + 'Cash Flow'+'失败')


if __name__ == '__main__':
    file_path = r"E:\data\fin_tab\numberlist.xlsx"
    stk_list = get_stocklist(file_path_1=file_path)
    # get_balancesheet(stock_id_list=stk_list)
    # get_profitstatement(stock_id_list=stk_list)
    get_cashflow(stock_id_list=stk_list)
