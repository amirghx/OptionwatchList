import pandas as pd
import requests
import json
import xlwings as xw
import time
List_one = []
List_two = []


def shifter(List, tag):
    final = [[tag]]
    for item in List:
        final.append([item])
    return final


def get_ask_price(data):
    return data.get("BidAskFirstRow").get("BestBuyPrice")


def get_ask_count(data):
    return data.get("BidAskFirstRow").get("BestBuyQuantity")


def get_bid_price(data):
    return data.get("BidAskFirstRow").get("BestSellPrice")


def get_bid_count(data):
    return data.get("BidAskFirstRow").get("BestSellQuantity")


def get_Last_traded_price(data):
    return data.get("LastTradedPrice")


def get_symbol_name(data):
    return data.get("SymbolFa")


def get_val(data):
    return data.get("TotalNumberOfSharesTraded")


def send_request(isin):
    response = requests.get("http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27" + isin + "%27")
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]
    return x


Main_df = pd.read_excel('Data/tradeOption.xls', header=3)

Name = Main_df['Unnamed: 0'].values.tolist()
ISIN_Code = Main_df['Unnamed: 1'].values.tolist()

temp = {'اسم اختیار': Name, 'کد ISIN': ISIN_Code}
df = pd.DataFrame(temp)
df.to_excel('data/source.xlsx', index=False)
check_List = []
for isin in ISIN_Code:
    data = send_request(isin)
    List_one.append(get_val(data))
    check_List.append('0')
while True:
    Name = []
    last_traded = []
    sell_price = []
    buy_price = []
    sell_count = []
    buy_count = []

    for isin in ISIN_Code:
        data = send_request(isin)
        Name.append(get_symbol_name(data))
        last_traded.append(get_Last_traded_price(data))
        sell_price.append(get_bid_price(data))
        buy_price.append(get_ask_price(data))
        sell_count.append(get_bid_count(data))
        buy_count.append(get_ask_count(data))
        List_two.append(get_val(data))

    for i in range(len(List_one)):
        if (List_two[i] - List_one[i]) >= 1000:
            check_List[i] = '1'
            List_one[i] = List_two[i]
        else:
            check_List[i] = '0'

    wb = xw.Book('data/watchList.xlsx')
    worksheet = wb.sheets('Sheet1')
    worksheet.range('A1').value = shifter(Name, 'نام')
    worksheet.range('B1').value = shifter(last_traded, 'قیمت اخرین معامله')
    worksheet.range('C1').value = shifter(sell_price, 'بهترین فروش')
    worksheet.range('D1').value = shifter(buy_price, 'بهترین خرید')
    worksheet.range('E1').value = shifter(sell_count, 'تعداد فروش')
    worksheet.range('F1').value = shifter(buy_count, 'تعداد خرید')
    worksheet.range('G1').value = shifter(check_List, 'وضعیت')
    print('done')
    time.sleep(120)

