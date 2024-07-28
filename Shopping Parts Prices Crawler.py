# address of "Gaming Parts List.xlsx": "D:\Gaming Parts List.xlsx"

import sys
import time
import random
import requests
import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import Font

amazon_url = [
    "https://www.amazon.ca/Blue-Solid-State-Drive-1TB/dp/B073SBQMCX",
    "https://www.amazon.ca/Seagate-IronWolf-5900RPM-Internal-3-5-Inch/dp/B07H289S79",
    "https://www.amazon.ca/Samsonite-89431-1041-Backpack-15-6-Inch-International/dp/B072KV8QGR",
    "https://www.amazon.ca/X-Rite-EODIS3-i1Display-Pro/dp/B0055MBQOW",
    "https://www.amazon.ca/Thrustmaster-T300-Racing-Wheel-English/dp/B00O8B7D02",
    "https://www.amazon.ca/ROG-Sheath-Gaming-Mouse-Pad/dp/B01G5ATZAE"
]

newegg_url = [
    "https://www.newegg.ca/western-digital-blue-1tb/p/N82E16820250088",
    "https://www.newegg.ca/seagate-ironwolf-st4000vn008-4tb/p/N82E16822179005"
]

canada_computers_url = [
    "https://www.canadacomputers.com/product_info.php?cPath=179_1927_1928&item_id=114570",
    "https://www.canadacomputers.com/product_info.php?cPath=15_1086_210&item_id=100441",
    "",
    "",
    "https://www.canadacomputers.com/product_info.php?cPath=13_1864_1866&item_id=083403",
    ""
]

best_buy_url = [
    "https://www.bestbuy.ca/en-ca/product/western-digital-blue-1tb-sata-internal-solid-state-drive/10906158",
    "https://www.bestbuy.ca/en-ca/product/seagate-ironwolf-4tb-3-5-5900-rpm-desktop-nas-internal-hard-drive-st4000vna08/11617909"
]

if len(sys.argv) < 2:
    test = True
else:
    excel = sys.argv[1]
    test = False

red_color = Font(color="FF0000")
nattempt = 5

user_agents = [
    # "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36",
    # "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36",
    # "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36",
    # "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36",
    # "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
]


print("The crawler is working...")

def checkAmazonPrice(url: str, user_agent: str) -> float:
    if url == "":
        return -1
    else:
        for i in range(nattempt):
            headers = {"User-Agent" : user_agent}
            try:
                page_html = requests.get(url,headers=headers)
                break
            except:
                print("Request timed out.")
                time.sleep(random.uniform(5,10))
        else:
            return -1
        page = BeautifulSoup(page_html.content,"html5lib")
        find_result = page.select_one(".a-offscreen")
        if find_result != None:
            price_text = find_result.get_text("", True)
            return float(price_text[1:])
    return -1

def checkCanadaComputersPrice(url: str, user_agent: str) -> float:
    if url == "":
        return -1
    else:
        for i in range(nattempt):
            headers = {"User-Agent" : user_agent}
            try:
                page_html = requests.get(url,headers=headers)
                break
            except:
                print("Request timed out.")
                time.sleep(random.uniform(5,10))
        else:
            return -1
        page = BeautifulSoup(page_html.content,"html.parser")
        if page.select_one(".order-md-1 span strong") != None:
            price_text = page.select_one(".order-md-1 span strong").get_text()
            if price_text == "Price too low to show!":
                price_text = page.select_one(".ordder-4 span").get_text()
                return float(price_text[1:price_text.find("&")])
            else:
                return float(price_text[1:])
    return -1

while True:
    user_agent = random.choice(user_agents)
    print(user_agent)
    
    if not test:
        pc_part_list = openpyxl.load_workbook(excel)
        sheet = pc_part_list[pc_part_list.sheetnames[0]]
    
        price_change = False
   
    amazon_price = []
    canada_computers_price = []
    min_price = []

    if test:
        print("Amazon:")
    for i in range(len(amazon_url)):
        amazon_price.append(checkAmazonPrice(amazon_url[i],user_agent))
        if test:
            print(amazon_price[-1])
        time.sleep(random.uniform(5,10))
    if test:
        print("Canada Computers:")
    for i in range(len(canada_computers_url)):
        canada_computers_price.append(checkCanadaComputersPrice(canada_computers_url[i],user_agent))
        if test:
            print(canada_computers_price[-1])
        time.sleep(random.uniform(5,10))

    if not test:
        for price in amazon_price:
            min_price.append(price)
        for i in range(len(canada_computers_price)):
            if canada_computers_price[i] == -2:
                min_price[i] = canada_computers_price[i]
            elif min_price[i] == -1:
                min_price[i] = canada_computers_price[i]
            elif canada_computers_price[i] < min_price[i] and canada_computers_price[i] != -1:
                min_price[i] = canada_computers_price[i]

        for i in range(len(min_price)):
            if test:
                print(sheet["D"+str(i+2)].value)
            if min_price[i] == -2:
                continue
            elif min_price[i] == -1:
                if sheet["D"+str(i+2)].value != "N/A":
                    sheet["D"+str(i+2)].value = "N/A"
                    sheet["D"+str(i+2)].font = red_color
                    price_change = True
            elif sheet["D"+str(i+2)].value == "N/A":
                sheet["D"+str(i+2)].value = min_price[i]
                sheet["D"+str(i+2)].font = red_color
                if sheet["C"+str(i+2)].value == "N/A" or min_price[i] < sheet["C"+str(i+2)].value:
                    sheet["C"+str(i+2)].value = min_price[i]
                    sheet["C"+str(i+2)].font = red_color
                price_change = True
            elif min_price[i] != sheet["D"+str(i+2)].value:
                sheet["D"+str(i+2)].value = min_price[i]
                sheet["D"+str(i+2)].font = red_color
                if sheet["C"+str(i+2)].value == "N/A" or min_price[i] < sheet["C"+str(i+2)].value:
                    sheet["C"+str(i+2)].value = min_price[i]
                    sheet["C"+str(i+2)].font = red_color
                price_change = True
        if price_change:
            try:
                pc_part_list.save(excel)
                print("Prices have changed at {}.".format(time.strftime("%H:%M:%S on %Y/%m/%d")))
                price_change = False
            except:
                print("Cannot save the Excel document.")

    if not test:
        time.sleep(random.uniform(300,600))
