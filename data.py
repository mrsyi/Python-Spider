# -*- coding:UTF-8 -*-
import requests
from urllib.parse import urlencode

if __name__ == '__main__':
    headers = {
    'Host':'sdhq.hust.edu.cn',
    'Referer':'http://sdhq.hust.edu.cn/ICBS/hust/html/recently.html',
    'User-Agent':'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.119 Mobile Safari/537.36',
    'X-Requested-With':'XMLHttpRequest',
    }
    def get_page(page):
        params = {
            'AmMeter_ID':'ammeter',
            'startDate':'2019%2F2%2F3',
            'endDate':'2019%2F3%2F5',
            'page':page
            }
        target = 'http://sdhq.hust.edu.cn/ICBS/PurchaseWebService.asmx/getMeterDayValue?' + urlencode(params)
        try:
            response = requests.get(target,headers = headers)
            if response.status_code == 200:
                print(response.text)
        except requests.ConnectionError as e:
                    print('Error',e.args)
