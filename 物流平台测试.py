import time
import queue
from typing import List

import pandas
import requests
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from settings import Settings

config = Settings


class DataModel:
    def __init__(
            self,
            waybill_number: str,
            to: str = None,
            latest_status: str = None,
            days_after_order: int = None,
            provider_name: str = None,
            provider_service_type: str = None,
            events: str = None

    ):
        self.waybill_number = waybill_number
        self.to = to
        self.latest_status = latest_status
        self.days_after_order = days_after_order
        self.provider_name = provider_name
        self.provider_service_type = provider_service_type
        self.events = events


class StartTask:
    def __init__(
            self,
            settings,
            proxies: dict = None
    ):
        self.settings = settings
        self.proxies = proxies

    def get_cookies(self, order_num):
        chrome = Service(self.settings.chrome_path)
        driver = webdriver.Chrome(service=chrome)

        driver.get(f'https://t.17track.net/zh-cn#nums={order_num}')
        time.sleep(20)
        cookie_list = driver.get_cookies()
        driver.close()
        if not cookie_list:
            return None
        for _ in cookie_list:
            name = _.get('name')
            value = _.get('value')
            if name == 'Last-Event-ID':
                return value

    def read_excel(self):
        df = pandas.read_excel(self.settings.order_excel_path)
        order_list = df.values
        cut_num = (len(order_list) // 40) + 1
        for i in range(cut_num):
            cut_list = list()
            for _ in order_list[i * 40: (i + 1) * 40]:
                num = _[0]
                cut_list.append(num)
            order_q.put(cut_list)

        return order_list[0][0]

    def request_order_info(self, order_num_list, cookies):
        data = list()
        for order_num in order_num_list:
            data.append(
                {
                    "fc": 0,
                    "num": f"{order_num}",
                    "sc": 0

                }
            )

        request_data = {
            "data": data,
            "guid": "",
            "timeZoneOffset": -480
        }
        cookies = {
            "Last-Event-ID": f"{cookies}",
            "country": "CN"
        }
        url = "https://t.17track.net/track/restapi"
        headers = {
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Content-Length": "85",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Host": "t.17track.net",
            "Origin": "https://t.17track.net",
            "Referer": "https://t.17track.net/zh-cn",
            "Sec-Ch-Ua": "\"Not.A/Brand\";v=\"8\", \"Chromium\";v=\"114\", \"Google Chrome\";v=\"114\"",
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": "\"Windows\"",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/114.0.0.0 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest"
        }
        response = requests.post(url, headers=headers, cookies=cookies, json=request_data, proxies=self.proxies).json()
        res = self.parse_response(response)
        return res

    @staticmethod
    def parse_response(json_data):
        if json_data.get('meta').get('code') == -8:
            raise PermissionError('cookies失效或者出现了验证码未通过！')
        return_data = list()
        shipments = json_data.get('shipments')
        for ship in shipments:
            prior_status = ship.get('prior_status')
            number = ship.get('number')
            carrier = ship.get('carrier')
            res = DataModel(waybill_number=number)
            res.provider_name = config.post_dict.get(str(carrier))
            if prior_status == 'NotFound':
                return_data.append(res)
                continue

            shipment: dict = ship.get('shipment', {})
            sender_country = config.country_dict.get(shipment.get('shipping_info').
                                                     get('shipper_address').get('country'))
            recipient_country = config.country_dict.get(shipment.get('shipping_info').
                                                        get('recipient_address').get('country'))

            to = f'{sender_country} -> {recipient_country}'
            res.to = to
            time_metrics = shipment.get('time_metrics').get('days_after_order')
            latest_status = config.package_info.get(shipment.get('latest_status').get('status'))
            if not latest_status:
                latest_status = config.package_info.get(shipment.get('latest_status').get('sub_status'))
                try:
                    latest_status = latest_status % config.post_dict.get(str(carrier))

                except TypeError:
                    latest_status = '正在等待揽收'
            res.latest_status = f'({time_metrics}天)--{latest_status}'
            events = shipment.get('tracking').get('providers')[0].get('events')
            event_list = list()
            for event in events:
                # time_info = time.strptime(event.get('time_iso')[:-6], '%Y-%m-%dT%H:%M:%S')
                event_str = f'{event.get("time_iso").replace("T", " ")[:-9]}  ' \
                            f'{event.get("location")}  ' \
                            f'{event.get("description")}'
                event_list.append(event_str)

            event = '\n'.join(event_list)
            res.events = event
            return_data.append(res)
        return return_data

    def write_to_excel(self, save_list: List[DataModel]):
        order_number = list()
        provider = list()
        status = list()
        country = list()
        event = list()
        for save in save_list:
            order_number.append(save.waybill_number)
            provider.append(save.provider_name)
            status.append(save.latest_status)
            country.append(save.to)
            event.append(save.events)
        headers = {
            '单号': order_number,
            '物流公司': provider,
            '包裹状态': status,
            '国家': country,
            '在途信息': event
        }
        fwrite = pandas.DataFrame(headers)
        fwrite.to_excel(f'./{config.res_excel_path}{config.res_excel_name}', index=False)

    def start(self):
        one_order_num = self.read_excel()
        cookie = self.get_cookies(one_order_num)
        print(cookie)
        if not cookie:
            raise PermissionError('检查是否出现验证码')
        save_list = list()
        while True:
            if order_q.empty():
                break
            res = self.request_order_info(order_q.get(), cookie)
            save_list += res

        self.write_to_excel(save_list)
        print('执行完成！')


if __name__ == '__main__':
    order_q = queue.Queue()
    StartTask(config).start()

# 很抱歉，系统检测到您的IP或者网段访问太频繁
