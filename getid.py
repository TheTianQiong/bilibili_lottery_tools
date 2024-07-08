# 网络工具 爬取数据
import datetime
import json
import openpyxl
import requests
import os
from colorama import Fore


class id_operation:
    def __init__(self, id: str, cookie: str):
        """

        :param id: 是每个动态的id，通常可以在动态页面的url中找到
        :param cookie: 是用户登录后请求内必有内容，若缺省则会导致获取不到内容。一般从开发者工具里请求的cookie中直接复制
        :return:
        """
        global end_message
        offset = ""  # 初始为空表示第一页
        page = 1
        lst = []
        url = "https://api.bilibili.com/x/polymer/web-dynamic/v1/detail/reaction"
        headers = {
            "Cookie": cookie,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
        }
        params = {"id": id, "offset": offset}
        begin_message,plus_value=1,0
        count=0
        self.wb_new()
        while True:
            r = requests.get(url, params=params, headers=headers).json()
            total = json.loads(json.dumps(r))["data"]["total"]
            page_total = total // 20 if total % 20 == 0 else total // 20 + 1
            print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 正在获取第{page}/{page_total}页的内容中……",
                  end="")
            dict_r = json.loads(json.dumps(r))
            lst = lst + dict_r["data"]["items"]
            plus_value = plus_value+len(dict_r["data"]["items"])
            count = count + len(dict_r["data"]["items"])
            print(f"已合并")
            if not dict_r["data"]['has_more']:
                break
            if page%20==0: # 防止数据量较大，每到20页存档一次，并清空
                self.wb_append(lst,begin_message+1)
                lst=[]
                begin_message=begin_message+plus_value
                plus_value=0
            offset = dict_r["data"]["offset"]
            params = {"id": id, "offset": offset}
            page += 1
        print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 共计获取到{count}条数据，给定数据量为{total}")
        if count != total:
            print(Fore.RED+"注意：这里获取到的数据量和给定的数据量不一致，可能爬取的是非本UP主的且数据量较大的动态，\n若想要完整的获取，请登录本人的账号再试；也可能是存在新的数据而导致的"+ Fore.RESET)
        self.wb_append(lst, begin_message + 1)
        print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 数据处理完成")

    def wb_new(self):
        print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 新建getid.xlsx文件……", end="")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        ws['A1'], ws['B1'], ws['C1'], ws['D1'] = "Name", "Mid", "Action", "Face"
        try:
            wb.save('getid.xlsx')
        except:
            print(Fore.RED+"此处无法新建文档，可能是文件被占用，请关闭可能的程序" + Fore.RESET)
        print("完成")
    def wb_append(self,lst,start):
        print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 打开getid.xlsx文件")
        wb = openpyxl.load_workbook('getid.xlsx')
        ws = wb.active
        w=start
        print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 正在存入目前内容……", end="")
        for i in lst:
            ws["A" + str(w)], ws["B" + str(w)], ws["C" + str(w)], ws["D" + str(w)] = i["name"], i["mid"], i["action"], i["face"]
            w += 1
        wb.save('getid.xlsx')
        print("完成")




if __name__ == '__main__':
    import keyboard
    id = str(input("请输入要获取的动态id号："))
    cookie = str(input("请输入cookie值："))
    id_operation(id, cookie)
    os.system('pause')
