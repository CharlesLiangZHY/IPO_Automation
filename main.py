import os
import re
import copy

from datetime import datetime, timedelta

import numpy as np
import pandas as pd 
import joblib

import xlsxwriter
from pypinyin import lazy_pinyin



import smtplib
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.mime.text import MIMEText 
from email.mime.application import MIMEApplication

from HTMLTable import HTMLTable

from corner_case import lower_hold_special_case, lottery_date_special_case, entry_special_case



IPO_dict = {
    "ID"                          : "Wind代码", ### IPO ID
    "IPO_name"                    : "Wind名称", ### IPO Name
    "online"                      : "网上网下标识", ### online
    "announcement_date"           : "招股公告日", ### Announcement Day
    "inquiry_date"                : "初步询价起始日", ### Inquiry Day
    "offline_subscription_date"   : "网下申购起始日", ### Offline Subscription Start Date
    "offline_payment_date"        : "网下申购缴款日", ### Offline Subscription Payment Date
    "offering_date"               : "上市日", ### Offering Day
    "purchase_limit"              : "申购上限", ### Purchase Limit
    "offline_purchase_limit"      : "网下申购上限", ### Offline Purchase Limit
    "funding"                     : "预计募集资金", ### Funding
    "issued_share"                : "新股发行数量", ### Number of New Shares Issued
    "price"                       : "首发价格", ### Price
    "online_purchase_limit"       : "网上申购上限", ### Online Purchase Limit
    "online_subscription_date"     : "网上申购日", ### Online Subscription Date
    "online_payment_date"         : "网上缴款日", ### Online Payment Date
}


subscription_dict = {
    "ID"                          : "wind代码", ### IPO ID
    "allotment_subject_name"      : "配售对象名称", ### Allotment Subject Name
    "valid_subscription_amount"   : "有效报价的申购数量", ### Valid Subscription Amount
    "allotment_amount"            : "获配数量", ### Allotment Amount
    "lockup_period"               : "锁定期", ### Lockup Period
    "quote"                       : "申报价格", ### Quote
    "valid"                       : "是否有效报价",
    "subscription_amount"         : "申报数量", ### Subscription Amount
}

calendar_dict = {
    "material_submitting"         : "交材",
    "inquiry"                     : "报价",
    "subscription"                : "申购",
    "payment"                     : "缴款",
    "lottery"                     : "摇号",
    "offering"                    : "上市",
}


_MAINBOARD_ = 0 ### mainboard
_SMALLMEDIUMBOARD_ =1 ### small and medium board
_SNTINNOVATIONBOARD_ = 2 ### S&T innovation board
_SECONDBOARD_ = 3 ### second board

_PAST_ = 0
_TODAY_ = 1
_TOMORROW_ =2
_FUTURE_ = 3




def online(string):
    if re.search("网下", string) is None:
        return True
    else:
        return False

def valid(string):
    if string == "有效":
        return True
    else:
        return False



def parse_ID(string):
    global _MAINBOARD_, _SMALLMEDIUMBOARD_, _SNTINNOVATIONBOARD_, _SECONDBOARD_ ### read only, not allow to modify
    if string[:2] == "60": ### mainboard
        return _MAINBOARD_
    elif string[:2] == "00": ### small and medium board
        return _SMALLMEDIUMBOARD_
    elif string[:2] == "68": ### S&T innovation board
        return _SNTINNOVATIONBOARD_
    elif string[:2] == "30": ### second board
        return _SECONDBOARD_


def date_str(date):
    if isinstance(date, datetime):
        return str(date.year)+str(date.month).rjust(2,'0') +str(date.day).rjust(2,'0')
    elif isinstance(date, str):
        return date
    elif date is None:
        return date
    else: ### to handle the case that date type is np.float
        return str(int(date))

def next_workday(date, workday_sheet, data=True):
    next_workday = None
    if data:
        index = -1
        for i in workday_sheet.index:
            if workday_sheet.loc[i][0] == date:
                index = i+1
        if index < len(workday_sheet.index) and index != -1:
            next_workday = workday_sheet.loc[index][0]
        else:
            dateStr = date_str(date)
            date = datetime(int(dateStr[:4]), int(dateStr[4:6]), int(dateStr[6:8]))
            weekday = date.weekday() ### 0: Mon, ..., 6: Sun
            interval = 1
            if weekday >= 4:
                interval += 6-weekday

            next_workday = date + timedelta(days = interval)
    else:
        dateStr = date_str(date)
        date = datetime(int(dateStr[:4]), int(dateStr[4:6]), int(dateStr[6:8]))
        weekday = date.weekday() ### 0: Mon, ..., 6: Sun
        interval = 1
        if weekday >= 4:
            interval += 6-weekday

        '''
        Official holiday cases have not implemented!!!
        '''
        pass

        next_workday = date + timedelta(days = interval)

    return int(date_str(next_workday))

def date_type(date):
    global today, tomorrow ### read only, not allow to modify
    global _PAST_, _TODAY_, _TOMORROW_, _FUTURE_ ### read only, not allow to modify
    _today = int(date_str(today))
    _tomorrow = int(date_str(tomorrow))
    if isinstance(date, str):
        return -1
    elif date is None:
        return -1
    else:
        if int(date) == _today:
            return _TODAY_
        elif int(date) == _tomorrow:
            return _TOMORROW_
        elif int(date) < _today:
            return _PAST_
        elif int(date) > _tomorrow:
            return _FUTURE_




class IPO:
    def __init__(self, info_dict, subscription_sheet, workday_sheet, history=None):
        self.ID = info_dict["ID"]
        self.announcement_date = info_dict["announcement_date"]
        self.inquiry_date = info_dict["inquiry_date"]
        self.offline_subscription_date = info_dict["offline_subscription_date"]
        self.offline_payment_date = info_dict["offline_payment_date"]
        self.offering_date = info_dict["offering_date"]
        self.purchase_limit = info_dict["purchase_limit"]
        self.offline_purchase_limit = info_dict["offline_purchase_limit"]
        self.funding = info_dict["funding"]
        self.issued_share = info_dict["issued_share"]
        self.price = info_dict["price"]
        self.online_purchase_limit = info_dict["online_purchase_limit"]
        self.online_subscription_date = info_dict["online_subscription_date"]
        self.online_payment_date = info_dict["online_payment_date"]

        self.online = online(info_dict["online"]) ### True / False

        self.lottery_date = self.cal_lottery_date(workday_sheet)
        self.entry = -1 ### default -1 (which means it is unnecessary to consider); 0 entried; 1 partly entried; 2 not entried
        self.allotment_subjects = []

        global today, tomorrow ### read only, not allow to modify
        if date_str(self.offline_subscription_date) == date_str(today) or date_str(self.offline_subscription_date) == date_str(tomorrow):
            self.entry, self.allotment_subjects = entry_or_not(self.ID, subscription_sheet)
        elif date_str(self.offline_subscription_date) == date_str(tomorrow) or date_str(self.offering_date) == date_str(today):
            _, self.allotment_subjects = entry_or_not(self.ID, subscription_sheet)


        ### online and partly/not entried flags are mutually exclusive
        if self.ID in entry_special_case:
            self.IPO_name = entry_special_case[self.ID]
            if history is not None:
                history[self.ID] = self.IPO_name
        elif self.online:
            self.IPO_name = info_dict["IPO_name"] + "（网上）"
        elif self.entry == 1:
            self.IPO_name = info_dict["IPO_name"] + "（部分入围）"
        elif self.entry == 2:
            self.IPO_name = info_dict["IPO_name"] + "（未入围）"
        elif history is not None and history.__contains__(self.ID):
            self.IPO_name = history[self.ID]
        else:
            self.IPO_name = info_dict["IPO_name"]

        if history is not None and self.entry != 0:
            history[self.ID] = self.IPO_name ### this works in python, the history dictionary passed in will be updated



    def cal_lottery_date(self, workday_sheet):
        global _MAINBOARD_, _SMALLMEDIUMBOARD_, _SNTINNOVATIONBOARD_, _SECONDBOARD_ ### read only, not allow to modify
        board_type = parse_ID(self.ID)
        if self.ID in lottery_date_special_case:
            return lottery_date_special_case[self.ID]
        if board_type == _MAINBOARD_ or board_type == _SMALLMEDIUMBOARD_:
            return None
        if self.online:
            return None
        elif board_type == _SNTINNOVATIONBOARD_:
            return next_workday(self.offline_payment_date, workday_sheet)
        elif board_type == _SECONDBOARD_:
            return "10%比例限售锁定"


def pinyin(chinese_str):
    chars = lazy_pinyin(chinese_str)
    pinyin_str = ""
    for char in chars:
        pinyin_str += char
    return pinyin_str

def entry_or_not(IPO_ID, subscription_sheet, sort=True):
    entry = -1
    allotment_subjects = []
    valid_flag = 0 
    invalid_flag = 0

    global subscription_dict ### read only, not allow to modify

    for i in subscription_raw.index:
        info_dict = copy.deepcopy(subscription_dict)
        ID = subscription_raw.loc[i][info_dict["ID"]]
        if ID == IPO_ID:
            for key in info_dict.keys():
                if pd.isna(subscription_raw.loc[i][info_dict[key]]):
                    info_dict[key] = None
                else:
                    info_dict[key] = subscription_raw.loc[i][info_dict[key]]
            allotment_subjects.append(Allotment(info_dict))
            if valid(info_dict["valid"]):
                valid_flag += 1
            else:
                invalid_flag += 1
        else:
            continue

    if valid_flag > 0 and invalid_flag == 0:
        entry = 0 ### entried
    elif valid_flag > 0 and invalid_flag > 0:
        entry = 1 ### partly entried
    elif valid_flag == 0 and invalid_flag > 0:
        entry = 2 ### not entried


    if sort:
        allotment_subjects.sort(key=lambda allotment: pinyin(allotment.name))



    return entry, allotment_subjects



class Allotment():
    def __init__(self, info_dict):
        self.ID = info_dict["ID"]
        self.name = info_dict["allotment_subject_name"]
        self.valid_subscription_amount = info_dict["valid_subscription_amount"]
        self.allotment_amount = info_dict["allotment_amount"]
        self.lockup_period = info_dict["lockup_period"]
        self.quote = info_dict["quote"]
        self.valid = valid(info_dict["valid"])
        self.subscription_amount = info_dict["subscription_amount"]

class IPO_calendar():
    def __init__(self, IPO_raw, subscription_raw, workday_sheet, history = None):
        global today, tomorrow ### read only, not allow to modify
        self.today = date_str(today)
        self.tomorrow = date_str(tomorrow)
        self.IPO_sheet = IPO_raw
        self.subscription_sheet = subscription_raw
        self.workday_sheet = workday_sheet

        self.today_IPO = {"material_submitting":[], "inquiry":[], "subscription":[], "payment":[], "lottery":[], "offering":[]}
        self.tomorrow_IPO = {"inquiry":[], "subscription":[], "payment":[], "lottery":[], "offering":[]}

        for i in self.IPO_sheet.index:
            info_dict = copy.deepcopy(IPO_dict)
            for key in info_dict.keys():
                if pd.isna(self.IPO_sheet.loc[i][info_dict[key]]):
                    info_dict[key] = None
                else:
                    '''
                    The data types of raw sheet have not aligned.
                    Some date info is np.float64 type. Therefore, they need to be handle separately

                    '''
                    info_dict[key] = self.IPO_sheet.loc[i][info_dict[key]]

            ipo = IPO(info_dict, self.subscription_sheet, self.workday_sheet, history)
            self.parse_IPO_date(ipo)
            



    def parse_IPO_date(self, ipo):
        if date_str(ipo.announcement_date) == self.today:
            self.today_IPO["material_submitting"].append(ipo)
        if ipo.online:
            if date_str(ipo.online_subscription_date)  == self.today:
                self.today_IPO["subscription"].append(ipo)
            if date_str(ipo.online_payment_date) == self.today:
                self.today_IPO["payment"].append(ipo)
            if date_str(ipo.online_subscription_date) == self.tomorrow:
                self.tomorrow_IPO["subscription"].append(ipo)
            if date_str(ipo.online_payment_date) == self.tomorrow:
                self.tomorrow_IPO["payment"].append(ipo)
        else:
            if date_str(ipo.offline_subscription_date)  == self.today:
                self.today_IPO["subscription"].append(ipo)
            if date_str(ipo.offline_payment_date) == self.today:
                self.today_IPO["payment"].append(ipo)
            if date_str(ipo.offline_subscription_date) == self.tomorrow:
                self.tomorrow_IPO["subscription"].append(ipo)
            if date_str(ipo.offline_payment_date) == self.tomorrow:
                self.tomorrow_IPO["payment"].append(ipo)

            if date_str(ipo.inquiry_date) == self.today:
                self.today_IPO["inquiry"].append(ipo)
            if date_str(ipo.inquiry_date) == self.tomorrow:
                self.tomorrow_IPO["inquiry"].append(ipo)


        if date_str(ipo.lottery_date) == self.today:
            self.today_IPO["lottery"].append(ipo)
        if date_str(ipo.offering_date) == self.today:
            self.today_IPO["offering"].append(ipo)

        if date_str(ipo.lottery_date) == self.tomorrow:
            self.tomorrow_IPO["lottery"].append(ipo)
        if date_str(ipo.offering_date) == self.tomorrow:
            self.tomorrow_IPO["offering"].append(ipo)


class excel_writer:
    def __init__(self, filename):
        global today, tomorrow ### read only, not allow to modify
        self.today = date_str(today)
        self.tomorrow = date_str(tomorrow)
        self.workbook = xlsxwriter.Workbook(filename) 
        self.worksheets = {}

    def add_worksheet(self, sheet_name):
        if self.worksheets.__contains__(sheet_name):
            print("Error: The sheet has already existed.")
        else:
            self.worksheets[sheet_name] = [self.workbook.add_worksheet(sheet_name), 0] ### 0 is the current row for writing
        return self.worksheets[sheet_name]

    def save(self):
        self.workbook.close()

    def generate_IPO_calendar(self, data):
        global calendar_dict ### read only, not allow to modify

        html_tables = []

        calendar = self.add_worksheet("新股日历")
        title_format = self.workbook.add_format({
                                                 'bold':True, 
                                                 'align':'left', 
                                                })
        write(calendar, [0], ["新股日历"], [title_format])

        table = HTMLTable(caption='新股日历')
        table.caption.set_style({
            'color': '#000000',
            'font-size': '25px',
            'width': '500px',
        })

        html_tables.append(table.to_html())


        ### today calendar
        today_title_format = self.workbook.add_format({
                                                       'bold':True, 
                                                       'border':1, 
                                                       'align':'center', 
                                                       'fg_color':'#DC143C',
                                                     })

        
        table = HTMLTable(caption=self.today)
        table.caption.set_style({
            'color': '#000000',
            'font-size': '15px',
            'width': '500px',
            'background-color': '#DC143C',
        })
        
        
        calendar[0].merge_range(calendar[1],0,calendar[1],1, self.today, today_title_format)



        calendar[1] += 1
        calendar_cell_format = self.workbook.add_format({
                                                    'bold':True, 
                                                    'border':1, 
                                                    'align':'left', 
                                                  })

        table_content = []
        for key in data.today_IPO.keys():
            content = ""
            for ipo in data.today_IPO[key]:
                content += f" {ipo.IPO_name} "
            write(calendar,[0,1], [calendar_dict[key], content], [calendar_cell_format for i in range(2)])
            table_content.append((calendar_dict[key], content))

        table.append_data_rows(table_content)
        table.set_style({
          'color': '#000000',
          'border-collapse': 'collapse',
          'word-break': 'keep-all',
          'white-space': 'nowrap',
          'font-size': '15px',
          'width': '100%',
        })

        table.set_cell_style({
          'color': '#000000',
          'border-color': '#000000',
          'border-width': '1px',
          'border-style': 'solid',
          'padding': '5px',
        })


        html_tables.append(table.to_html())

        calendar[1] += 1 ### An empty line

        ### tomorrow calendar
        tomorrow_title_format = self.workbook.add_format({
                                                       'bold':True, 
                                                       'border':1, 
                                                       'align':'center', 
                                                       'fg_color':'#CD7F32',
                                                     })

        calendar[0].merge_range(calendar[1],0,calendar[1],1, self.tomorrow, tomorrow_title_format)
        
        table = HTMLTable(caption=self.tomorrow)
        table.caption.set_style({
            'color': '#000000',
            'font-size': '15px',
            'width': '500px',
            'background-color': '#CD7F32',
        })


        calendar[1] += 1
        table_content = []
        for key in data.tomorrow_IPO.keys():
            content = ""
            for ipo in data.tomorrow_IPO[key]:
                content += f" {ipo.IPO_name} "
            write(calendar,[0,1], [calendar_dict[key], content], [calendar_cell_format for i in range(2)])
            table_content.append((calendar_dict[key], content))

        table.append_data_rows(table_content)
        table.set_style({
          'color': '#000000',
          'border-collapse': 'collapse',
          'word-break': 'keep-all',
          'white-space': 'nowrap',
          'font-size': '15px',
          'width': '100%',
        })

        table.set_cell_style({
          'border-color': '#000000',
          'border-width': '1px',
          'border-style': 'solid',
          'padding': '5px',
        })



        html_tables.append(table.to_html())

        calendar[1] += 1 ### An empty line


        write(calendar,[0], ["详细信息:"], [title_format])

        table = HTMLTable(caption='详细信息')
        table.caption.set_style({
            'color': '#000000',
            'font-size': '15px',
            'width': '500px',
        })

        html_tables.append(table.to_html())

        offline = []
        online = []
        for key in data.today_IPO.keys():
            for ipo in data.today_IPO[key]:
                if not ipo.online and ipo not in offline:
                    offline.append(ipo)
                elif ipo.online and ipo not in online:
                    online.append(ipo)
        for key in data.tomorrow_IPO.keys():
            for ipo in data.tomorrow_IPO[key]:
                if not ipo.online and ipo not in offline:
                    offline.append(ipo)
                elif ipo.online and ipo not in online:
                    online.append(ipo)

        global _MAINBOARD_, _SMALLMEDIUMBOARD_, _SNTINNOVATIONBOARD_, _SECONDBOARD_ ### read only, not allow to modify
        global _PAST_, _TODAY_, _TOMORROW_, _FUTURE_ ### read only, not allow to modify

        if len(offline) > 0:
            table = HTMLTable(caption="网下")
            table.caption.set_style({
                'color': '#000',
                'font-size': '15px',
                'width': '100%',
            })

            ### offline
            table_content = []
            mark = []
            titles = ["代码", "简称", "股价测算", "底仓要求", "申报金额上限", "招股公告日", "初步询价起始日", "网下申购起始日", "网下申购缴款日", "网下摇号日", "上市日"]
            table_content.append(tuple(titles))


            calendar[0].set_column(0,len(titles),15)
            write(calendar, range(0,len(titles)), titles, [self.workbook.add_format({'border':1}) for i in range(len(titles))])
        
            for ipo in offline:
                contents = []
                cell_formats = []

                contents.append(ipo.ID)
                cell_formats.append(self.workbook.add_format({'border':1}))

                contents.append(ipo.IPO_name)
                cell_formats.append(self.workbook.add_format({'border':1}))

                price = None
                if ipo.price is not None:
                    price = round(ipo.price, 2)
                   
                else:
                    if parse_ID(ipo.ID) == _SNTINNOVATIONBOARD_ or parse_ID(ipo.ID) == _SECONDBOARD_:
                        pass
                    elif parse_ID(ipo.ID) == _MAINBOARD_ or parse_ID(ipo.ID) == _SMALLMEDIUMBOARD_:
                        if ipo.funding is None:
                            print(f"预计募集资金 of {ipo.ID} missed. The price can not be calculated. Please check the raw data!")
                            price = None
                        elif ipo.issued_share is None:
                            print(f"新股发行数量 of {ipo.ID} missed. The price can not be calculated. Please check the raw data!")
                            price = None
                        else:
                            price = round(ipo.funding / ipo.issued_share, 2)
                contents.append(price if price is None else '%.02f'%price)
                cell_formats.append(self.workbook.add_format({'border':1}))

                holder = None
                if parse_ID(ipo.ID) == _SNTINNOVATIONBOARD_ or parse_ID(ipo.ID) == _SECONDBOARD_ or ipo.ID in lower_hold_special_case:
                    holder = "6000"
                elif parse_ID(ipo.ID) == _MAINBOARD_ or parse_ID(ipo.ID) == _SMALLMEDIUMBOARD_:
                    holder = "1000/6000"
                contents.append(holder)
                cell_formats.append(self.workbook.add_format({'border':1}))

                application_limit = None
                if price is None:
                    pass
                else:
                    if ipo.offline_purchase_limit is None:
                        print(f"网下申购上限 of {ipo.ID} missed. 申购上限 can not be calculated. Please check the raw data!")

                    else: 
                        application_limit = price * ipo.offline_purchase_limit
                contents.append(application_limit if application_limit is None else '%.02f'%application_limit)
                cell_formats.append(self.workbook.add_format({'border':1}))

                announcement_date = ipo.announcement_date
                contents.append(announcement_date if announcement_date is None else int(announcement_date))
                if date_type(announcement_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1,  
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(announcement_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C','border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(announcement_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(announcement_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                inquiry_date = ipo.inquiry_date
                contents.append(inquiry_date if inquiry_date is None else int(inquiry_date))
                if date_type(inquiry_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(inquiry_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(inquiry_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(inquiry_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                offline_subscription_date = ipo.offline_subscription_date
                contents.append(offline_subscription_date if offline_subscription_date is None else int(offline_subscription_date))
                if date_type(offline_subscription_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0','border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(offline_subscription_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(offline_subscription_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(offline_subscription_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                offline_payment_date = ipo.offline_payment_date
                contents.append(offline_payment_date if offline_payment_date is None else int(offline_payment_date))
                if date_type(offline_payment_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(offline_payment_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(offline_payment_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(offline_payment_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                lottery_date = ipo.lottery_date
                contents.append(lottery_date if lottery_date is None or type(lottery_date) is str else int(lottery_date))
                if date_type(lottery_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(lottery_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(lottery_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(lottery_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                offering_date = ipo.offering_date
                contents.append(offering_date if offering_date is None else int(offering_date))
                if date_type(offering_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(offering_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(offering_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(offering_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                write(calendar, range(0,len(titles)), contents, cell_formats)
                table_content.append(contents)

            table.append_data_rows(table_content)

            table.set_style({
              'color': '#000000',
              'border-collapse': 'collapse',
              'word-break': 'keep-all',
              'white-space': 'nowrap',
              'font-size': '15px',
              'width': '100%',
            })

            table.set_cell_style({
              'border-color': '#000000',
              'border-width': '1px',
              'border-style': 'solid',
              'padding': '5px',
            })

            for cell in mark:
                table[cell[0][0]][cell[0][1]].set_style({
                    'background-color' : cell[1],
                })

            html_tables.append(table.to_html())

            calendar[1] += 1 ### An empty line
            calendar[1] += 1 ### An empty line

        if len(online) > 0:
            table = HTMLTable(caption="网上")
            table.caption.set_style({
                'color': '#000',
                'font-size': '15px',
                'width': '100%',
            })

            ### online
            table_content = []
            mark = []
            titles = ["代码", "简称", "股价", "网上申购上限(股)", "网上申购资金上限", "招股公告日", "网上申购起始日", "网上申购缴款日", "上市日"]
            write(calendar, range(0,len(titles)), titles, [self.workbook.add_format({'border':1}) for i in range(len(titles))])
            table_content.append(tuple(titles))

            for ipo in online:
                contents = []
                cell_formats = []

                contents.append(ipo.ID)
                cell_formats.append(self.workbook.add_format({'border':1}))

                contents.append(ipo.IPO_name)
                cell_formats.append(self.workbook.add_format({'border':1}))

                contents.append(ipo.price)
                cell_formats.append(self.workbook.add_format({'border':1}))

                contents.append(ipo.online_purchase_limit)
                cell_formats.append(self.workbook.add_format({'border':1}))

                if ipo.price is not None and ipo.online_purchase_limit is not None:
                    contents.append(ipo.price * ipo.online_purchase_limit)
                else:
                    contents.append(self.workbook.add_format({'border':1}))
                cell_formats.append(self.workbook.add_format({'border':1}))

                announcement_date = ipo.announcement_date
                contents.append(announcement_date if announcement_date is None else int(announcement_date))
                if date_type(announcement_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(announcement_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(announcement_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(announcement_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                online_subscription_date = ipo.online_subscription_date
                contents.append(online_subscription_date if online_subscription_date is None else int(online_subscription_date))
                if date_type(online_subscription_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(online_subscription_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(online_subscription_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(online_subscription_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                online_payment_date = ipo.online_payment_date
                contents.append(online_payment_date if online_payment_date is None else int(online_payment_date))
                if date_type(online_payment_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(online_payment_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(online_payment_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(online_payment_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))


                offering_date = ipo.offering_date
                contents.append(offering_date if offering_date is None else int(offering_date))
                if date_type(offering_date) == _PAST_:
                    past_format = self.workbook.add_format({
                                                           'fg_color':'#C0C0C0', 'border':1, 
                                                           })
                    cell_formats.append(past_format)
                    mark.append(((len(table_content), len(contents)-1), '#C0C0C0') )
                elif date_type(offering_date) == _TODAY_:
                    today_format = self.workbook.add_format({
                                                           'fg_color':'#DC143C', 'border':1, 
                                                           })
                    cell_formats.append(today_format)
                    mark.append(((len(table_content), len(contents)-1), '#DC143C') )
                elif date_type(offering_date) == _TOMORROW_:
                    tomorrow_format = self.workbook.add_format({
                                                           'fg_color':'#CD7F32', 'border':1, 
                                                           })
                    cell_formats.append(tomorrow_format)
                    mark.append(((len(table_content), len(contents)-1), '#CD7F32') )
                elif date_type(offering_date) == _FUTURE_:
                    cell_formats.append(self.workbook.add_format({'border':1}))
                else: ### -1 
                    cell_formats.append(self.workbook.add_format({'border':1}))

                write(calendar, range(0,len(titles)), contents, cell_formats)
                table_content.append(contents)

            table.append_data_rows(table_content)

            table.set_style({
              'color': '#000000',
              'border-collapse': 'collapse',
              'word-break': 'keep-all',
              'white-space': 'nowrap',
              'font-size': '15px',
              'width': '100%',
            })

            table.set_cell_style({
              'border-color': '#000000',
              'border-width': '1px',
              'border-style': 'solid',
              'padding': '5px',
            })

            for cell in mark:
                table[cell[0][0]][cell[0][1]].set_style({
                    'background-color' : cell[1],
                })

            html_tables.append(table.to_html())

        return html_tables

    def today_offering(self, data):

        html_tables = []

        table = HTMLTable(caption='今日上市')
        table.caption.set_style({
            'color': '#000000',
            'font-size': '25px',
            'width': '500px',
            'background-color':'#DC143C',
        })

        html_tables.append(table.to_html())

        calendar = self.add_worksheet("今日上市")

        title_format = self.workbook.add_format({
                                                 'bold':True, 
                                                 'align':'center', 
                                                })

        write(calendar, [0], ["今日上市"], [title_format])

        table_cell_format = self.workbook.add_format({
                                                      'border':1, 
                                                      'align':'center', 
                                                    })

        titles = ["配售对象名称", "有效报价的申购数量(万股)", "获配数量(股)", "锁定期(月)"]


        calendar[0].set_column(0,len(titles),30)
        ipos = data.today_IPO["offering"]

        if len(ipos) == 0:
            write(calendar, [0], ["无"], [title_format])
            table = HTMLTable(caption='无')
            table.caption.set_style({
                'color': '#000000',
                'font-size': '15px',
                'width': '100%',
            })

            html_tables.append(table.to_html())


        for ipo in ipos:
            table = HTMLTable(caption = ipo.ID + "   " + ipo.IPO_name)
            table.caption.set_style({
                'color': '#000000',
                'font-size': '25px',
                'width': '100%',
            })

            

            write(calendar, [0], [ipo.ID + "   " + ipo.IPO_name], [title_format])
            write(calendar, range(0,len(titles)), titles, [table_cell_format for i in range(len(titles))])
            
            table_content = []
            table_content.append(tuple(titles))

            for allotment in ipo.allotment_subjects:
                if not allotment.valid:
                    continue
                if parse_ID(ipo.ID) == _SECONDBOARD_:
                    contents = [allotment.name, '%.02f'%(allotment.valid_subscription_amount*0.1) if int(allotment.lockup_period) == 6 else '%.02f'%(allotment.valid_subscription_amount*0.9), '%.02f'%allotment.allotment_amount]
                else:
                    contents = [allotment.name, '%.02f'%allotment.valid_subscription_amount, '%.02f'%allotment.allotment_amount]

                lockup_period = allotment.lockup_period
                if lockup_period is not None and int(lockup_period) == 6:
                    contents.append(6)
                elif lockup_period is not None and int(lockup_period) == 0:
                    contents.append("-")
                else:
                    contents.append(None)




                write(calendar, range(0,len(titles)), contents, [table_cell_format for i in range(len(titles))])

                table_content.append(contents)

            table.append_data_rows(table_content)

            table.set_style({
              'color': '#000000',
              'border-collapse': 'collapse',
              'word-break': 'keep-all',
              'white-space': 'nowrap',
              'font-size': '15px',
              'width': '100%',
            })

            table.set_cell_style({
              'border-color': '#000000',
              'border-width': '1px',
              'border-style': 'solid',
              'padding': '5px',
            })

            html_tables.append(table.to_html())

            calendar[1] += 1 ### An empty line
            calendar[1] += 1 ### An empty line

        return html_tables


    def today_purchase(self, data):

        html_tables = []

        table = HTMLTable(caption='今日申购')
        table.caption.set_style({
            'color': '#000000',
            'font-size': '25px',
            'width': '500px',
            'background-color':'#DC143C',
        })

        html_tables.append(table.to_html())

        calendar = self.add_worksheet("今日申购")

        title_format = self.workbook.add_format({
                                                 'bold':True, 
                                                 'align':'center', 
                                                })

        write(calendar, [0], ["今日申购"], [title_format])

        table_cell_format = self.workbook.add_format({
                                                      'border':1, 
                                                      'align':'center', 
                                                    })

        titles = ["配售对象名称", "申报价格(元)", "是否有效报价", "申报数量(万股)"]
        

        calendar[0].set_column(0,len(titles),30)
        ipos = data.today_IPO["subscription"]

        if len(ipos) == 0:
            write(calendar, [0], ["无"], [title_format])
            table = HTMLTable(caption='无')
            table.caption.set_style({
                'color': '#000000',
                'font-size': '15px',
                'width': '100%',
            })

            html_tables.append(table.to_html())


        for ipo in ipos:
            if ipo.entry == 2:
                continue
            table = HTMLTable(caption = ipo.ID + "   " + ipo.IPO_name)
            table.caption.set_style({
                'color': '#000000',
                'font-size': '25px',
                'width': '100%',
            })



            write(calendar, [0], [ipo.ID + "   " + ipo.IPO_name], [title_format])
            write(calendar, range(0,len(titles)), titles, [table_cell_format for i in range(len(titles))])


            table_content = []
            table_content.append(tuple(titles))
            
            for allotment in ipo.allotment_subjects:
                if not allotment.valid:
                    continue

                contents = [allotment.name, allotment.quote, "是" if allotment.valid else "否", allotment.subscription_amount]
                write(calendar, range(0,len(titles)), contents, [table_cell_format for i in range(len(titles))])
                table_content.append(contents)

            table.append_data_rows(table_content)
            table.set_style({
              'color': '#000000',
              'border-collapse': 'collapse',
              'word-break': 'keep-all',
              'white-space': 'nowrap',
              'font-size': '15px',
              'width': '100%',
            })

            table.set_cell_style({
              'border-color': '#000000',
              'border-width': '1px',
              'border-style': 'solid',
              'padding': '5px',
            })

            html_tables.append(table.to_html())

            


            calendar[1] += 1 ### An empty line
            calendar[1] += 1 ### An empty line

        return html_tables

    def tomorrow_purchase(self, data):

        html_tables = []

        table = HTMLTable(caption='明日申购')
        table.caption.set_style({
            'color': '#000000',
            'font-size': '25px',
            'width': '500px',
            'background-color':'#CD7F32'
        })

        html_tables.append(table.to_html())


        calendar = self.add_worksheet("明日申购")

        title_format = self.workbook.add_format({
                                                 'bold':True, 
                                                 'align':'center', 
                                                })

        write(calendar, [0], ["明日申购"], [title_format])

        table_cell_format = self.workbook.add_format({
                                                      'border':1, 
                                                      'align':'center', 
                                                    })

        titles = ["配售对象名称", "申报价格(元)", "是否有效报价", "申报数量(万股)"]
        calendar[0].set_column(0,len(titles),30)
        ipos = data.tomorrow_IPO["subscription"]

        if len(ipos) == 0:
            write(calendar, [0], ["无"], [title_format])
            table = HTMLTable(caption='无')
            table.caption.set_style({
                'color': '#000000',
                'font-size': '15px',
                'width': '100%',
            })

            html_tables.append(table.to_html())


        for ipo in ipos:
            if ipo.entry == 2:
                continue
            table = HTMLTable(caption = ipo.ID + "   " + ipo.IPO_name)
            table.caption.set_style({
                'color': '#000000',
                'font-size': '25px',
                'width': '100%',
            })

            write(calendar, [0], [ipo.ID + "   " + ipo.IPO_name], [title_format])
            write(calendar, range(0,len(titles)), titles, [table_cell_format for i in range(len(titles))])
            
            table_content = []
            table_content.append(tuple(titles))

            for allotment in ipo.allotment_subjects:
                if not allotment.valid:
                    continue

                contents = [allotment.name, allotment.quote, "是" if allotment.valid else "否", allotment.subscription_amount]
                write(calendar, range(0,len(titles)), contents, [table_cell_format for i in range(len(titles))])
                table_content.append(contents)

            table.append_data_rows(table_content)
            table.set_style({
              'color': '#000000',
              'border-collapse': 'collapse',
              'word-break': 'keep-all',
              'white-space': 'nowrap',
              'font-size': '15px',
              'width': '100%',
            })

            table.set_cell_style({
              'border-color': '#000000',
              'border-width': '1px',
              'border-style': 'solid',
              'padding': '5px',
            })

            html_tables.append(table.to_html())

            calendar[1] += 1 ### An empty line
            calendar[1] += 1 ### An empty line

        return html_tables

def write(worksheet, cols, contents, cell_formats = None):
    if cell_formats is None:
        for i in range(len(cols)):
            worksheet[0].write(worksheet[1], cols[i], contents[i])
    else:
        for i in range(len(cols)):
            worksheet[0].write(worksheet[1], cols[i], contents[i], cell_formats[i])
    worksheet[1] += 1



class mail:

    def __init__(self, address, authorization_code, server, port):
        self.address = address
        self.authorization_code = authorization_code
        self.server = server
        self.port = port
        # self.SMTP_server = smtplib.SMTP_SSL(self.server, self.port)
        self.SMTP_server = smtplib.SMTP()
        self.SMTP_server.connect(self.server, self.port)
        self.SMTP_server.login(self.address, self.authorization_code)


    def send(self, mail_list, contents, attachment = None, today = date_str(datetime.today())):
        msg = MIMEMultipart()
        msg['From'] = formataddr(["LZY", self.address])
        msg['To'] = formataddr(["Test",",".join(mail_list)])
        print(",".join(mail_list))
        msg['Subject'] = date_str(today)+"_IPO_Info"

        if type(contents) is list:
            for content in contents:
                msg.attach(content)
        else:
            msg.attach(contents)

        if attachment is not None:
            file = MIMEApplication(open(attachment, 'rb').read())
            file.add_header("Content-Disposition", "attachment", filename = attachment)
            msg.attach(file)

        try:
            self.SMTP_server.sendmail(self.address, mail_list, msg.as_string())
            self.SMTP_server.quit()
            print ('Succeed.')
        except Exception as e:
            print ('Failed--' + str(e))







if __name__ == '__main__':

    # today = datetime(2021,7,16) ### for testing
    today = datetime.today()

    root_path = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.join(os.path.join(root_path, "RawData"), date_str(today))

    filename = "wind新股数据" + date_str(today) + ".xlsx"
    file_path = os.path.join(data_path, filename)

    history_save_path = os.path.join(root_path, "history.save")
    excel_save_path = os.path.join(os.path.join(root_path, "IPO_calendar"), date_str(today)+".xlsx")

    workday_sheet = pd.read_excel(file_path, sheet_name = 2)

    # tomorrow = today + timedelta(days = 1)
    tomorrow = next_workday(today, workday_sheet)

    IPO_raw = pd.read_excel(file_path, sheet_name = 0)
    subscription_raw = pd.read_excel(file_path, sheet_name = 1)
    

    # history = {}
    # joblib.dump(history, history_save_path)

    history = joblib.load(history_save_path)
    # print(history)

    data = IPO_calendar(IPO_raw, subscription_raw, workday_sheet, history)

    joblib.dump(history, history_save_path)

    writer = excel_writer(excel_save_path) 
    html_tables = writer.generate_IPO_calendar(data)
    html_tables += writer.today_offering(data)
    html_tables += writer.today_purchase(data)
    html_tables += writer.tomorrow_purchase(data)
    writer.save()
    print(today)
    # print(history)


    
    # receiver = ["charlesliangzhiyuan@gmail.com"]
    # receiver = ["liangzhy@shanghaitech.edu.cn"]
    receiver = ["charlesliangzhiyuan@gmail.com", "wujn@huianfund.cn"]

    # sender = "1450925659@qq.com"
    # authorization_code = "ffquxggkasxxiagg"
    # smtp_server='smtp.qq.com'
    # smtp_port=465

    # sender = "research@huianfund.cn"
    # authorization_code = "abc!1234"
    # smtp_server='mail.huianfund.cn'
    # smtp_port=25

    # Email = mail(sender, authorization_code, smtp_server, smtp_port)
    # mainbody = ""
    # for html_table in html_tables:
    #     mainbody += html_table
    # contents = MIMEText(mainbody, 'html', 'utf-8')
    # Email.send(receiver, contents, attachment=excel_save_path, today=date_str(today))






