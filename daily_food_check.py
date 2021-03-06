import datetime
import os
import threading
import time

import pandas as pd
import requests
import xlrd
import easyocr

from concurrent.futures import ThreadPoolExecutor,as_completed
import threading

# @Author:hairu,wu
# @Time: 2022/5/6
# @Location: fudan u.

pool = ThreadPoolExecutor(max_workers=3)

save_path = './result.xls'

def read_file(path):
    return pd.read_excel(path)

# 获取基本数据
def basic_data(data):
    student_list = [] # 最后数据的列表
    for row in data.index.values:
        row_data = data.iloc[row]
        student = dict()
        student_name = row_data['姓名（必填）']
        student_school = row_data['所在学院（必填）']
        breakfast = row_data['早餐']
        lunch = row_data['午餐']
        lunch_rice = row_data['午餐白米饭']
        diner = row_data['晚餐']
        diner_rice = row_data['晚餐白米饭']
        all_money = row_data['支付总金额']
        alipay = row_data['支付宝付款截图上传']
        student['早餐'] = row_data['早餐']
        student['午餐'] = row_data['午餐']
        student['午餐白米饭'] = row_data['午餐白米饭']
        student['晚餐'] = row_data['晚餐']
        student['晚餐白米饭'] = row_data['晚餐白米饭']
        student['student_name'] = student_name
        student['student_school'] = student_school
        student['breakfast'] = breakfast
        student['lunch'] = lunch
        student['lunch_rice'] = lunch_rice
        student['diner'] = diner
        student['diner_rice'] = diner_rice
        student['all_money'] = all_money
        student['支付宝付款截图上传'] = alipay
        student_list.append(student)
    return student_list

def download_imgs(student):
    i = 0
    count = len(student)
    for stu in student:
        url = stu['支付宝付款截图上传']
        # 日期/学院/姓名.jpg
        path = r"./imgs/" + str(datetime.date.today ()) + "/" + stu['所在学院（必填）']+"/"
        img_name = r"./imgs/" + str(datetime.date.today ()) + "/" \
                   + stu['所在学院（必填）']+"/"+stu['姓名（必填）']+".jpg"
        if not os.path.exists(path):
            os.makedirs(path)
        if not os.path.isfile(img_name):
            img = requests.get(url)
            time.sleep(0.5)
            with open(img_name, "wb") as code:
                code.write(img.content)
        # os.system('cls' if os.name == 'nt' else 'clear')
        i += 1
        print("下载图片：已完成：", i / count*100  , "%")
    pass

# 判断坐标的包含关系
def check_location(location_1, location_2):
    # 判断location_1是否再location_2中
    idx = 0
    # 判断纵轴,location1的x坐标都在location_2的范围内
    left_low = location_2[0][0]
    left_high = location_2[1][0]
    flag_1 = True
    for item in location_1:
        left = item[0]
        if left < left_low or left > left_high:
            flag_1 = False
            break
    # 判断横轴
    flag_2 = True
    right_low = location_2[0][1]
    right_high = location_2[2][1]
    for item in location_1:
        right = item[1]
        if right < right_low or right >  right_high:
            flag_2 = False
            break
    return flag_1 and flag_2

# thread_worker
def thread_worker(stu):
    print("开启了一个工具")
    ret = dict()
    ret['所在学院（必填）'] = stu['所在学院（必填）']
    ret['姓名（必填）'] = stu['姓名（必填）']
    ret['支付宝付款截图上传'] = stu['支付宝付款截图上传']
    reader = easyocr.Reader(['ch_sim', 'en'], gpu=True)  # need to run only once to load model into memory
    uri = r"./imgs/" + str(datetime.date.today()) + "/" + stu['所在学院（必填）'] + "/" + stu['姓名（必填）'] + ".jpg"
    with open(uri, 'rb') as f:
        img = f.read()
        result = reader.readtext(img)
        # print(result)
        # 处理单张图片内容
        # 读取支付宝图片信息 [<支付人姓名>，<支付人学院>,<支付日期，xx>, <支付金额，xx>, <收款人，xx>,]
        # https://www.jaided.ai/easyocr/documentation/
        data_receiver = ''  # 收款人
        data_time = ''  # 转账时间
        data_money = ''  # 转账金额
        idx = len(result) - 1
        while idx >= 0:
            item = result[idx]  # item 为每一个元组
            idx -= 1
            # item[0] 为坐标框
            # 00, 10 ,     01 , 21
            # item[1] 为具体数据
            if '世英' in item[1]:
                data_receiver = item[1]
            if str(datetime.date.today()) in item[1]:
                data_time = str(datetime.date.today())
            if '.00' in item[1]:
                data_money = item[1]
        # 打印错误信息
        err = ''
        if len(data_receiver) == 0:
            err += '没有收款人信息;'
            ret['收款人'] = ''
        else:
            ret['收款人'] = data_receiver

        if len(data_money) == 0:
            err += "没有转账记录;"
            ret['转账金额'] = 0
        else:
            ret['转账金额'] = data_money[1:-1]

        if len(data_time) == 0:
            err += '没有转账时间;'
            ret['转账时间'] = ''
        else:
            ret['转账时间'] = data_time
        ret['err'] = err
    return ret


# 通过ocr来获取学生数据
def get_alipay(student):
    # 读取支付宝图片信息 [<支付人姓名>，<支付人学院>,<支付日期，xx>, <支付金额，xx>, <收款人，xx>,]
    # print(student)
    ans = []
    i = 1
    count = len(student)
    future_list = []
    for stu in student:
        future1 = pool.submit(thread_worker, stu)
        future_list.append(future1)
    for future1 in as_completed(future_list):
        ret = future1.result()
        lock = threading.Lock()
        lock.acquire()  # 加锁
        # ans是互斥变量
        ans.append(ret)
        print("处理进度：", i / count * 100, "%")
        i += 1
        time.sleep(0.01)
        lock.release()  # 解锁
    # print(ans)
    pool.shutdown()
    return ans


def alipay_data(path):
    # 根据excel中的链接读取信息
    data = xlrd.open_workbook(path, formatting_info=True)
    sheet_1 = data.sheet_by_index(0)
    # 读取所有列的信息
    keys = []
    for col in range(sheet_1.ncols):
        keys.append(sheet_1.cell_value(0, col))
    # 读取行的信息
    student = []
    for row_index in range(1, sheet_1.nrows):
        line_data = {}
        for col_index in range(sheet_1.ncols):
            if keys[col_index] == '支付宝付款截图上传':
                # print(sheet_1.cell_value(row_index, col_index))
                link = sheet_1.hyperlink_map.get((row_index,col_index))
                line_data[keys[col_index]] = link.url_or_path
            else:
                line_data[keys[col_index]] = sheet_1.cell_value(row_index, col_index)
        student.append(line_data)

    # 下载支付宝图片
    download_imgs(student)
    # 读取支付宝图片信息 [<支付日期，xx>, <支付金额，xx>, <收款人，xx>,<支付人姓名>，<支付人学院>]
    ret = get_alipay(student)
    return ret

# 检查总金额是否相等
def check_basic_data(basic_value):
    # print(basic_value)
    ret = []
    for item in basic_value:
        res = dict()
        res['学院'] = item['student_school']
        res['姓名'] = item['student_name']
        res['支付宝付款截图上传'] = item['支付宝付款截图上传']
        res['早餐'] = item['早餐']
        res['午餐'] = item['午餐']
        res['午餐白米饭'] = item['午餐白米饭']
        res['晚餐'] = item['晚餐']
        res['晚餐白米饭'] = item['晚餐白米饭']

        res['早餐'] = 0
        breakfast = item['breakfast']

        # 早餐
        res['早餐'] = 0
        if '1份' in breakfast:
            res['早餐'] = 10

        res['午餐'] = 0
        if '元' in str(item['lunch']):
            left = str(item['lunch']).index('（')
            right = str(item['lunch']).index('）')
            res['午餐'] = int(str(item['lunch'])[left + 1:right - 1])

        res['午餐米饭'] = 0
        if '元' in str(item['lunch_rice']):
            left = str(item['lunch_rice']).index('（')
            right = str(item['lunch_rice']).index('）')
            res['午餐米饭'] = int(str(item['lunch_rice'])[left + 1:right - 1])

        res['晚餐'] = 0
        if '元' in str(item['diner']):
            left = str(item['diner']).index('（')
            right = str(item['diner']).index('）')
            res['晚餐'] = int(str(item['diner'])[left + 1:right - 1])

        res['晚餐米饭'] = 0
        if '元' in str(item['diner_rice']):
            left = str(item['diner_rice']).index('（')
            right = str(item['diner_rice']).index('）')
            res['晚餐米饭'] = int(str(item['diner_rice'])[left + 1:right - 1])

        all_money = item['all_money']
        real_money = res['早餐'] + res['午餐'] +res['午餐米饭'] +res['晚餐'] + res['晚餐米饭']
        res['err'] = ''
        res['表格填写金额'] = real_money
        if all_money != real_money:
            res['err'] += '总金额填写不正确！'
        ret.append(res)
    # print(ret)
    return ret

def check_alipay_data(alipay_value, basic_check_value):
    res = []
    for stu in basic_check_value:
        res_item = dict()
        for item in alipay_value:
            # print(item)
            if stu['学院'] == item['所在学院（必填）'] and stu['姓名'] == item['姓名（必填）']:
                res_item['支付宝付款截图上传'] = item['支付宝付款截图上传']
                res_item['学院'] = item['所在学院（必填）']
                res_item['姓名'] = item['姓名（必填）']
                res_item['早餐具体'] = stu['早餐']
                res_item['午餐具体'] = stu['午餐']
                res_item['午餐白米饭具体'] = stu['午餐白米饭']
                res_item['晚餐具体'] = stu['晚餐']
                res_item['晚餐白米饭具体'] = stu['晚餐白米饭']
                res_item['转账时间'] = item['转账时间']
                res_item['表格填写金额'] = stu['表格填写金额']
                if float(stu['表格填写金额']) == float(item['转账金额']):
                    res_item['info'] = '正确'
                    res_item['转账金额'] = float(item['转账金额'])
                else:
                    res_item['转账金额'] = float(item['转账金额'])
                    res_item['info'] = '转账错误！'
                    res_item['info'] += item['err']
                break
        res.append(res_item)
    return res

# 将文字转为链接
def excel_url_fun(input_value):
    ''' 把一个网址字符串转换为excel公式 '''
    return f'=HYPERLINK("{input_value}","图片")'

# 将错误的数据写入到excel中
def writeExcel(alipay_check_value, data_origin):
    error = dict()
    cols = alipay_check_value[0].keys()
    for col in cols:
        error[col] = []

    for item in alipay_check_value:
        if not '正确' in item['info']:
            for key in cols:
                error[key].append(item[key])
    error_data = pd.DataFrame(error)
    from pandas import DataFrame
    df = DataFrame(error_data)
    # df['link'] = df['支付宝付款截图上传'].apply(lambda x: excel_url_fun(x))
    # df.to_excel(save_path)
    with pd.ExcelWriter(  save_path , engine='xlsxwriter',options={'strings_to_urls': True}) as writer:
        df.to_excel(writer, index=False)
    return error

if __name__ == '__main__':
    # 读取excel数据，获取姓名-总金额，获取早餐、午餐、午餐白米饭、晚餐、晚餐白米饭，图片结果
    # <学院+'_'+姓名，[<早餐，money>,<午餐，money>,<晚餐，money>,<总金额，money>,<图片结果,money>,<备注, [图片日期不对，总金额不匹配]>]>
    # ------ 必须将excel另存为.xls
    path = '0507-2.xls'
    data_origin = read_file(path)
    basic_value = basic_data(data_origin)    # 获取基本数据
    alipay_value = alipay_data(path)     #读取支付宝图片信息 [<支付日期，xx>, <支付金额，xx>, <收款人，xx>,<支付人姓名>，<支付人学院>]
    # 1、首先检查总金额和三餐需求是否匹配
    basic_check_value = check_basic_data(basic_value)
    # 2、其次检查支付结果是否和总金额匹配
    alipay_check_value = check_alipay_data(alipay_value, basic_check_value)
    # 3、输出所有结果 : [<学院，姓名，错误>]
    # print("所有结果列表", alipay_check_value)
    # 4、输出错误结果
    error = writeExcel(alipay_check_value, data_origin)
    # print("错误列表", error)
    pass
