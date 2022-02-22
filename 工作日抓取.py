import requests
import arrow

import xlwt
import time
import json
def isLeapYear(years):
  '''
  通过判断闰年，获取年份years下一年的总天数
  :param years: 年份，int
  :return:days_sum，一年的总天数
  '''
  # 断言：年份不为整数时，抛出异常。
  assert isinstance(years, int), "请输入整数年，如 2018"

  if ((years % 4 == 0 and years % 100 != 0) or (years % 400 == 0)):  # 判断是否是闰年
    # print(years, "是闰年")
    days_sum = 366
    return days_sum
  else:
    # print(years, '不是闰年')
    days_sum = 365
    return days_sum


def getAllDayPerYear(years):
  '''
  获取一年的所有日期
  :param years:年份
  :return:全部日期列表
  '''
  start_date = '%s-1-1' % years
  a = 0
  all_date_list = []
  days_sum = isLeapYear(int(years))
  print()
  while a < days_sum:
    b = arrow.get(start_date).shift(days=a).format("YYYY-MM-DD")
    a += 1
    all_date_list.append(b)
  # print(all_date_list)
  return all_date_list


if __name__ == '__main__':
  # years = "2001"
  # years = int(years)
  # # 通过判断闰年，获取一年的总天数
  # days_sum = isLeapYear(years)


  # 创建一个workbook 设置编码
  workbook = xlwt.Workbook(encoding='utf-8')
  # 创建一个worksheet
  worksheet = workbook.add_sheet('My Worksheet')

  # 写入excel
  # 参数对应 行, 列, 值
  worksheet.write(0, 0, label='date')
  worksheet.write(0, 1, label='workmk')
  worksheet.write(0, 2, label='worknm')
  worksheet.write(0, 3, label='week_1')
  worksheet.write(0, 4, label='week_2')
  worksheet.write(0, 5, label='week_3')
  worksheet.write(0, 6, label='week_4')
  worksheet.write(0, 7, label='remark')

  # url  = 'http://api.k780.com/?app=life.workday&date={}&appkey=10003&sign=b59bc3ef6191eb9f747dd4e83c99f2a4&format=json'.format("20220202")
  #
  # res = requests.get(url).text
  # res = json.loads(res)
  # if(res['success']=='1'):
  #   print(res['result']['date'])
  #   print(res['result']['workmk'])
  #   print(res['result']['worknm'])
  #   print(res['result']['week_1'])
  #   print(res['result']['week_2'])
  #   print(res['result']['week_3'])
  #   print(res['result']['week_4'])
  #   print(res['result']['remark'])
  url = 'http://api.k780.com'
  params = {
    'app': 'life.workday',
    'appkey': '64415',
    'sign': '841280974fec50865d3b8b6b8a6778c7',
    'format': 'json',
  }
  col = 1
  # 获取一年的所有日期
  all_date_list = getAllDayPerYear("2023")
  all_date_list=[i.replace("-","") for i in all_date_list]
  print(len(all_date_list)//80)
  for i in range(4):
      params['date'] = ','.join(all_date_list[i*80:(i+1)*80])
      print(params)
      res = requests.get(url,params).text
      res = json.loads(res)
      print(res)
      if (res['success'] == '1'):
        for i in res['result']:
          worksheet.write(col, 0, label=i['date'])
          worksheet.write(col, 1, label=i['workmk'])
          worksheet.write(col, 2, label=i['worknm'])
          worksheet.write(col, 3, label=i['week_1'])
          worksheet.write(col, 4, label=i['week_2'])
          worksheet.write(col, 5, label=i['week_3'])
          worksheet.write(col, 6, label=i['week_4'])
          worksheet.write(col, 7, label=i['remark'])
          col=col+1
  print(col)
  params['date'] = ','.join(all_date_list[320:])
  print(params)
  res = requests.get(url, params).text
  res = json.loads(res)
  if (res['success'] == '1'):
    for i in res['result']:
      worksheet.write(col, 0, label=i['date'])
      worksheet.write(col, 1, label=i['workmk'])
      worksheet.write(col, 2, label=i['worknm'])
      worksheet.write(col, 3, label=i['week_1'])
      worksheet.write(col, 4, label=i['week_2'])
      worksheet.write(col, 5, label=i['week_3'])
      worksheet.write(col, 6, label=i['week_4'])
      worksheet.write(col, 7, label=i['remark'])
      col = col + 1
  workbook.save('2023年节假日详情.xls')