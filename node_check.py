import pandas as pd
from UliPlot.XLSX import auto_adjust_xlsx_column_width
import pymysql
import os
import sys
import time

if __name__ == '__main__':
    ids = sys.argv[1].split(',')
    print(ids)
work_dttm = time.strftime('%Y년 %m월 %d일 %H시', time.localtime(time.time()))
result_list = list()


def make_dataframe(writer):
    conn = pymysql.connect(host='localhost',
                       port=3306,
                       user='root',
                       password='12341234',
                       db ='sensor',
                       charset='utf8', autocommit=False)
    cur = conn.cursor(pymysql.cursors.DictCursor)

    limit_amount = 50
    loop_count = 0

    for id in ids:
        sql = f"select * from connect where id=%s order by num desc limit %s"
        cur.execute(sql, (id, limit_amount))

        datas = cur.fetchall()
        data_frame = pd.DataFrame(datas)
        data_frame = data_frame.sort_values(by='num')

        date_datas = pd.to_datetime(data_frame['date'], format='%a %b %d %Y %H:%M:%S GMT%z (대한민국 표준시)')

        data_frame['diff'] = date_datas.diff().shift(-1)

        data_frame = data_frame.sort_values(by='num', ascending=False)


        print(data_frame)

        result_min = data_frame['diff'].min()
        result_max = data_frame['diff'].max()
        result_mean = data_frame['diff'].mean()
        result_list.append({'id': id, 'min': result_min, 'max': result_max, 'mean': result_mean})
        print('-' * 10)

        data_frame['diff'] = data_frame['diff'].astype('str')

        data_frame.to_excel(writer,
                            index=False,
                            sheet_name=work_dttm,
                            startcol=0,
                            startrow=6 + ((limit_amount + 2) * loop_count),
                            na_rep='N/A')
        auto_adjust_xlsx_column_width(data_frame, writer, sheet_name=work_dttm, margin=0)

        loop_count = loop_count + 1

    make_result(writer)
    cur.close()
    conn.close()

def make_result(writer):
    result_frame = pd.DataFrame(result_list)
    result_frame = result_frame.transpose()
    result_frame.rename(columns=result_frame.iloc[0], inplace=True)
    result_frame = result_frame.drop(result_frame.index[0])
    result_frame = result_frame.astype('str')

    result_frame.to_excel(writer,
                          index=True,
                          header=ids,
                          sheet_name=work_dttm,
                          startcol= 0,
                          startrow=0)

    auto_adjust_xlsx_column_width(result_frame, writer, sheet_name=work_dttm, margin=0)

    print(result_frame)


work_file = f'C:/Users/User/Desktop/node_check.xlsx'

if not os.path.exists(work_file):
    with pd.ExcelWriter(work_file, mode='w', engine='openpyxl') as writer:
        make_dataframe(writer)
else:
    with pd.ExcelWriter(work_file, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        make_dataframe(writer)