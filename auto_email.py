# coding=utf-8
import os
from datetime import datetime, date
from time import strftime, localtime

import xlrd

from email import encoders
from email.mime.base import MIMEBase
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

import smtplib

from xlrd import xldate_as_tuple


def read_file(file_path):
    file_list = []
    work_book = xlrd.open_workbook(file_path)
    sheet_data = work_book.sheet_by_name('Sheet1')
    Nrows = sheet_data.nrows

    for i in range(1, Nrows):
        #cell = sheet_data.cell_value(i, 2)
        #date = datetime(*xldate_as_tuple(cell, 0))
        #sheet_data.row_values(i)[2] = date.strftime('%Y/%d/%m')
        ##sheet_data.row_values(i)[2] = date(*xldate_as_tuple(sheet_data.cell(i, 2).value, work_book.datemode))[:3].strftime('%Y-%m-%d')
        file_list.append(sheet_data.row_values(i))
    print 'read_file file_list={}'.format(file_list)
    return file_list


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def sendEmail(from_addr, password, smtp_server, file_list):
    for i in range(len(file_list)):
        try:
            person_info = file_list[i]
            email_add = str(person_info[0])
            name = person_info[1].encode('utf-8')
            ruzhi = person_info[2]
            print 'ruzhi={}'.format(ruzhi)
            zhiwei = str(person_info[3])
            zhiji = str(person_info[4])
            shifa_shouru = '￥' + str(person_info[5])
            shuiqian_shouru = '￥' + str(person_info[6])
            mingyi_shouru = '￥' + str(person_info[7])
            suodeshui = '￥' + str(person_info[8])
            gongzi = '￥' + str(person_info[9])
            quanqing_jixiao_jiang = '￥' + str(person_info[10])
            jintie_huizong = '￥' + str(person_info[11])
            jiaban = '￥' + str(person_info[12])
            tepi = '￥' + str(person_info[13])
            queqing = '￥' + str(person_info[14])
            shebao_gere = '￥' + str(person_info[15])
            gongjijing_geren = '￥' + str(person_info[16])


            html_content = \
                '''
                <html>
                <body>
                    <h3 align="center">2020年3月工资单</h3>
                    <p>一直承蒙关照：</p>
                    <table border="1">
                    <tr>
                        <td>姓名</td>
                        <td>入职</td>
                        <td>职位</td>
                        <td>职级</td>
                        <td>实发收入</td>
                        <td>税前收入</td>
                        <td>名义收入</td>
                        <td>所得税</td>
                        <td>工资</td>
                        <td>全勤绩效奖</td>
                        <td>津贴汇总</td>
                        <td>加班</td>
                        <td>特批</td>
                        <td>缺勤</td>
                        <td>社保个人</td>
                        <td>公积金个人</td>
                    </tr>
                    <tr>
                      <td>{name}</td>
                      <td>{ruzhi}</td>
                      <td>{zhiwei}</td>
                      <td>{zhiji}</td>
                      <td>{shifa_shouru}</td>
                      <td>{shuiqian_shouru}</td>
                      <td>{mingyi_shouru}</td>
                      <td>{suodeshui}</td>
                      <td>{gongzi}</td>
                      <td>{quanqing_jixiao_jiang}</td>
                      <td>{jintie_huizong}</td>
                      <td>{jiaban}</td>
                      <td>{tepi}</td>
                      <td>{queqing}</td>
                      <td>{shebao_gere}</td>
                      <td>{gongjijing_geren}</td>
                    </tr>
                    </table>
    
    
                </body>
                </html>
                '''.format(name=name, ruzhi=ruzhi, zhiwei=zhiwei, zhiji=zhiji,
                           shifa_shouru=shifa_shouru, shuiqian_shouru=shuiqian_shouru,
                           mingyi_shouru=mingyi_shouru, suodeshui=suodeshui,
                           gongzi=gongzi, quanqing_jixiao_jiang=quanqing_jixiao_jiang,
                           jintie_huizong=jintie_huizong, jiaban=jiaban,
                           tepi=tepi, queqing=queqing, shebao_gere=shebao_gere,
                           gongjijing_geren=gongjijing_geren)

            # msg = MIMEText(text_content, 'html', 'utf-8') # html邮件
            msg = MIMEMultipart()
            msg.attach(MIMEText(html_content, 'html', 'utf-8'))

            msg['From'] = _format_addr('赵鹏鑫 <%s>' % from_addr)
            msg['To'] = _format_addr(name + '<%s>' % email_add)
            msg['Subject'] = Header('TEXT17:19-2020年3月工资明细', 'utf-8').encode()

            #server = smtplib.SMTP(smtp_server, 25)
            #server.starttls()  # 调用starttls()方法，就创建了安全连接
            ## server.set_debuglevel(1) # 记录详细信息
            #server.login(from_addr, password)  # 登录邮箱服务器
            #server.sendmail(from_addr, [email_add], msg.as_string())  # 发送信息
            #server.quit()
            print 'num:{0}, name:{1} send success.'.format(i+1, name)
        except Exception as e:
            print(e)


if __name__ == '__main__':
    root_dir = 'D:\PyWorkSpace\middlewaredemo\excel_email'
    file_path = root_dir + "\测试.xlsx"
    from_addr = 'xxxxx'  # 邮箱登录用户名
    password = 'xxxxx'  # 登录密码
    smtp_server = 'smtp.office365.com'  # 微软SMTP服务器地址，默认端口号587

    file_list = read_file(file_path)
    sendEmail(from_addr, password, smtp_server, file_list)
    print('all is ok')