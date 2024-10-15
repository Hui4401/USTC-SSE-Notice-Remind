import requests
from bs4 import BeautifulSoup
import datetime
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email import encoders
from urllib.parse import unquote

import config


BASE_URL = 'http://mis.sse.ustc.edu.cn'
VALID_URL = BASE_URL + '/ValidateCode.aspx?ValidateCodeType=1&0.011150883024061309'
SSE_URL = BASE_URL + '/default.aspx'
HOME_PAGE = BASE_URL + '/homepage/StuHome.aspx'


# 计算验证码数字之和
def calculate_code(codes):
    res = 0
    for code in codes:
        res += int(code)
    return res


def sendmail(title, author, time, content, url="", attachments=[], possible_error=0):

    msg = MIMEMultipart()
    msg['subject'] = title + ' ' + author + ' ' + time
    msg["from"] = formataddr(['NoticeRemainder', config.SMTP_SENDER])
    msg["to"] = ','.join(config.SMTP_RECIVER)
    
    if possible_error > 0:
        content += "<br>可能有" + str(possible_error) + "个附件下载失败，请前往信息化平台手动下载"
    content = "<br><a href='" + url + "'>公告链接</a>" + content
    msg.attach(MIMEText(content, 'html', 'utf-8'))
    
    for attachment in attachments:
        filename = attachment['filename']
        content = attachment['content']
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(content)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)
    
    with SMTP_SSL(config.SMTP_HOST, config.SMTP_PORT) as smtp:
        smtp.login(config.SMTP_USERNAME, config.SMTP_PASSWORD)
        smtp.sendmail(config.SMTP_SENDER, config.SMTP_RECIVER, msg.as_string())


# 解析公告列表，每个公告形式：(标题，发布人，时间，详细链接)
def parse_notice(html):
    notices = []
    soup = BeautifulSoup(html, 'lxml')
    notice_nodes = soup.find(id="global_LeftPanel_UpRightPanel_ContentPanel2_ContentPanel3_content").find_all('tr')
    for node in notice_nodes:
        title = node.find_all('td')[0].text
        author = node.find_all('td')[1].text
        time = node.find_all('td')[2].text
        link = BASE_URL + node.find_all('td')[0].a['href']
        notices.append((title, author, time, link))
    return notices

# 解析公告内容，将所有链接转换为绝对链接
def parse_attachment(url, html_content, session):
    possible_error = 0
    attachments = []
    soup = BeautifulSoup(html_content, "html.parser")
    forms = soup.find_all('form')
    for form in forms:
        try:
            method = form.get('method')
            data = {x.get('name'): x.get('value') for x in form.find_all('input')}
            links = form.find_all('a')
            for link in links:
                # 每个链接对应一个附件
                href = link.get('href')
                if href.startswith('javascript'):
                    params = href.split("javascript:__doPostBack(")[1].split(")")[0].split(",")
                    data['__EVENTTARGET'] = params[0][1:-1] # 移除单引号
                    data['__EVENTARGUMENT'] = params[1][1:-1]
                    method = session.post if method == 'post' else session.get
                    res = method(url, data=data)
                    if res.headers.get('Content-Type') == 'application/octet-stream':
                        filename = res.headers.get('Content-Disposition').split('filename=')[1]
                        attachments.append({
                            'filename': unquote(filename),
                            'content': res.content
                        })
        except Exception as e:
            possible_error += 1
    return attachments, possible_error
                           
def main():
    year  = datetime.datetime.now().year
    month = datetime.datetime.now().month
    day   = datetime.datetime.now().day
    cur_date = str(year) + "-" + str(month) + "-" + str(day)
    with requests.Session() as s:
        res = s.get(VALID_URL)
        codes = res.cookies['CheckCode']
        code = calculate_code(codes)
        data = {
            '__EVENTTARGET' : 'winLogin$sfLogin$ContentPanel1$btnLogin',
            'winLogin$sfLogin$txtUserLoginID' : config.USERNAME,
            'winLogin$sfLogin$txtPassword' : config.PASSWORD,
            'winLogin$sfLogin$txtValidate' : code,
        }
        s.post(SSE_URL, data=data)
        res = s.get(HOME_PAGE)
        notices = parse_notice(res.text)
        for notice in notices:
            if notice[2] == cur_date:
                res = s.get(notice[3])
                attachments, possible_error = parse_attachment(notice[3], res.text, s)
                sendmail(notice[0], notice[1], notice[2], res.text, notice[3], attachments, possible_error)



if __name__ == '__main__':
    main()