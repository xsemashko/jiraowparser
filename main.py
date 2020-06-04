#! /usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Для работы скрипта необходимо первоначально авторизоваться в Jira OW в FireFox
и посмотерть сколько всего страниц с задачами необходимо парсить. При запросе
количества страниц ввести интересующее значение и ждать выгрузки файла.
"""

import requests
import browser_cookie3
from bs4 import BeautifulSoup as BS
import html5lib
import re
import json
import xlsxwriter

cj = browser_cookie3.firefox()

users = []

lines = open("users", "r")

for line in lines:
	line = line.rstrip("\n")
	users.append(line)

i=0
p=input("Введите число страниц для парсинга :")
header = ['Номер задачи', 'Описание', 'Автор обращения', 'Привязано', 'Статус', 'Приоритет', 'Дата обновления', 'Комментарий']
issue = [header]
issues = []
TAG_RE = re.compile(r'<[^>]+>')

def remove_tags(text):
    return TAG_RE.sub('', text)

while i < int(p):
	url2 = 'https://jira.openwaygroup.com/servicedesk/customer/user/requests?reporter=all&status=open&page=' + str(i)
	r = requests.get(url2, cookies=cj)
	soup = BS(r.content, 'html5lib')
	r1 = re.findall(r"OWCPE-\d+", str(soup))
	issues = issues + r1
	i = i + 1

for item in issues:

	url3 = 'https://jira.openwaygroup.com/servicedesk/customer/portal/29/' + item
	r = requests.get(url3, cookies=cj)
	soup = BS(r.content, 'html.parser')

	element = soup.find("div", class_="cv-json-fragment", id="jsonPayload").text
	jsonData = json.loads(element) 
	jdata = jsonData['reqDetails']['issue']

	a = jdata['key']
	b = jdata['summary']
	c = jdata['reporter']['displayName']
	d = jdata['assignee']['displayName']
	e = jdata['status']
	f = jdata['fields'][2]['value']['html']

	try:
		g = jdata['activityStream'][0]['friendlyDate']
		h = jdata['activityStream'][0]['comment']

		h = remove_tags(h)

	except:
		g = 'Null'
		h = 'Null'

	if c in users:
		body = [a,b,c,d,e,f,g,h]
		issue.append(body)


with xlsxwriter.Workbook('OW_tasks.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    for row_num, data in enumerate(issue):
        worksheet.write_row(row_num, 0, data)

print("Done!")