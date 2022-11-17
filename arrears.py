import os
import shutil
from requests import Session
from yaml import safe_load
from bs4 import BeautifulSoup
from datetime import date
from xlsxwriter import Workbook


def parse(soup, scc):
    res = []
    soup_list = soup.select(scc)
    soup_list = [el.text for el in soup_list]
    for i in range(0, len(soup_list), 10):
        if i > 0 and soup_list[i] != 'Дата заявки':
            res.append(tuple(soup_list[i:i+10]))
    return res


def write_xlsx(res, name, fio):
    wordbook = Workbook(name)
    worksheet = wordbook.add_worksheet()
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:D', 64)
    worksheet.set_column('E:E', 25)
    worksheet.set_column('F:F', 12)
    worksheet.set_column('G:G', 40)
    worksheet.write(0, 0, f'Нет в архиве {fio}')
    row = 1
    col = 0
    for tpl in res:
        for i in range(len(tpl)):
            worksheet.write(row, col + i, tpl[i])
        row += 1
    wordbook.close()


if os.path.exists("RES"):
    shutil.rmtree('RES')
os.mkdir('RES')

with open('config.yaml') as cfg:
    config = safe_load(cfg)

usc_url = config['SERT_SITE']
data = {
    'email': config['SERT_LOGIN'],
    'password': config['SERT_PASS'].encode('CP1251')
}
with Session() as s:
    s.post(usc_url, data)
    req = s.get(usc_url + '/ru/registration_adminsUK/')
    soup = BeautifulSoup(req.text, 'lxml')
    mydiler = soup.select('select.in-text2 option')
    dilers = {}
    for node in mydiler:
        if len(node['value']) <= 3:
            dilers[node['value']] = node.text
    for catalog in dilers.values():
        os.mkdir(f'RES/{catalog}')
    for diler_kod, diler_name in dilers.items():
        data = {'mydiler': diler_kod}
        req = s.post(usc_url + '/ru/registration_adminsUK/', data)
        soup = BeautifulSoup(req.text, 'lxml')
        inn = soup.select('select.in-text option')
        for node in inn:
            if len(node['value']) == 10:
                data = {
                    'dt1': '2014.01.01',
                    'dt2': date.today().strftime('%Y.%M.%D'),
                    'inn': node['value'],
                    'act': '5',
                    'mydiler': diler_kod
                }
                r = s.get(usc_url+'/inc/popup_adminsUK.php', params=data)
                r.encoding = 'CP1251'
                soup = BeautifulSoup(r.text, 'lxml')
                if links := soup.select('a[href*="popup"]'):
                    last_links = links[-1]
                    result = []
                    result += parse(soup, 'table tr td')
                    for i in links[1:]:
                        r = s.get(usc_url + '/inc/popup_adminsUK.php', params=dict(data, num=i))
                        r.encoding = 'CP1251'
                        soup = BeautifulSoup(r.text, 'lxml')
                        result += parse(soup, 'table tr td')
                file_name = f"RES/{diler_name}/{node['value']}.xlsx"
                write_xlsx(result, file_name, node.text)
                # with open(f"RES/{diler_name}/{node['value']}.xlsx", 'w') as admreg_file:
                #     admreg_file.write(soup.prettify())



