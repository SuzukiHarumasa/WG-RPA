import sys
sys.path.append('./env/lib/python3.10/site-packages')
sys.path.append('./pytools')
import requests
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import urllib3
import pandas as pd
from gsheets import GSheets

# 検索上位取得


# インスタンスを生成


def firstPrepFunc(all_input_dict):
    creds = './seacrets/rpa-shokikikaku-a6bba716cf27.json'

    gs = GSheets(creds)

    # 必要な情報を取得する
    input_id = '1hQikhwgt8VSNC6GcfhFhHqQPAfoLygBJrK0cG_Ja3xc'
    sheets = gs.gc.open_by_key(input_id)

    sheet = sheets.worksheet('input')
    content = sheet.get_values()

    sheet = sheets.worksheet('input')
    content = sheet.get_values()
    # シートの中身をデータフレーム化

    # 必要なファイルの作成

    creiant_name = all_input_dict['対策サイト_サイト名']
    # SEO課題分析調査
    file_name = f'{creiant_name}様_SEO課題分析調査'

    first_prep_file_name = f'{creiant_name}様_初期企画準備ファイル（テンプレ） 【最終編集 : 2022/01/05】'

    keyword_file_name = f'{creiant_name}様_キーワード調査テンプレ'

    # テンプレートを作成
    DRIVE_ID = '1ma2kClZS0NYY9tipellTbvnAD89FU30v'
    FILE_ID = '1LNNuszsOHeLc7R8LVMNv3zOLejkCURUt-tyntBk1wgw'
    FIRST_PREP_FILE_ID = '1t3Vc_MB0NcVD52qgshnv06uU4WTlJzwXqVjOnaXGr2A'
    KEYWOED_FILE_ID = '1rgt1y0Gh1gvrnZKfxUZGLGS7X_GA8Km-3pfA6ImP3Ic'

    new_file_body = {
        'name': file_name,  # 新しいファイルのファイル名. 省略も可能
        'parents': [DRIVE_ID],  # Copy先のFolder ID. 省略も可能
    }

    first_prep_new_file_body = {
        'name': first_prep_file_name,  # 新しいファイルのファイル名. 省略も可能
        'parents': [DRIVE_ID],  # Copy先のFolder ID. 省略も可能
    }

    keyword_new_file_body = {
        'name': keyword_file_name,  # 新しいファイルのファイル名. 省略も可能
        'parents': [DRIVE_ID],  # Copy先のFolder ID. 省略も可能
    }

    new_file = gs.api.files().copy(
        fileId=FILE_ID, body=new_file_body
    ).execute()

    first_prep_new_file = gs.api.files().copy(
        fileId=FIRST_PREP_FILE_ID, body=first_prep_new_file_body
    ).execute()

    keyword_new_file = gs.api.files().copy(
        fileId=KEYWOED_FILE_ID, body=keyword_new_file_body
    ).execute()

    # SEO課題分析調査
    output_id = new_file['id']

    output_sheets = gs.gc.open_by_key(output_id)

    seo_field1 = [
        "対策サイト",
        "対策サイト_サイト名",
        "対策サイト（ドメイン抽出）",
        "競合①_URL",
        "競合①_サイト名",
        "競合②_URL",
        "競合②_サイト名",
        "競合③_URL",
        "競合③_サイト名",
        "競合④_URL",
        "競合④_サイト名",
        "競合⑤_URL",
        "競合⑤_サイト名",
        "競合⑥_URL",
        "競合⑥_サイト名",
        "最重要キーワード",
        "上位表示分析結果URL",
    ]

    seo_field2 = [
        "GAのアカウントID",
        "GAのビューID"
    ]

    def to_input_dict(fields):
        input_dict = {}
        for field in fields:
            input_dict[field] = all_input_dict[field]

        return input_dict

    input_dict1 = to_input_dict(seo_field1)
    input_dict2 = to_input_dict(seo_field2)

    output_sheet = output_sheets.worksheet('【※はじめに確認※】調査概要')
    ds = output_sheet.range('D2:D18')
    for i, item in enumerate(input_dict1.items()):
        ds[i].value = item[1]

    output_sheet.update_cells(ds)

    ds = output_sheet.range('D20:D21')
    for i, item in enumerate(input_dict2.items()):
        ds[i].value = item[1]

    output_sheet.update_cells(ds)

    # 初期企画準備ファイル（テンプレ）
    first_prep_id = first_prep_new_file['id']

    first_prep_sheets = gs.gc.open_by_key(first_prep_id)

    first_prep_sheet = first_prep_sheets.worksheet('【※はじめに確認※】調査概要')
    first_prep_content = first_prep_sheet.get_values()
    first_prep_field = ['対策サイト', '取引先webサイト',
                        'メインキーワード', '指定競合サイト×3社以上', '初期企画依頼日']
    tact_url = all_input_dict['上位表示分析結果URL'].split('seo_improvement')[0]

    first_prep_input = to_input_dict(first_prep_field)

    def update_sheet(range, sheet, input_dict):
        ds = sheet.range(range)

        for i, item in enumerate(input_dict.items()):
            ds[i].value = item[1]

        sheet.update_cells(ds)

    update_sheet('E2:E6', first_prep_sheet, first_prep_input)

    # キーワード調査テンプレ
    keyword_id = keyword_new_file['id']

    keyword_sheets = gs.gc.open_by_key(keyword_id)

    keyword_sheet = keyword_sheets.worksheet('【※はじめに確認※】調査概要')
    keyword_content = keyword_sheet.get_values()
    first_prep_url = f'https://docs.google.com/spreadsheets/d/{first_prep_id}'

    keyword_sheet.update_acell('D2', first_prep_url)

    # 初期企画準備ファイル（テンプレ） 途中から
    seo_url = f'https://docs.google.com/spreadsheets/d/{output_id}'
    keyword_url = f'https://docs.google.com/spreadsheets/d/{keyword_id}'
    first_prep_sheet.update_acell('H3', seo_url)
    first_prep_sheet.update_acell('H4', keyword_url)
    first_prep_sheet.update_acell('H6', tact_url)

    article_comps = []
    article_comps = all_input_dict["記事系競合（最大8つまで）"].splitlines()

    service_comps = []
    service_comps = all_input_dict["カテゴリ系競合（最大8つまで）※除外するディレクトリがある場合には『,』を記載してください"].splitlines(
    )
    service_comps

    dicts = {}
    http = urllib3.PoolManager()
    for i, comp in enumerate(service_comps):
        key = service_comps[i].split(",")[0]
        response = http.request('get', key)
        soup = BeautifulSoup(response.data, 'html.parser')
        title = soup.find('title').text
        try:
            val = service_comps[i].split(",")[1]
        except:
            val = ''
        dicts[key] = [val, title]

    first_prep_sheet = first_prep_sheets.worksheet('【#16】競合キーワード抽出「初期設定」')
    ds = first_prep_sheet.range('D11:D18')
    for i, items in enumerate(dicts.items()):
        ds[i].value = items[0]

    first_prep_sheet.update_cells(ds)

    ds = first_prep_sheet.range('E11:E18')
    for i, items in enumerate(dicts.items()):
        ds[i].value = items[1][0]
    first_prep_sheet.update_cells(ds)

    ds = first_prep_sheet.range('C11:C18')
    for i, items in enumerate(dicts.items()):
        ds[i].value = items[1][1]

    first_prep_sheet.update_cells(ds)

    article_dicts = {}
    http = urllib3.PoolManager()
    for comp in article_comps:
        key = comp
        response = http.request('get', key)
        soup = BeautifulSoup(response.data, 'html.parser')
        title = soup.find('title').text

        article_dicts[key] = title

    first_prep_sheet = first_prep_sheets.worksheet('【#16】競合キーワード抽出「初期設定」')
    ds = first_prep_sheet.range('J11:J18')
    for i, items in enumerate(article_dicts.items()):
        ds[i].value = items[0]

    first_prep_sheet.update_cells(ds)

    ds = first_prep_sheet.range('I11:I18')
    for i, items in enumerate(article_dicts.items()):
        ds[i].value = items[1]

    first_prep_sheet.update_cells(ds)

    search_word = all_input_dict['メインキーワード']

    pages_num = 25 + 1

    print(f'【検索ワード】{search_word}')

    # Googleから検索結果ページを取得する
    url = f'https://www.google.co.jp/search?hl=ja&num={pages_num}&q={search_word}'
    request = requests.get(url)

    # Googleのページ解析を行う
    soup = BeautifulSoup(request.text, "html.parser")
    search_site_list = soup.select('div.kCrYT > a')

    url_list = []

    # ページ解析と結果の出力
    for rank, site in zip(range(1, pages_num), search_site_list):
        try:
            site_title = site.select('h3.zBAuLc')[0].text
        except IndexError:
            site_title = site.select('img')[0]['alt']
        site_url = site['href'].replace('/url?q=', '').split('&sa=U')[0]
        url_list.append(site_url)

    domain_list = []
    my_site = all_input_dict['対策サイト']

    for url in url_list:
        domain = '{uri.scheme}://{uri.netloc}/'.format(uri=urlparse(url))
        domain_list.append(domain)

    domain_list = list(filter(lambda x: x != my_site, domain_list))

    if len(domain_list) != 20:
        domain_list = list(set(domain_list))
        domain_list = domain_list[:20]

    first_prep_sheet = first_prep_sheets.worksheet('【#11~12】競合調査2.0')
    ds = first_prep_sheet.range('D43:D62')
    for i, key in enumerate(domain_list):
        ds[i].value = domain_list[i]

    first_prep_sheet.update_cells(ds)

    return first_prep_url, keyword_url, seo_url
