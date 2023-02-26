import streamlit as st
import datetime
import sys
sys.path.append('./env/lib/python3.10/site-packages')
# sys.path.append('./pytools/')

import glob


from firstPrepModule import firstPrepFunc

input_fields = ['日時', 'メールアドレス', '対策サイト', '対策サイト_サイト名', '対策サイト（ドメイン抽出）', '取引先webサイト', 'メインキーワード', '指定競合サイト×3社以上', '初期企画依頼日', '競合①_URL', '競合①_サイト名', '競合②_URL', '競合②_サイト名', '競合③_URL', '競合③_サイト名', '競合④_URL',
                '競合④_サイト名', '競合⑤_URL', '競合⑤_サイト名', '競合⑥_URL', '競合⑥_サイト名', '最重要キーワード', '上位表示分析結果URL', '自社及び競合サイトのディレクトリ構造', 'GAのアカウントID', 'GAのビューID', 'カテゴリ系競合（最大8つまで）※除外するディレクトリがある場合には『,』を記載してください', '記事系競合（最大8つまで）']

with st.form("my_form"):
    # st.write(sys.path)
    # files = glob.glob("./*")
    # for file in files:
    #     st.write(file)
    st.write("初期企画準備ファイル　用意フォーム")

    request_date = st.date_input(
        "依頼日",
        datetime.date.today())
    email = st.text_input('メールアドレス','suzuki.harumasa@willgate.co.jp')
    mysite = st.text_input('対策サイト','https://www.wifi-rental.com/')
    mysite_name = st.text_input('対策サイト_サイト名','WiFiレンタルどっとこむ')
    mysite_dmain = st.text_input('対策サイト（ドメイン抽出）','https://www.wifi-rental.com/')
    website = st.text_input('取引先webサイト','株式会社ビジョン（wifi-rental.com）')
    main_keyword = st.text_input('メインキーワード','Wifiレンタル')
    name = st.text_input('指定競合サイト×3社以上','https://www.netage.ne.jp/')
    first_request_date = st.date_input(
        "初期企画依頼日",
        datetime.date.today())
    comp1 = st.text_input('競合①_URL','https://www.netage.ne.jp/')
    comp1_name = st.text_input('競合①_サイト名','【公式】NETAGE')
    comp2 = st.text_input('競合②_URL','https://wifistore.jp/')
    comp2_name = st.text_input('競合②_サイト名','【公式】WiFiストア')
    comp3 = st.text_input('競合③_URL','https://www.rental-store.jp/')
    comp3_name = st.text_input('競合③_サイト名','WiFiレンタル屋さん')
    comp4 = st.text_input('競合④_URL','https://wimax-broad.jp/')
    comp4_name = st.text_input('競合④_サイト名','Broad WiMAX')
    comp5 = st.text_input('競合⑤_URL','https://www.just-size.net/')
    comp5_name = st.text_input('競合⑤_サイト名','Just-Size.Networks')
    comp6 = st.text_input('競合⑥_URL','https://shibarinashi-wifi.jp')
    comp6_name = st.text_input('競合⑥_サイト名','《公式》縛りなしWiFi')
    imp_keyword = st.text_input('最重要キーワード','海外 wifi レンタル')
    tact_url = st.text_input('上位表示分析結果URL','https://tact-seo.com/clients/815/sites/915/seo_improvement/rank_factor_analysis/page_analysis?analysis_id=43837')
    dict_stracture = st.text_input('自社及び競合サイトのディレクトリ構造')
    ga_id = st.text_input('GAのアカウントID','t3dwillgate@gmail.com')
    ga_view = st.text_input('GAのビューID','52781251')
    comp_category = st.text_area(
        'カテゴリ系競合（最大8つまで）※除外するディレクトリがある場合には『,』を記載してください','https://www.wifi-rental.com/,media')
    comp_article = st.text_area('記事系競合（最大8つまで）','https://www.wifi-rental.com/media/')

    data_list = [request_date, email, mysite, mysite_name, mysite_dmain, website, main_keyword, name, first_request_date, comp1, comp1_name, comp2, comp2_name,
                 comp3, comp3_name, comp4, comp4_name, comp5, comp5_name, comp6, comp6_name, imp_keyword, tact_url, dict_stracture, ga_id, ga_view, comp_category, comp_article]
    # Every form must have a submit button.
    submitted = st.form_submit_button("Submit")

    if submitted:
        with st.spinner('ありがとうございます！少々お待ちください！'):
            input_dict = {}

            for key, data in zip(input_fields, data_list):
                input_dict[key] = data

            input_dict['日時'] = input_dict['日時'].strftime('%Y/%m/%d')
            input_dict['初期企画依頼日'] = input_dict['初期企画依頼日'].strftime('%Y/%m/%d')
            

            first_prep_url, keyword_url, seo_url= firstPrepFunc(input_dict)

            if first_prep_url and keyword_url and seo_url:
                st.balloons()
                st.success('完了しました！!!')
        
        

                st.write("初期企画準備ファイル :sunglasses:")
                st.write(first_prep_url)
                st.write("SEO課題分析 :sunglasses:")
                st.write(seo_url)
                st.write("キーワードシート :sunglasses:")
                st.write(keyword_url)


# form = st.form("my_form")
# form.text_input('メールアドレス')
# form.text_input('対策サイト')
# form.text_input('対策サイト_サイト名')
# form.text_input('対策サイト（ドメイン抽出）')
# form.text_input('取引先webサイト')

# form.form_submit_button("Submit")

# st.write(form)
