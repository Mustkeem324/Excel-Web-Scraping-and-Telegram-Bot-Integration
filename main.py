import requests
from bs4 import BeautifulSoup  as s
import json
import openpyxl
#import urllib.parse
import os.path
import telebot
#telgram bot detail
TOKEN = "6002286894:AAFV-QFYZqdAg63nTYuvE2D_qmjoBfdlG1E"
bot = telebot.TeleBot(TOKEN, parse_mode=None, threaded=True)
# Define the range of rows to process
start_row = 1
end_row = 26000

# Load the last processed row from a file
try:
    with open("last_processed_row.txt", "r") as f:
        last_processed_row = int(f.read().strip())
except FileNotFoundError:
    last_processed_row = start_row

# Load the Excel spreadsheet
wb = openpyxl.load_workbook('193198.xlsx')
print(wb.sheetnames)
sheet = wb['Sheet1']

# Loop through the URLs in the spreadsheet
for i2, row in enumerate(sheet.iter_rows(min_row=start_row, max_row=end_row, max_col=1, values_only=True), start=start_row):
    if i2 <= last_processed_row:
        continue
    url = row[0]
    print(f"Processing row {i2}: {url}")
    try:
        payload2={}
        headers2 = {
        'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:87.0) Gecko/20100101 Firefox/87.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'ar,en-US;q=0.7,en;q=0.3',
        'Cookie': '_gid=GA1.2.1448162549.1669728615; _gat_UA-20169249-1=1; __attentive_id=f0b67eb44dad446e804aef87f37bafd2; _attn_=eyJ1Ijoie1wiY29cIjoxNjY5NzI4NjE1MDE1LFwidW9cIjoxNjY5NzI4NjE1MDE1LFwibWFcIjoyMTkwMCxcImluXCI6ZmFsc2UsXCJ2YWxcIjpcImYwYjY3ZWI0NGRhZDQ0NmU4MDRhZWY4N2YzN2JhZmQyXCJ9In0=; __attentive_cco=1669728615019; _tt_enable_cookie=1; _ttp=46201817-79fb-4b27-8ff4-4163b4fa9943; ajs_anonymous_id=358cc210-e160-4839-8664-02026ee0b450; _hjSessionUser_1400488=eyJpZCI6ImM0NzllOGU4LWZlZTktNTY3YS1iMDM4LTU4ZTY4YTY0NzhiYSIsImNyZWF0ZWQiOjE2Njk3Mjg2MTUwNzcsImV4aXN0aW5nIjp0cnVlfQ==; G_ENABLED_IDPS=google; tpc_a=d8a26a2810a84f84b31fcf69693d4cbc.1673019785.hlA.1675611602; _gcl_au=1.1.77948795.1678011865; viewedPagesCount=3; cookieconsent_status=dismiss; writeUserStatus=nonPremium; _gat_UA-93748-2=1; _ga_HXSFB2SGLZ=GS1.1.1678888367.2.0.1678888367.0.0.0; _ga=GA1.2.701438214.1669728615; ABTasty=uid=vk598ddz2ww2k4fs&fst=1669728613009&pst=1678888373674&cst=1678980265593&ns=27&pvt=97&pvis=2&th=; __attentive_pv=1; __attentive_ss_referrer=ORGANIC; __attentive_dv=1; _hjIncludedInSessionSample_1400488=1; _hjSession_1400488=eyJpZCI6Ijc5ZDQ5YjNjLWZmN2EtNGYxYS1hYTE3LWU2OWJhZWQzZjQ0MiIsImNyZWF0ZWQiOjE2NzkxNjg5Mzc3NjUsImluU2FtcGxlIjp0cnVlfQ==; _uetsid=ec555310c5c511ed86cbf3c857a727a7; _uetvid=f530c5306fe911ed873503b892b60087; OptanonConsent=isGpcEnabled=0&datestamp=Sun+Mar+19+2023+01%3A18%3A59+GMT%2B0530+(India+Standard+Time)&version=6.32.0&isIABGlobal=false&hosts=&consentId=5a26afa5-5852-4149-8d9a-65d0ffacbd74&interactionCount=1&landingPath=NotLandingPage&groups=C0001%3A1%2CC0003%3A1%2CC0002%3A0%2CC0005%3A0%2CC0004%3A0%2CBG156%3A0&AwaitingReconsent=false'
        }
        #encoded_sentence = urllib.parse.quote(url)
        urltext = "https://www.bartleby.com/search?scope=Solutions&q="+str(url)  # Replace with your desired URL
        response22 = requests.get(urltext, headers=headers2 ,data=payload2)
        soup22 = s(response22.content, 'html.parser')
        element = soup22.find(id="__NEXT_DATA__")
        json_data = json.loads(element.text)
        documents = json_data['props']['pageProps']['solutionResults']['documents']

        urlquestion = None

        if len(documents) > 0:
            urlquestion = documents[0].get('url')
        if urlquestion is None and len(documents) > 1:
            urlquestion = documents[1].get('url')
        if urlquestion is None and len(documents) > 2:
            urlquestion = documents[2].get('url')
        if urlquestion is None and len(documents) > 3:
            urlquestion = documents[3].get('url')
        if urlquestion is None and len(documents) > 4:
            urlquestion = documents[4].get('url')         
        if urlquestion is None:
            # Handle case where there is no URL for the documents
            print("No URL found for document")
            with open('url_not_found.txt', 'a') as f:
                f.write(str(url) + "\n")
        else:
            qurl = "https://www.bartleby.com" + urlquestion
            print(f'Your Link :{qurl}')
            # Process the URL as needed
            payload = {}
            headers = {
                'User-Agent':
                'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:87.0) Gecko/20100101 Firefox/87.0',
                'Accept':
                'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language':
                'ar,en-US;q=0.7,en;q=0.3',
                'Cookie':'cookieconsent_status=dismiss; writeUserStatus=nonPremium; _gid=GA1.2.334021265.1674776205; _gat_UA-93748-2=1; _fbp=fb.1.1674776205556.866881809; _tt_enable_cookie=1; _ttp=usg9UASAqZLgoAjgP49nm8Nyf4p; fs.session.id=ff963f23-d134-4581-b020-b9168a44c4cb; __qca=P0-656787387-1674776205430; _pbjs_userid_consent_data=3524755945110770; _lr_env_src_ats=false; _au_1d=AU1D-0100-001674776208-H7CDR6FR-W3MT; __gads=ID=9a6c8d38584c5e4f-225c5e236cd90001:T=1674776206:S=ALNI_MYTwo53TZTq9OKm8gGgKN8Xpq8WZA; _hjSessionUser_1400488=eyJpZCI6IjcwNzIyMGNlLTA0ZjYtNTczYy04MzgwLWRhZjk1ZWRkZWMzOCIsImNyZWF0ZWQiOjE2NzQ3NzYyMDU1OTAsImV4aXN0aW5nIjp0cnVlfQ==; _gcl_au=1.1.327136428.1674965019; _gat_UA-20169249-1=1; __attentive_id=a3a3b2df02ae4029a45aeb2a842b3398; _attn_=eyJ1Ijoie1wiY29cIjoxNjc0OTY1MDI3NjA5LFwidW9cIjoxNjc0OTY1MDI3NjA5LFwibWFcIjoyMTkwMCxcImluXCI6ZmFsc2UsXCJ2YWxcIjpcImEzYTNiMmRmMDJhZTQwMjlhNDVhZWIyYTg0MmIzMzk4XCJ9In0=; __attentive_cco=1674965027615; G_ENABLED_IDPS=google; cdn.bartleby.126401.ka.ck=d8eb4b2b4eb6236cd4ba9a8225f8576591c16f6b52298d32faf9b6d713603bf9af8793971bfa13c97dcece9222ece6dcf779248159a104b078bfa747f64e265a3cb89fdbb260da2ce91c631ba08205917b1f8076011cb3265533a682ead9c1b5a4e97b300bbc660a9d27e412c38ed76245cc3905129c6a86ebf2bda7c17558c09478027a874c7806926dc791f9dee8f03865e145539621b367618a; ajs_anonymous_id=a5e2245e-e140-47cf-ba0d-8f083cbc5803; ki_t=1675351091608%3B1675955497014%3B1675970659032%3B11%3B51; viewedPagesCount=3; _awl=2.1676835255.5-82711991fcab23f57c1d4cc8b6d259d9-6763652d617369612d6561737431-0; cookie=238f83cb-a38e-4622-aed7-ddef83101938; _au_last_seen_pixels=eyJhcG4iOjE2Nzc2MDA1MjgsInR0ZCI6MTY3NzYwMDUyOCwicHViIjoxNjc3NjAwNTI4LCJydWIiOjE2Nzc2MDA1MjgsInRhcGFkIjoxNjc3NjAwNTI4LCJhZHgiOjE2Nzc2MDA1MjgsImdvbyI6MTY3NzYwMDUyOCwidGFib29sYSI6MTY3NDc3NjIwOCwidW5ydWx5IjoxNjc0Nzc2MjA4LCJpbXByIjoxNjc1NzUxNTg4LCJhZG8iOjE2NzU3NTE1ODgsInBwbnQiOjE2Nzc2MDA1MjgsImJlZXMiOjE2Nzc2MDA1MjgsInNtYXJ0IjoxNjc3NjAwNTI4fQ%3D%3D; tpc_a=097e23659d52498e892d4ff38be55626.1674965027.hlA.1677782200; cto_bidid=ujOXj19Ea09FYVJJUXRFeFZPVGVGbE9OVWc5S0F5QjJObGNnWFBtU2FPdVpMOGJIcW4wSWM2a2RxV3kyRjJqUTV6ZTBiWk1YMTBxYUtqWVFQRnduakZ5NTFrV3UxenNYeGdQQXlFU3B4eElObUJaNCUzRA; SusyAnalyticsCookie=fe5f6d30-cff4-11ed-94d3-fb5a2f6d9770.1680288666499; pbjs-unifiedid=%7B%22TDID%22%3A%22d4047405-adfc-4196-9238-d84993744b53%22%2C%22TDID_LOOKUP%22%3A%22TRUE%22%2C%22TDID_CREATED_AT%22%3A%222023-03-05T20%3A24%3A18%22%7D; pbjs-unifiedid_last=Wed%2C%2005%20Apr%202023%2020%3A24%3A19%20GMT; cto_bundle=a710Z195ZE00aTFZZFY2MGFPdzdBTHN6Sm5tN1NGWElvMWslMkJjRmVUNkw3VnA2JTJGOU8lMkJQTUJuc1B1aVBBbmdqUnpiWkxqTm1VQjVpUUt5RHMza2Z0VzVWTjZQTjJmcDI3U2gwMTRMaTNOelY2a2hwcyUyRkRJTThRQ2tORGVIZW9Tako5MDhTMDdEaWZVUkJRNHFtY1RHMzQ2aVdSdyUzRCUzRA; _cc_id=6056d91fe263be7d78d2246d0a91177b; cdn.studymodellc.126400.ka.ck=426ede393098d49afa870746c9c9285d9327d956d3e3fc625e0e6913b157c1cf69616134b4866bb90633eb8198bb2d8cf39d571d179553685e99e954a85a505039f99df4d2469f9eb9fbe0b1a063bc85947715b47a3e010d7e14f5cc948ef3b6ddc365dc0da090266122642a11dfd2abc47d95863003a61636c7e741dbb65e81ee057e89943b3668ee46d6a3bc893c10d8582f46a40095d1738fe8; splat_user=%7B%22userId%22%3A%22406156299%22%2C%22isPremium%22%3Afalse%7D; _ga=GA1.2.1772280364.1674776205; _ga_HXSFB2SGLZ=GS1.1.1680726251.13.1.1680729328.0.0.0; __gpi=UID=00000badb83cb622:T=1674776206:RT=1680970601:S=ALNI_MbvNAqVLaGB4tI3LCrqMvhWdkBEOg; __attentive_dv=1; ipqsd=172748559920147000; _hjSession_1400488=eyJpZCI6IjViOTZlZmY5LWI5ZjctNGM2OS04M2I2LTI3N2ExMzFjZTIyMCIsImNyZWF0ZWQiOjE2ODEyMDk0OTM4ODIsImluU2FtcGxlIjp0cnVlfQ==; __attentive_ss_referrer=ORGANIC; __attentive_block=true; analytics_session_id=1681209532144; attntv_mstore_email=surensays@tutanota.com:0; refreshToken=088eab10e552d5198bd7fe259329e175922eb4a0; userId=05fdd8bd-68ee-422b-aa74-e218284ea617; userStatus=A1; promotionId=; sku=bb499firstweek_999_intl_learn_3; __attentive_pv=3; btbHomeDashboardAnimationTriggerDate=2023-04-12T10:39:38.493Z; btbHomeDashboardTooltipAnimationCount=1; btbHomeDashboardBonusChallengeModalCount=0; majors=Chemical%2520Engineering; _uetsid=ae6cd3e0d7a811eda78f4d95e0d1566e; _uetvid=4bab35709dd211ed968aa700bdfb7773; OptanonConsent=isIABGlobal=false&datestamp=Tue+Apr+11+2023+16%3A12%3A19+GMT%2B0530+(India+Standard+Time)&version=202302.1.0&hosts=&consentId=6deee5be-521b-4dbd-91cd-a16ef6f85708&interactionCount=1&landingPath=NotLandingPage&groups=C0001%3A1%2CC0003%3A1%2CC0002%3A1%2CC0005%3A1%2CC0004%3A1%2CSPD_BG%3A1&AwaitingReconsent=false&isGpcEnabled=0&geolocation=IN%3BDL; OptanonAlertBoxClosed=2023-04-11T10:42:19.848Z; analytics_session_id.last_access=1681209761951; accessToken=3221f296be320451c082a0003e5d6a1a2496a71b; bartlebyRefreshTokenExpiresAt=2023-05-11T10:55:49.756Z; ABTastySession=mrasn=&lp=https%253A%252F%252Fwww.bartleby.com%252Flogin; ABTasty=uid=wx9ekahgyw2y215a&fst=1674776205745&pst=1681202631482&cst=1681209475100&ns=122&pvt=412&pvis=5&th=831945.1034156.1.1.1.1.1678371591601.1678371591601.1.74_919607.1146677.37.27.2.1.1678022104184.1679163222365.1.92'
                    }
            print("url bartleby")
            r = requests.get(str(qurl), headers=headers, data=payload)
            soup = s(r.content, 'html.parser')
            print(r)
            element = soup.find(id="__NEXT_DATA__")
            json_data2 = json.loads(element.text)
            # Do something with the imgTag, such as print it
            try:
               imgTag = json_data2['props']['pageProps']['questionAnswer']['question']['images'][0]['imageUrl']
               print("image found")
            except:
               imgTag = "none"
               print("img not found")  
            try:
                subject = json_data2['props']['pageProps']['questionAnswer']['question']['subjects'][0]['title']
                subject1 = json_data2['props']['pageProps']['questionAnswer']['question']['subjects'][1]['title']
                topic = json_data2['props']['pageProps']['questionAnswer']['question']['topics'][0]['title']
                idno = json_data2['props']['pageProps']['questionAnswer']['question']['id']
                qu = json_data2['props']['pageProps']['questionAnswer']['question']['selectedText']
                qu2 = json_data2['props']['pageProps']['questionAnswer']['question']['text']
            except :
                subject = 'none'
                subject1 = 'none'
                topic = 'none'
                idno = 'none'
                qu = 'none'
                qu2 = 'none'

            for step in json_data2['props']['pageProps']['questionAnswer']['answer']['steps']:
                text22 = step['text']
                f = open('DDbb.html', 'a')
                f.write(str(text22))
                f.close()
            # Do something with the text, such as print it
            f2 = open('DDbb.html', 'r')
            anshtml = str("""<!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>NX pro</title>
            <meta name="description" content="">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <link rel="shortcut icon" type="image/x-icon" href="assets/img/favicon.ico">
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bulma/0.9.3/css/bulma.min.css">
            <script src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/3.2.0/es5/tex-mml-chtml.min.js"></script>
            </head>
            <body>
            <div class="container">
            <div id="app">
            <div class="container">
            <div class="section">
            <div class="box" style="word-break: break-all;">
            <h1>Question Link</h1>
            <div class="url">"""+str(qurl)+"""</div></div>
            <div class="box" style="word-break: break-all;">
            <div class="query">"""+str(url)+"""</div></div>
            <div class="box">
            <div class="content">
            <div class="subject"><h2>"""+str(subject)+"""</h2></div>
            <div class="subject1"><h2>"""+str(subject1)+"""</h2></div>
            <div class="topic"><h2>"""+str(topic)+"""</h2></div>
            </div></div>
            <div class="box">
            <div class="content">
            <h1>Question</h1>
            <div class="questionnx">"""+str(qu)+"\n"+str(qu2)+"\n""""<img src="""+str(imgTag)+"""></div>
            </div>
            </div>
            <div class="box">
            <div class="content">
            <h1>Answer</h1>
            <div class="answernx">"""+f2.read()+"""</div></div>
            </div>
            </div>
            </div>
            </div>
            </div>
            <script type="text/x-mathjax-config">MathJax.Hub.Config({ config: ["MMLorHTML.js"], jax: ["input/TeX","input/MathML","output/HTML-CSS","output/NativeMML"], extensions: ["tex2jax.js","mml2jax.js","MathMenu.js","MathZoom.js"], TeX: { extensions: ["AMSmath.js","AMSsymbols.js","noErrors.js","noUndefined.js"] } });</script> <script type="text/javascript" src="https://cdn.mathjax.org/mathjax/2.0-latest/MathJax.js?config=TeX-AMS-MML_HTMLorMML"></script></body>
            </html>""")
            folder_path = 'allbart3'
            file_name = 'Answer_{}.html'.format(idno)
            file_path = os.path.join(folder_path, file_name)
            f = open(file_name, 'w')
            f.write(str(anshtml))
            f.close()
            i = open(file_name, 'rb')
            bot.send_document(2110818173, i ,parse_mode='Markdown')
            os.remove(file_name)
            os.remove("DDbb.html")
            # Store the URLs in text.txt and qurl.txt
            with open("text.txt", "a") as f:
                f.write(str(url) + "\n")

            with open("qurl.txt", "a") as f:
                f.write(str(qurl) + "\n")
            
        # Update the last processed row after every iteration
        last_processed_row = i2
        with open("last_processed_row.txt", "w") as f:
            f.write(str(last_processed_row))

    except Exception as e:
        print(f"Error processing row {i2}: {e}")
        with open("error2.txt", "a") as f:
            f.write(str(url) + "\n")
        with open("error_qurl.txt", "a") as f:
            f.write(str(qurl) + "\n")
        continue
