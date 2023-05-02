import requests
from bs4 import BeautifulSoup  as s
import json
import openpyxl
#import urllib.parse
import os.path
import telebot
#telgram bot detail
TOKEN = "6002286894:AAEB3FK-M41QnPxSzPdjN-0biGV16UsBMb8"
bot = telebot.TeleBot(TOKEN, parse_mode=None, threaded=True)
# Define the range of rows to process
start_row = 1
end_row = 75000

# Load the last processed row from a file
try:
    with open("last_processed_row.txt", "r") as f:
        last_processed_row = int(f.read().strip())
except FileNotFoundError:
    last_processed_row = start_row

# Load the Excel spreadsheet
wb = openpyxl.load_workbook('68894.xlsx')
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
                i = open("url_not_found.txt", 'rb')
                bot.send_document(-1001534695986, i ,parse_mode='Markdown')
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
                'Cookie':'__attentive_id=cd1b4fb9a28c4a2c96cacadab99b4e4d; _attn_=eyJ1Ijoie1wiY29cIjoxNjgyNTM1Nzc4MzE4LFwidW9cIjoxNjgyNTM1Nzc4MzE4LFwibWFcIjoyMTkwMCxcImluXCI6ZmFsc2UsXCJ2YWxcIjpcImNkMWI0ZmI5YTI4YzRhMmM5NmNhY2FkYWI5OWI0ZTRkXCJ9In0=; __attentive_cco=1682535778322; __attentive_ss_referrer=https://www.google.com/; _gcl_au=1.1.279222654.1682535779; _hjFirstSeen=1; _hjIncludedInSessionSample_1400488=1; _hjSession_1400488=eyJpZCI6IjJhYTA0ZmM4LTQ0NmUtNDRlMC1hZDJkLWY1NGY1MGEwNGVkOSIsImNyZWF0ZWQiOjE2ODI1MzU3NzkzNTIsImluU2FtcGxlIjp0cnVlfQ==; __attentive_dv=1; _gid=GA1.2.1487647807.1682535780; _gat_UA-20169249-1=1; _tt_enable_cookie=1; _ttp=z7Las12Az8lHqB5QB9OOetScXAK; _fbp=fb.1.1682535781008.1138832355; G_ENABLED_IDPS=google; _hjSessionUser_1400488=eyJpZCI6ImZiZTBjNmE0LTE0ZDQtNTljYi04ZDVmLWNmZDg0ZDhmMTA5ZSIsImNyZWF0ZWQiOjE2ODI1MzU3NzkzMzgsImV4aXN0aW5nIjp0cnVlfQ==; ajs_anonymous_id=6f6ee3e5-c48f-49e0-ab07-989f3560192b; analytics_session_id=1682535838858; attntv_mstore_email=surensays@tutanota.com:0; refreshToken=afd283eb140d9e79236c13a2c41d0aca568169bc; userId=05fdd8bd-68ee-422b-aa74-e218284ea617; userStatus=A1; promotionId=; sku=bb499firstweek_999_intl_learn_3; accessToken=e95030b9fe7cbfb23a7aa46fcd349ff0dffde1a8; bartlebyRefreshTokenExpiresAt=2023-05-26T19:04:11.477Z; category=Learn; btbHomeDashboardAnimationTriggerDate=2023-04-27T19:04:15.479Z; btbHomeDashboardTooltipAnimationCount=1; _gat_UA-93748-2=1; viewedPagesCount=2; isNoQuestionAskedModalClosed=true; majors=Chemical%20Engineering; fs.bot.check=true; __attentive_block=true; __gads=ID=6c8034578fdea643-2256a4b1bbdf0085:T=1682535880:RT=1682535880:S=ALNI_MYCs_U5qn5VKjxVxRZHvf0cLcIEOg; __gpi=UID=00000bfdf4ff5977:T=1682535880:RT=1682535880:S=ALNI_MYFSXu2VjRDt5OJagyW4-NYattq2g; _pbjs_userid_consent_data=3524755945110770; analytics_session_id.last_access=1682535885909; cookie=c187d453-3d1b-4f75-ace4-b021e6b2a21f; _lr_retry_request=true; _lr_env_src_ats=false; __qca=P0-1980189146-1682535887216; ABTastySession=mrasn=&lp=https%3A%2F%2Fwww.bartleby.com%2F; ABTasty=uid=vty81frbzfgkwgkf&fst=1682535779497&pst=-1&cst=1682535779497&ns=1&pvt=7&pvis=7&th=; _uetsid=f62df460e46411eda2d8a74ae2fb4f03; _uetvid=f6307a70e46411eda5bf97d7dea356a0; _ga=GA1.2.734742159.1682535780; _ga_R3RTBJZFE8=GS1.1.1682535779.1.1.1682535898.0.0.0; __attentive_pv=7; OptanonAlertBoxClosed=2023-04-26T19:04:58.751Z; OptanonConsent=isGpcEnabled=0&datestamp=Thu+Apr+27+2023+00:34:59+GMT+0530+(India+Standard+Time)&version=202302.1.0&isIABGlobal=false&hosts=&consentId=ca60fbd5-c588-4768-afe7-1fde2559dc6a&interactionCount=1&landingPath=NotLandingPage&groups=C0001:1,C0003:1,SPD_BG:1,C0004:1,C0005:1,C0002:1&AwaitingReconsent=false&geolocation=IN;DL'
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
            bot.send_document(2110818173, i,caption=str("Row:"+str(i2)) ,parse_mode='Markdown')
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
            i = open("error2.txt", 'rb')
            try:
                bot.send_document(-1001534695986, i ,parse_mode='Markdown')
                print(f"Bot Send error file!")
            except Exception as e:
                print(f"Error processing row : {e}")
                continue
        #with open("error_qurl.txt", "a") as f:
            #f.write(str(qurl) + "\n")
        continue
