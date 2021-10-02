from selenium import webdriver
import smtplib
from email.message import EmailMessage
from openpyxl import load_workbook
#Web scraping for "yad2" website
#Make sure your link includes your preferences such as price and area

driver = webdriver.Chrome(r'dir.exe')
driver.get(
    'https://www.yad2.co.il/realestate/rent?topArea=2&area=1&city=5000&neighborhood=1520&rooms=3-3&price=6500-8700')
driver.maximize_window()
driver.implicitly_wait(300)

ids = []
#Loading excel data base
wb = load_workbook('yad.xlsx')
ws = wb.active
#Loading existing apts
L = [i.value for i in ws['A'] if i.value]


def my_app_bot(ids_of_apt):
    global L
    global ids
    global ws
    global wb

    driver.refresh()
    new_apts = []
    scrolling = 400
    for i in range(0, 15):
        scrolling += 400
        print(i)
        iStr = str(i)

        if driver.find_element_by_id('feed_item_' + iStr + '_price'):
            apt_id = driver.find_element_by_id('feed_item_' + iStr + '_price')
            apt_id.click()
            ad_num = driver.find_element_by_class_name('num_ad').text[-8:]
            driver.execute_script("window.scrollTo(0, %s)" % str(scrolling))

            if ad_num not in ids_of_apt:
                apt_link = driver.find_element_by_xpath(
                    """//*[@id="accordion_wide_%s"]/div/div[3]/ul/li[4]/a""" % iStr).get_attribute('href')
                new_apts.append(apt_link)
                ws["A"+str(ws.max_row+1)] = ad_num
                wb.save('yad.xlsx')

        else:
            print('no')
        apt_id.click()

    if new_apts:
        #send email
        sender = "sender@gmail.com"
        rec = "reciver@gmail.com"
        password = 'password'
        mes = '''
        '''.join(new_apts)
        msg = EmailMessage()
        msg.set_content(mes)

        msg['Subject'] = 'New Apartments!'
        msg['From'] = sender
        msg['To'] = rec

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        print('email sent')
    my_app_bot(L)


my_app_bot(L)

