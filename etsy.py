import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep, time
from selenium.webdriver.support.select import Select
import pyautogui as pg
import pandas as pd
import smtplib
import xlrd

#if any exception is happened, this function sends mail with short explanation
import pillow_config


def send_mail(imagesNotFound=None, product=None, sku_num=None):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login("hof2018hof@gmail.com", "hof12345")
    if imagesNotFound is None and product is None and sku_num is None:
        msg = "Hata!"
    elif product is None and sku_num is None:
        msg = "Hata! \n Image List Number : " + str(imagesNotFound)
    elif sku_num is None:
        msg = "Hata! \nImage List Number : " + str(imagesNotFound) + " Product : " + str(product)
    else:
        msg = "Hata! \nImage List Number : " + str(imagesNotFound) + " Product : " + str(product) + " SKU : " + str(sku_num)

    server.sendmail("hof2018hof@gmail.com", "hasanemreari@gmail.com", msg)
    print(msg)
    server.quit()
########################################################################################################################
#function that creates tags from title and default tags
def creatingTags(elem):
    defaultTagList=['Decorative Pillow','Handmade Pillow','Cushion Cover']
    title = str(elem).split(" ")
    #print(title[0])
    a=""
    if str(title[0]).__contains__('"'):
        for i in title[0]:
            if i=='"':
                print("Tags are creating...")
            else:
                a+=i
        #print(True)
    #print(title)
    title.remove(title[0])
    title.insert(0,a)
    if(title.__len__()<8):
        for i in title:
            defaultTagList.append(i)
    else:
        for j in range(0,8):
            defaultTagList.append(title[j])

    if defaultTagList.__contains__("Pillow"):
        defaultTagList.remove('Pillow')
    if defaultTagList.__contains__("and"):
        defaultTagList.remove("and")

    tagSet=set(defaultTagList)
    #print(tagSet)

    stringTags=""
    for tag in tagSet:
        stringTags+=tag+","

    originalTags=stringTags[:-1]
    #print(stringTags)
    #print(originalTags)
    return originalTags

def modifyCategory(category):
    # nonpillow=category.split('>')

    # print(nonpillow)
    cat=category
    # print(cat)
    b = str(cat).split('(')
    inche=(b[0])
    # print(inche)

    char=[]
    catList=[]
    for i in inche:
        char.append(i)
    for t in cat:
        catList.append(t)

    if char.__contains__('\xa0'):
        char.remove('\xa0')
    if char.__contains__('\xa0'):
        char.remove('\xa0')
    if char.__contains__(' '):
        char.remove(' ')
    if char.__contains__(' '):
        char.remove(' ')

    if catList.__contains__('\xa0'):
        catList.remove('\xa0')
    if catList.__contains__('\xa0'):
        catList.remove('\xa0')
    if catList.__contains__(' '):
        catList.remove(' ')
    if catList.__contains__(' '):
        catList.remove(' ')

    result1 = ''
    result2 = ''
    for j in char:
        result1+=j

    for s in catList:
        result2 +=s
    return (str(result1),str(result2))

def getInches(input):

    return (str(input[0:2]),str(input[4:6]))

#####################Editlenecekler#######################################################################
file = pillow_config.file
dir_input = pillow_config.dir_input

excel = pd.read_excel(file, index_col=None, header=None)
total_number_of_row = excel.count().iloc[0]

#################Open Browser ########################

browser = webdriver.Chrome()
browser.get("https://www.etsy.com/your/shops/me/dashboard?ref=mcpa")
browser.maximize_window()

#################Login#######################
try:
    elem = browser.find_element_by_id("join_neu_email_field")
    #print("Test Pass : Email ID found")
except Exception as e:
    print("Exception found on LOGIN "+str(e))

elem.send_keys("hetyemez@yahoo.com")

try:
    elem = browser.find_element_by_id("join_neu_password_field")
    #print("Test Pass : Password ID found")
except Exception as e:
    print("Exception found on LOGIN "+ str(e))
elem.send_keys("etyemez57")
elem.send_keys(Keys.ENTER)

print("Login is done successfully")
#########################################Product########################################
##########################################Entry#########################################

total_number_of_row = pillow_config.endRow    # hem dosyadan okuyo hem de biz giriyoruz. Düzelince hallet
row = pillow_config.startRow
imagesNotFound=[]
try:
    while row < total_number_of_row:

        ###############Gerekli bilgiler dosyadan okunarak alındı###########################
                       # Burası düzeltilecek !!!!!!!!!!!!!!!!!!
        sku_num = excel[0][row]
        materials = excel[2][row]
        price = excel[4][row]
        title2 = excel[8][row]
        section = excel[10][row]
        title = str(section[0:7]) + ' ' + title2
        inches = getInches(title)

        modifiedCategoryTuple = modifyCategory(section)


        description = "Vintage kilim pillow cover. In good condition.\n\n" + "Usually shipped via Fedex or UPS.\n" + "Item will be ready for the shipment within 1-3 business days with a tracking number.\n" + "Estimated delivery 3-7 business days.\n" + "Properly washed and ready to use.\n\n" + "Please contact us if you need additional information or have any questions.\n\n" + "Thanks\n\n" + str(
            sku_num)
        sleep(3)
        browser.get("https://www.etsy.com/your/shops/WovenHane/tools/listings/create")
        sleep(3)
        print("*******************************Ürün Bilgileri Alındı********************************************")
##################Resim dosyada var mı yok mu kontrol et. Varsa yükle. yoksa loga bunu yazıp. Yüklemeye devam et.################
        ####################    Resim 1 #####################
        try:
            elem = browser.find_element_by_name("image-upload")
            print("Test Pass : Image ID1 found successfully")
        except Exception as e:
            print("Exception found at FINDING IMAGE ID 1" + str(e))

        if os.path.isfile(dir_input + str(sku_num) + ".jpg"):
            elem.click()
            sleep(1)
            pg.typewrite(dir_input + str(sku_num) + ".jpg")  # "C:\\Users\\hasan\\Desktop\\deneme\\a.jpg")
        else:
            imagesNotFound.append(sku_num)
            print("Bu ürünün resmi bulunamadı 1: " + str(sku_num)+ " Satır numarası: "+ str(row))
            row = row + 1
            print("*******************************Sıradaki ürüne geçildi********************************************")
            continue
        sleep(2)
        pg.press("enter")
        sleep(0.5)
        ####################    Resim 2 #####################
        try:
            elem = browser.find_element_by_name("image-upload")
            print("Test Pass : Image ID2 found successfully")
        except Exception as e:
            print("Exception found at FINDING IMAGE ID 2" + str(e))

        elem.click()
        sleep(1)
        if os.path.isfile(dir_input + str(sku_num) + "a.jpg"):
            pg.typewrite(dir_input + str(sku_num) + "a.jpg")  # "C:\\Users\\hasan\\Desktop\\deneme\\a.jpg")
        else:
            imagesNotFound.append(sku_num)
            print("Bu ürünün resmi bulunamadı 2: " + str(sku_num) + " Satır numarası: " + row)
            print("*******************************Sıradaki ürüne geçildi********************************************")
            row = row + 1
            continue
        sleep(2)
        pg.press("enter")
        sleep(0.5)
        ####################    Resim 3 #####################
        try:
            elem = browser.find_element_by_name("image-upload")
            print("Test Pass : Image ID3 found successfully")
        except Exception as e:
            print("Exception found at FINDING IMAGE ID 3" + str(e))
        elem.click()
        sleep(1)
        if os.path.isfile(dir_input + str(sku_num) + "b.jpg"):
            pg.typewrite(dir_input + str(sku_num) + "b.jpg")  # "C:\\Users\\hasan\\Desktop\\deneme\\a.jpg")
        else:
            imagesNotFound.append(sku_num)
            print("Bu ürünün resmi bulunamadı 1: " + str(sku_num) + " Satır numarası: " + row)
            print("*******************************Sıradaki ürüne geçildi********************************************")
            row = row + 1
            continue
        sleep(0.5)
        pg.press("enter")
        sleep(0.5)
##################  Title entry ####################################
        try:
            elem = browser.find_element_by_id("title")
            print("Test Pass : Title ID found successfully")
        except Exception as e:
            print("Exception found at TITLE ENTRY" + str(e))
        elem.send_keys(title)
        ##################  Who made Entry ####################################
        dropDownId = 'who_made'
        try:
            elem = browser.find_element_by_id(dropDownId)
            print("Test Pass : Who made ID found successfully")
        except Exception as e:
            print("Exception found at WHO MADE" + str(e))
        Select(elem).select_by_visible_text("A member of my shop")
        ##################  A finished product entry ####################################
        dropDownId = 'is_supply'
        try:
            elem = browser.find_element_by_id(dropDownId)
            print("Test Pass : What is it ID found successfully")
        except Exception as e:
            print("Exception found at IS SUPLY" + str(e))
        Select(elem).select_by_visible_text("A finished product")
        ################## 2010 - 2019 entry ####################################
        dropDownId = 'when_made'
        try:

            elem = browser.find_element_by_id(dropDownId)

            print("Test Pass : When made ID found successfully")
        except Exception as e:
            print("Exception found at WHEN MADE" + str(e))
        Select(elem).select_by_visible_text("2010 - 2019")
        ################## Decorative Pillow Entry entry ####################################
        try:
            elem = browser.find_element_by_id('taxonomy-search')

            print("Test Pass : taxonomy search ID found successfully")
        except Exception as e:
            print("Exception found at DECORATİVE PİLLOW ENTRY" + str(e))
        elem.send_keys('Decorative Pillows')
        # elem.click()
        sleep(3)
        elem.send_keys(Keys.ENTER)
        sleep(2)
        ################## 16 Inches entry ####################################
        try:
            elem = browser.find_element_by_id('attribute-55-input')
            print("Test Pass : lenght ID found successfully")
        except Exception as e:
            print("Exception found at 16 (1)" + str(e))
        sleep(1)
        elem.send_keys(inches[0])                                    #???????????????????????????????????????????????????

        dropDownId = 'attribute-55-unit'

        try:
            elem = browser.find_element_by_id(dropDownId)
            print("Test Pass : Lenght unit ID found successfully")
        except Exception as e:
            print("Exception found at INCHES (1) " + str(e))
        Select(elem).select_by_visible_text("Inches")
        ################## 16 Inches entry ####################################
        try:
            elem = browser.find_element_by_id('attribute-68-input')
            print("Test Pass : Width ID found successfully")
        except Exception as e:
            print("Exception found at 16 (2)" + str(e))
        elem.send_keys(inches[1])                                    #???????????????????????????????????????????????????
        dropDownId = 'attribute-68-unit'
        try:
            elem = browser.find_element_by_id(dropDownId)
            print("Test Pass : Width unit ID found successfully")
        except Exception as e:
            print("Exception found at INCHES (2)" + str(e))
        Select(elem).select_by_visible_text("Inches")
        ################## Description entry ####################################
        try:
            elem = browser.find_element_by_id('description')
            print("Test Pass : descriptions ID found successfully")
        except Exception as e:
            print("Exception found at DESCRIPTION" + str(e))
        elem.send_keys(description)
        ################## Section (Category) entry ####################################
        try:
            elem = browser.find_element_by_id('sections')
            print("Test Pass : sections ID found successfully")
        except Exception as e:
            print("Exception found at SECTIONS" + str(e))
        Select(elem).select_by_visible_text(section)
        ################## Tag entry ####################################
        try:
            elem = browser.find_element_by_id('tags')
            print("Test Pass : tags ID found successfully")
        except Exception as e:
            print("Exception found at TAGS" + str(e))

        elem.send_keys(creatingTags(title))
        elem.send_keys(Keys.ENTER)
        ################## Material entry ####################################
        try:
            elem = browser.find_element_by_id('materials')
            print("Test Pass : materials ID found successfully")
        except Exception as e:
            print("Exception found at MATERIAL" + str(e))

        elem.send_keys(materials) # bu kısımda exelden alınıp girilebilir !!!!!!!!!!!!!!!!!!
        elem.send_keys(Keys.ENTER)
        ################## Price entry ####################################
        try:
            elem = browser.find_element_by_name('price-input')
            print("Test Pass : price ID found successfully")
        except Exception as e:
            print("Exception found at Price ENTRY" + str(e))
        elem.send_keys(str(price))
        ################## SKU entry ####################################
        try:
            elem = browser.find_element_by_name('sku-input')
            print("Test Pass : sku ID found successfully")
        except Exception as e:
            print("Exception found at SKU" + str(e))

        elem.send_keys(str(sku_num))
        ################## PUBLISH click ####################################
        pg.press('tab')
        sleep(0.1)
        pg.press('tab')
        sleep(0.1)
        pg.press('tab')
        sleep(0.1)
        pg.press('tab')
        sleep(0.1)
        pg.press('tab')
        sleep(0.1)
        pg.press('tab')
        sleep(0.1)
        pg.press('space')

        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('tab')
        # sleep(0.1)
        # pg.press('space')


        try:
            elem = browser.find_element_by_xpath('//*[@id="page-region"]/div/div/div[3]/div/div[1]/div/div/div[2]/button[3]')
            print("Test Pass : first publish ID found")
        except Exception as e:
            print("Exception found at PUBLISH " + str(e))

        elem.click()

        try:
            elem = browser.find_element_by_xpath('//*[@id="page-region"]/div/div/div[3]/div/div[1]/div/div/div[2]/button[3]')
            print("Test Pass : first publish ID found")
        except Exception as e:
            print("Exception found at PUBLISH " + str(e))
        sleep(10)
        elem.click()
        sleep(2)
        try:
            elem = browser.find_element_by_xpath('//*[@id="overlay-region"]/div/div[3]/button[1]')
            # xpath sonundaki  button[2] ---> cancele bas ------------------------------------^
            #                  button[1] --->publishe bas ------------------------------------|
            print("Test Pass : second publish ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()
        row = row + 1
        print("*******************************Sıradaki ürüne geçildi********************************************")
    print("*************************************************************************************************************************")
    print("****************************************        ÜRÜN YÜKLEMESİ BİTTİ       ***********************************************")
    print("*************************************************************************************************************************")

except Exception as e:
    print("Son Exception found "+str(e))
    send_mail(1, row, sku_num)

send_mail(imagesNotFound,row,sku_num)
# send_mail('Ürün girişi bitti!',row,sku_num)

