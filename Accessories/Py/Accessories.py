from Scraping.Accessories.Py.poorvika import Model
###################################################################################################


def scraping_accessories():
    worksheet_name = "All Accessories "
    Title = "Accessories Brand"
    SV = 1
    Acc_brand = "https://www.poorvika.com/accessories-brands?&stock=1"
    Acc_brand_list = "https://www.poorvika.com/accessories-brands?&fmin=99&fmax=77999&stock=1&order=l-h&page="
    acc_brand = Model(url=Acc_brand,url_list=Acc_brand_list,row_num=SV,title=Title,wb_title=worksheet_name)
    acc_brand.product()

    ###################################################################################################

    Title = "Mobile Accessories"
    AB = acc_brand.product_num()
    Mob_Acc = "https://www.poorvika.com/mobiles-accessories?&stock=1"
    Mob_Acc_list = "https://www.poorvika.com/mobile-accessories?&fmin=99&fmax=77999&stock=1&order=l-h&page="
    mob_acc = Model(url=Mob_Acc,url_list=Mob_Acc_list,row_num=AB,title=Title,wb_title=worksheet_name)
    mob_acc.product()

    ####################################################################################################

    Title = "Smart Technology"
    MA = mob_acc.product_num()
    Smart_Tech = "https://www.poorvika.com/smart-technology?&stock=1"
    Smart_Tech_List = "https://www.poorvika.com/smart-technology?&fmin=299&fmax=77900&stock=1&order=l-h&page="
    smart_tech = Model(url=Smart_Tech,url_list=Smart_Tech_List,row_num=MA,title=Title,wb_title=worksheet_name)
    smart_tech.product()

    ####################################################################################################

    Title = "Audio"
    ST = smart_tech.product_num()
    Audio = "https://www.poorvika.com/audio?&stock=1"
    Audio_List = "https://www.poorvika.com/audio?&fmin=99&fmax=77999&stock=1&order=l-h&page="
    audio = Model(url=Audio,url_list=Audio_List,row_num=ST,title=Title,wb_title=worksheet_name)
    audio.product()

    #####################################################################################################

    Title = "Personal_Health_Care"
    AD = audio.product_num()
    Personal_Health = "https://www.poorvika.com/personal-health-care?&stock=1"
    per_hel = Model(url=Personal_Health,url_list=None,row_num=AD,title=Title,wb_title=worksheet_name)
    per_hel.product1()
    PH = per_hel.product_num()

    #####################################################################################################

    print("Accessories Brand", AB)
    print("Mobile Accessories", MA-AB)
    print("Smart Technology", ST-MA)
    print("Audio", AD-ST)
    print("Personal Health", PH-AD)
    print("Total value", PH)

    #####################################################################################################



