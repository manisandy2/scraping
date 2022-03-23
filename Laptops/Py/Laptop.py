from Scraping.Accessories.Py.poorvika import Model
###################################################################################################


def scraping_laptop():
    worksheet_name = "All Laptops "
    title = "Laptop"
    lop = 2
    laptop = "https://www.poorvika.com/laptops-computers"
    laptop_list = "https://www.poorvika.com/laptops-computers?&fmin=21990&fmax=329900&order=l-h&page="
    laptop_brand = Model(url=laptop,url_list=laptop_list,row_num=lop,title=title,wb_title=worksheet_name)
    laptop_brand.product()


