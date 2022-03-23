import datetime
import pandas as pd

date = datetime.date.today().strftime("%#d-%#m-%Y")
acc_data = pd.read_excel(r"D:\Durai\Scraping\Accessories\Save File\Scraping Data\Scraping File\Accessories all Price List " + date + ".xlsx")

data = pd.DataFrame(acc_data)
print(data.columns)

all_price = pd.DataFrame(data[["Model Name",'Poorvika price','Flipkart Price','Amazon Price',
                          "Croma Price",'Vijay Sale price','Reliance Digital Price']])

print(all_price)
all_price.to_excel(r"D:\Durai\Scraping\Accessories\Save File\Scraping Data\Accessories " + date + ".xlsx",index=False)
