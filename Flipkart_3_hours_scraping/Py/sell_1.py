from Scraping.Flipkart_3_hours_scraping.Py.seller import Filpkart_Scraping_3_hours
import datetime

date = datetime.datetime.now().strftime("%d-%m-%Y %I %p")


fk = Filpkart_Scraping_3_hours()
fk.flikart_seller(start=1,end=271,path=1)

