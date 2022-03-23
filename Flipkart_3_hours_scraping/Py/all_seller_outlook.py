import datetime


def all_seller_outlook():
    date = datetime.datetime.now().strftime("%d-%m-%Y %I %p")
    att_file = r"D:\Durai\Scraping\Flipkart_3_hours_scraping\save data\Flipkart_All_Sellers 1 " + date + ".xlsx"
    print(att_file)

    import win32com.client as client

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.To = "Rvk@poorvika.com"
    message.CC = "karthik@poorvika.in; mani2005poorvika@gmail.com; saravanavelu0482@poorvika.com; karpagam0064@poorvika.com; yasararafath1147@poorvika.com"
    message.Subject = "Flipkart Price " + datetime.datetime.now().strftime("%d-%m-%Y")
    message.Body = """Hi Team,

            Kindly find the attachment of Flipkart All Sellers list.

    With Regards,
    Duraikannan R

    +91-8682997570
    """
    message.Attachments.Add(att_file)


all_seller_outlook()