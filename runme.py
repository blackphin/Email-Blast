import xlrd, smtplib


email_addres = "ENTER EMAIL HERE"
email_password = "ENTER PASSWORD HERE"
subject = "ENTER SUBJECT LINE HERE"


wb = xlrd.open_workbook(r"D:\Email Blast\College List.xls")
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

file = open(r"D:\Email Blast\message.txt", "r")
message = file.read()

with smtplib.SMTP("smtp.gmail.com") as connection:
    connection.starttls()
    connection.login(user=email_addres, password=email_password)
    for i in range(sheet.nrows):
        college_name = sheet.cell_value(i, 0)
        to_email_address = sheet.cell_value(i, 1)
        connection.sendmail(
            from_addr=email_addres,
            to_addrs=to_email_address,
            msg="Subject:" + subject + "\n\n" + "Dear " + college_name + "\n" + message,
        )
        print(i + ". Email sent to " + college_name + " successfully")
