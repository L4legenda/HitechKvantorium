
from openpyxl import load_workbook, drawing
import qrcode

wb_form = load_workbook("data/data.xlsx")

seria = input("Серия: ")
num = int(input("Номер: ").strip())
i = num
money = input("Цена: ")

qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=7,
    border=1,
)
qr.add_data(input("QR code: "))
qr.make(fit=True)

img = qr.make_image(fill_color="black", back_color="white")
img.save("qr/qrcode.png")

wb_form["Лист1"]["I7"]  = seria
wb_form["Лист1"]["I23"] = seria
wb_form["Лист1"]["I39"] = seria
wb_form["Лист1"]["I55"] = seria
wb_form["Лист1"]["Y7"]  = seria
wb_form["Лист1"]["Y23"] = seria
wb_form["Лист1"]["Y39"] = seria
wb_form["Лист1"]["Y55"] = seria

wb_form["Лист1"]["D10"] = seria
wb_form["Лист1"]["D26"] = seria
wb_form["Лист1"]["D42"] = seria
wb_form["Лист1"]["D58"] = seria
wb_form["Лист1"]["T10"] = seria
wb_form["Лист1"]["T26"] = seria
wb_form["Лист1"]["T42"] = seria
#Nums
wb_form["Лист1"]["D2"] = i
wb_form["Лист1"]["I9"] = i
i += 1
wb_form["Лист1"]["D18"] = i
wb_form["Лист1"]["I25"] = i
i += 1
wb_form["Лист1"]["D34"] = i
wb_form["Лист1"]["I41"] = i
i += 1
wb_form["Лист1"]["D50"] = i
wb_form["Лист1"]["I57"] = i
i += 1
wb_form["Лист1"]["T2"] = i
wb_form["Лист1"]["Y9"] = i
i += 1
wb_form["Лист1"]["T18"] = i
wb_form["Лист1"]["Y25"] = i
i += 1
wb_form["Лист1"]["T34"] = i
wb_form["Лист1"]["Y41"] = i
i += 1
wb_form["Лист1"]["T50"] = i
wb_form["Лист1"]["Y57"] = i

wb_form["Лист1"]["E2"] = money
wb_form["Лист1"]["E18"] = money
wb_form["Лист1"]["E34"] = money
wb_form["Лист1"]["E50"] = money
wb_form["Лист1"]["U2"] = money
wb_form["Лист1"]["U18"] = money
wb_form["Лист1"]["U34"] = money
wb_form["Лист1"]["U50"] = money

ws = wb_form.active
my_png = drawing.image.Image("qr/qrcode.png")
ws.add_image(my_png, "K3")
ws.add_image(my_png, "K19")
ws.add_image(my_png, "K35")
ws.add_image(my_png, "K51")
ws.add_image(my_png, "AA3")
ws.add_image(my_png, "AA19")
ws.add_image(my_png, "AA35")
ws.add_image(my_png, "AA51")

wb_form.save("export.xlsx")