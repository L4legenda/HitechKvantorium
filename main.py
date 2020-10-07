
from openpyxl import load_workbook, drawing, Workbook
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

list1 = wb_form["Лист1"]
cols = "ABCDEFGHIJKLMNO"
ws = wb_form.active

lists = {}

for c in cols:
    for r in range(1, 17):
        val = None
        style = None

        try:
            val = ws[c + str(r)].value
            style = ws[c + str(r)].style
        except:
            print(c, r, "Не доступна")

        lists[c + str(r)] = {
            "val": val,
            "style": style
        }
print(lists)

for l in lists:
    id = l[0] + str( int(l[1:]) + 16 )
    try:
        wb_form["Лист1"][id] = lists[l]["val"]
        list1[id] = lists[l]["val"]
    except:
        print(id, "Ошибка")
print("Good")

#
#my_png = drawing.image.Image("qr/qrcode.png")
#ws.add_image(my_png, "K3")


wb_form.save("export.xlsx")