from pystrich.qrcode import QRCodeEncoder
from pystrich.code128 import Code128Encoder
from pystrich.code39 import Code39Encoder
from pystrich.ean13 import EAN13Encoder
from pyzbar.pyzbar import decode
from PIL import Image
import os

USER = "古内翔平"
#SAVE_PATH = "C:/Users/{}/Desktop/".format(USER)
SAVE_PATH = "C:/Users/{}/OneDrive - 株式会社　ＴＲＩＮＩＴＹ　/業務会議/1⃣管理部/社員バーコード/".format(USER)


# encoder = QRCodeEncoder("600300563001")
# encoder.save("01.png")

# encoder = EAN13Encoder("600300563001")
# encoder.save("02.png")
code = 1024

Data1 = "00{}".format(code)
Data2 = "古内 翔平"
Image_name = str(code) + ".png"
CreatePath = os.path.join(SAVE_PATH,Image_name)
print(CreatePath)
encoder = Code39Encoder(str(Data1) , options={"show_label": True})

encoder.save(CreatePath,bar_width=2)

print(decode(Image.open(CreatePath))[0].data.decode())
