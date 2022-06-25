#将条形码生成图片

import os, datetime, time
import barcode
from barcode.writer import ImageWriter

EAN_list = []
f = open('1.txt', 'r+') #读文件
while True:
  line = f.readline()
  if line == '':
    f.close()
    break
  else:
#    line = eval(line)
#    EAN_list.append(str(line))
    EAN_list.append(line)

# currentTime = datetime.date.today()
currentTime = time.time()
os.mkdir("./images/" + str(currentTime) + "输出")

x = 1
for i in EAN_list:
  EAN = barcode.get_barcode_class("code128")
  ean = EAN(i, writer=ImageWriter)
  #去除特殊字符
  j = i.replace("\n", "")
  k = j.replace(":", "")
  #输出文件夹
  route = ean.save("./images/" + str(currentTime) + "输出/" + str(k) + "barcode")
  x = x + 1
  print("第" + str(x) + "个已成功，输出文件：" + route)


"""
EAN = barcode.get_barcode_class("code128")

message = "12312313"
ean = EAN(message, writer = ImageWriter())
fullname = ean.save("images/条形码")
print("success")
"""