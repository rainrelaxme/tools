import time
from tkinter import filedialog

from modules.create_barcode.create_barcode import EAN_list

CURRENT_TIME = time.time()  #获取当前时间float时间戳

open_file()

"""
# 文件写入内容
fd = os.open("./1.txt", os.O_RDWR)
os.write(fd, str(CURRENT_TIME).encode())  #要求byte_like object,str和byte_like可以用encode和decode转化
os.close(fd)
print("It has been closed !")
"""

def open_file():
  #打开多个文件
  file_path = filedialog.askopenfiles()
  for i in file_path:
    f = open(i, "r+")
    while True:
      line = f.readline() #读文件
      if line == '':
        f.close()
        break
      else:
        EAN_list.append(line)
        print(line)