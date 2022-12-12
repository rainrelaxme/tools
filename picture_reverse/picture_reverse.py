from PIL import Image, ImageSequence
import os
 
# 用于读取gif动图
im = Image.open("D:\learnpython\changzhang.gif")
# gif图片流的迭代器
iter = ImageSequence.Iterator(im)
 
index = 1
 
for frame in iter:
    # 每一帧图片
    print("image %d: mode %s, size %s" % (index, frame.mode, frame.size))
    # 若保存的路径不存在，则创建文件夹
    path = "imgs"
    if path not in os.listdir():
        os.makedirs(path)
    # 将每一帧图片保存到imgs文件夹下
    frame.save("./imgs/frame%d.png" % index)
    index += 1
 
# 将gif拆分成图片流
imgs = [frame.copy() for frame in ImageSequence.Iterator(im)]
 
# 输出原图
# imgs[0].save("./out.gif", save_all=True, append_images=imgs[1:])
 
# 将图片流反序
imgs.reverse()
 
# 将反序后的图片流保存并输出
imgs[0].save("./reverse_out.gif", save_all=True, append_images=imgs[1:])