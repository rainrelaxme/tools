#!/usr/bin/env python3
# -*- coding:utf-8 -*-
"""
@Project: vector_distance.py
@File   : __init__.py.py
@Version:
@Author : RainRelaxMe
@Date   : 2025/9/1 22:16
@Info   : 翻译word文件
"""


"""
由于国际化翻译，低代码部分无法利用工具快速翻译，考虑提取其中中文替换。
1. 先将文件拷贝到input.txt内（一定是这个文件名称）！ 
2. 将文件（input.txt）放到文件夹下。 
3. 运行word2excel.py,得到output.xlsx. 
4. 将output.xlsx上传到google translate（https://translate.google.com/?hl=zh-cn&sl=auto&tl=en&op=docs ），翻译并下载文件。 
5. 将下载的文件的英文列，复制到output.xlsx的第三列，并将此文件置于程序目录下。 
6. 运行excel2word.py，将其中输入文件excel_file=output.xlsx,执行得到output.txt,replaced_results.xlsx。 
7. 此时output.txt已经可以替换为对应的语言。 
8. 将output拷贝出来，按照“页面名称-语言”命名。 
9. 可以再执行check_CN.py，输入文件为output.py，得到output-EN.xlsx，以此检查上一个程序是否有遗漏的中文。
"""