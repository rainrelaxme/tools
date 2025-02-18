由于国际化翻译，低代码部分无法利用工具快速翻译，考虑提取其中中文替换。
1. 先将文件（input.txt）放到文件夹下。
2. 运行word2excel.py,得到output.xlsx.
3. 将output.xlsx拷贝到桌面，拷贝一份output-EN.xlsx.
4. 将output.xlsx上传到google translate（https://translate.google.com/?hl=zh-cn&sl=auto&tl=en&op=docs ），翻译并下载文件。
5. 将下载的文件的英文列，复制到output-EN.xlsx的第三列，并将此文件拷贝到程序目录下。
6. 运行excel2word.py，将其中输入文件excel_file=output-EN.xlsx,执行得到output.txt,replaced_results.xlsx。
7. 此时output.txt已经可以替换为对应的语言。

8. 可以再执行check_CN.py，输入文件为output.py，得到output-EN.xlsx，以此检查上一个程序是否有遗漏的中文。