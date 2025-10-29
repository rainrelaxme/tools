import pandas as pd

# 以二进制模式读取文件，然后尝试解码
file = "C:/Users/CM001620/Desktop/匹配结果/7vs7-1.txt"
with open(file, 'rb') as file:
    raw_data = file.read()

# 尝试常见的编码
encodings = ['gbk', 'gb2312', 'utf-8', 'latin-1']
content = None

for encoding in encodings:
    try:
        content = raw_data.decode(encoding)
        print(f"成功使用 {encoding} 编码")
        break
    except UnicodeDecodeError:
        continue

if content is None:
    # 如果所有编码都失败，使用错误忽略模式
    content = raw_data.decode('utf-8', errors='ignore')
    print("使用UTF-8编码并忽略错误")

# 处理内容
lines = content.split('\n')
data = []
current_group = {}

for line in lines:
    line = line.strip()
    if line.startswith('-----------------------------------------------'):
        if current_group:
            data.append(current_group)
            current_group = {}
    elif line and ':' in line:
        parts = line.split(':', 1)
        if len(parts) == 2:
            key = parts[0].strip()
            value = parts[1].strip().replace('[[', '').replace(']]', '')
            current_group[key] = value

# 添加最后一组数据
if current_group:
    data.append(current_group)

# 创建DataFrame并保存
df = pd.DataFrame(data)
df.to_excel('output.xlsx', index=False)
print(f"成功保存 {len(data)} 组数据到 output.xlsx")