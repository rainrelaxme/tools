功能：
1. 从excel收集更新的数据，添加到influxdb
2. redis记录现在最新的一条

过程or注意点：
# 如果文件被占用
	- 重试
	- Csvhelper.fileshare
	- 只读