1.通过串口的数据显示曲线,在关闭窗口后会把数据保存在excel中

2.串口要求:115200,格式:current_speed:2108,current_duty:2590,target_speed:2100

3.数据分列是根据冒号和逗号来分割的

4.安装库: pip install pyserial matplotlib openpyxl