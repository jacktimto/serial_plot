import serial
import tkinter as tk
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.animation import FuncAnimation
import matplotlib.pyplot as plt
from openpyxl import Workbook
from datetime import datetime

# 配置串口
SERIAL_PORT = 'COM4'  # 设置为 COM4
BAUD_RATE = 115200  # 设置为 115200

try:
    ser = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=1)
    ser.flushInput()
except serial.SerialException as e:
    print(f"无法打开串口 {SERIAL_PORT}: {e}")
    exit()

# 初始化数据字典
data_dict = {
    'current_speed': [],
    'current_duty': [],
    'target_speed': []
}

# 设置数据点的最大数量
MAX_POINTS = 100
paused = False  # 用于控制暂停和继续

# 创建 Excel 工作簿和表格
wb = Workbook()
ws = wb.active
ws.title = "Serial Data"
ws.append(list(data_dict.keys()))  # 写入标题行

# 创建Tkinter窗口
root = tk.Tk()
root.title("实时串口数据曲线")

# 创建Figure对象，用于在Tkinter窗口中嵌入Matplotlib图
fig = Figure(figsize=(8, 6), dpi=100)
ax = fig.add_subplot(1, 1, 1)

# 设置支持中文的字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体字体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 设置图表标签和标题
ax.set_xlabel("数据点")
ax.set_ylabel("值")
ax.set_title("实时串口数据")
ax.grid(True)

# 初始化线条对象和复选框状态
lines = {}
line_visible = {}  # 保存各曲线的显示状态

# 为每条数据创建曲线和复选框
for key in data_dict.keys():
    (line,) = ax.plot([], [], label=key)
    lines[key] = line
    line_visible[key] = tk.BooleanVar(value=True)  # 初始化为显示状态

    # 创建复选框
    chk = tk.Checkbutton(root, text=key, variable=line_visible[key], onvalue=True, offvalue=False,
                         command=lambda: update_visibility())
    chk.pack(anchor='w')  # 左对齐放置

ax.legend()

canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack()

# 创建一个注释对象，用于显示数据点的值
annot = ax.annotate("", xy=(0, 0), xytext=(20, 20),
                    textcoords="offset points",
                    bbox=dict(boxstyle="round", fc="w"),
                    arrowprops=dict(arrowstyle="->"))
annot.set_visible(False)


# 更新可见性函数
def update_visibility():
    for key, line in lines.items():
        line.set_visible(line_visible[key].get())


# 更新注释位置和文本
def update_annot(ind, key):
    x, y = lines[key].get_data()
    annot.xy = (x[ind["ind"][0]], y[ind["ind"][0]])
    text = f"{key}: ({x[ind['ind'][0]]}, {y[ind['ind'][0]]})"
    annot.set_text(text)
    annot.get_bbox_patch().set_alpha(0.8)


# 鼠标移动事件处理
def on_move(event):
    vis = annot.get_visible()
    if event.inaxes == ax:
        for key, line in lines.items():
            if line_visible[key].get():  # 仅在曲线可见时显示注释
                cont, ind = line.contains(event)
                if cont:
                    update_annot(ind, key)
                    annot.set_visible(True)
                    canvas.draw()
                    return
    if vis:
        annot.set_visible(False)
        canvas.draw()


canvas.mpl_connect("motion_notify_event", on_move)


# 更新绘图函数
def update_plot(frame):
    if paused:
        return  # 如果处于暂停状态，不更新绘图

    while ser.in_waiting:
        try:
            line = ser.readline().decode('utf-8').strip()
            if not line:
                continue
            # 解析数据
            entries = line.split(',')
            temp_data = {}
            for entry in entries:
                if ':' in entry:
                    key, value = entry.split(':', 1)
                    key = key.strip()
                    value = float(value.strip())
                    if key in data_dict:
                        temp_data[key] = value
            # 更新数据字典
            row_data = []
            for key, value in temp_data.items():
                data_dict[key].append(value)
                row_data.append(value)
                if len(data_dict[key]) > MAX_POINTS:
                    data_dict[key].pop(0)
            # 将数据写入 Excel
            ws.append(row_data)
        except ValueError as ve:
            print(f"数据解析错误: {ve}")
        except Exception as e:
            print(f"其他错误: {e}")

    # 更新每条线的数据
    for key, line in lines.items():
        y_data = data_dict[key]
        x_data = list(range(len(y_data)))
        line.set_data(x_data, y_data)
        if y_data:
            ax.set_xlim(0, MAX_POINTS)
            y_min = min(min(data_dict[key] for key in data_dict), default=0) - 10
            y_max = max(max(data_dict[key] for key in data_dict), default=0) + 10
            ax.set_ylim(y_min, y_max)

    # 更新每条线的可见性
    update_visibility()
    canvas.draw()


# 放大、缩小、暂停、继续功能
def pause():
    global paused
    paused = True


def resume():
    global paused
    paused = False


# 启动动画
ani = FuncAnimation(fig, update_plot, interval=100,cache_frame_data=False)

# 在窗口中添加控制按钮
button_frame = tk.Frame(root)
button_frame.pack()

pause_btn = tk.Button(button_frame, text="暂停", command=pause)
pause_btn.pack(side="left")

resume_btn = tk.Button(button_frame, text="继续", command=resume)
resume_btn.pack(side="left")

# 运行Tkinter主循环
try:
    root.mainloop()
except KeyboardInterrupt:
    pass
finally:
    ser.close()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"串口数据记录_{timestamp}.xlsx"
    wb.save(filename)  # 保存Excel文件
