import tkinter as tk
from tkinter import font as tkFont
from tkinter import simpledialog, filedialog, messagebox
from tkcalendar import Calendar
from datetime import datetime, timedelta
from PIL import Image, ImageTk
import win32ui, win32print
import json, os, shutil, re

# 设置图片存储目录
# Set up the image storage directory
img_directory = 'img'
if not os.path.exists(img_directory):
    os.makedirs(img_directory)

# 定义打印内容的函数
# Define the function to print content
def print_content(content):
    # 使用正则表达式匹配数字和单位（天或个月）
    # Use regex to match numbers and units (days or months)
    match = re.match(r'(\d+)(天|个月)', content)
    if not match:
        print("打印内容必须包含时间单位，例如 '7天' 或 '3个月'")
        return

    # 提取数字和单位
    # Extract the number and unit
    number = int(match.group(1))
    unit = match.group(2)

    # 使用应用程序中存储的日期作为起始时间
    # Use the date stored in the application as the start date
    start_date_str = app.selected_date_str
    start_date = datetime.strptime(start_date_str, '%Y年%m月%d日')

    # 根据单位计算到期时间
    # Calculate the expiration date based on the unit
    if unit == '天':
        # 开始日期的当天算作第一天，所以到期日期是开始日期后的 number - 1 天
        # The day of the start date is counted as the first day, so the expiration date is the start date plus number - 1 days
        end_date = start_date + timedelta(days=number - 1)
    elif unit == '个月':
        # 开始日期的当天算作第一天，所以到期日期是开始日期后的 number 个月再减去一天
        # The day of the start date is counted as the first day, so the expiration date is the start date plus number months, minus one day
        month = start_date.month - 1 + number
        year = start_date.year + month // 12
        month = month % 12 + 1
        day = start_date.day
        end_date = datetime(year, month, day) - timedelta(days=1)

    # 格式化打印内容
    # Format the print content
    print_str = "开启时间：          {}{}\n{}\n到期时间：\n{}".format(
        number, unit, start_date.strftime("%Y年%m月%d日"), end_date.strftime("%Y年%m月%d日")
    )

    # 打印机配置
    # Printer configuration
    printer_name = win32print.GetDefaultPrinter()
    pdc = win32ui.CreateDC()
    pdc.CreatePrinterDC(printer_name)
    
    # 开始打印文档
    # Start printing the document
    pdc.StartDoc("打印内容")
    pdc.StartPage()

    # 设置打印起始位置
    # Set the start position for printing
    start_x = 100
    start_y = 100
    line_height = 100  # 设置行高 Set line height

    # 按行分割文本并逐行打印
    # Split text by line and print line by line
    lines = print_str.split('\n')
    for i, line in enumerate(lines):
        pdc.TextOut(start_x, start_y + i * line_height, line)

    # 结束打印
    # End printing
    pdc.EndPage()
    pdc.EndDoc()
    pdc.DeleteDC()

# 图片打印区域类定义
# Definition of the ImagePrintArea class
class ImagePrintArea(tk.Frame):
    # 构造函数
    # Constructor
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.config(borderwidth=2, relief="groove")
        self.images = []  # 存储图片路径、描述和打印内容 Store image paths, descriptions, and print content
        self.load_images_data()  # 加载现有图片数据 Load existing image data
        self.create_widgets()
    
    # 创建图片打印区域内的各种组件
    # Create various components in the image print area
    def create_widgets(self):
        # 清除所有子组件
        # Clear all child components
        for widget in self.winfo_children():
            widget.destroy()

        # 创建右键菜单
        # Create a context menu
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="修改", command=self.edit_image)
        self.context_menu.add_command(label="删除", command=self.remove_image)

        # 根据图片信息创建界面
        # Create the interface based on image information
        for i, (image_path, description, content_to_print) in enumerate(self.images):
            img = Image.open(image_path)
            img = img.resize((180, 150), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            label = tk.Label(self, image=photo, text=description, compound="top")
            label.image = photo  # 保持对图片的引用 Keep a reference to the photo
            label._image_index = i  # 保存图片索引 Save the image index
            label.grid(row=i // 6 * 3, column=i % 6 * 2, padx=5, pady=5, columnspan=2)
            label.bind('<Button-1>', lambda event, content=content_to_print: print_content(content))
            label.bind('<Button-3>', self.popup_context_menu)

    # 弹出右键菜单
    # Popup the context menu
    def popup_context_menu(self, event):
        try:
            self.context_menu._image_index = event.widget._image_index
            self.context_menu.post(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    # 编辑图片信息
    # Edit image information
    def edit_image(self):
        index = self.context_menu._image_index
        if 0 <= index < len(self.images):
            image_path, _, _ = self.images[index]
            description = simpledialog.askstring("修改图片说明", "请输入新的图片说明:", initialvalue=self.images[index][1])
            content_to_print = simpledialog.askstring("修改打印内容", "请输入新的打印内容:", initialvalue=self.images[index][2])
            if description and content_to_print:
                self.images[index] = (image_path, description, content_to_print)
                self.save_images_data()
                self.create_widgets()
    
    # 删除图片信息
    # Delete image information
    def remove_image(self):
        index = self.context_menu._image_index
        if 0 <= index < len(self.images):
            del self.images[index]
            self.save_images_data()
            self.create_widgets()

    # 添加新图片
    # Add a new image
    def add_image(self):
        image_path = filedialog.askopenfilename()
        if image_path:
            description = simpledialog.askstring("图片说明", "请输入图片说明:")
            content_to_print = simpledialog.askstring("打印内容", "请输入打印内容:")
            if description and content_to_print:
                # 将图片复制到 'img' 目录并更新图片路径
                # Copy the image to the 'img' directory and update the image path
                img_filename = os.path.basename(image_path)
                target_path = os.path.join(img_directory, img_filename)
                shutil.copy(image_path, target_path)
                self.images.append((target_path, description, content_to_print))
                self.save_images_data()  # 保存更新的数据 Save updated data
                self.create_widgets()

    # 保存图片信息到 JSON 文件
    # Save image information to a JSON file
    def save_images_data(self):
        images_info_path = os.path.join(img_directory, 'images_data.json')
        with open(images_info_path, 'w') as f:
            json.dump(self.images, f)

    # 从 JSON 文件加载图片信息
    # Load image information from a JSON file
    def load_images_data(self):
        images_info_path = os.path.join(img_directory, 'images_data.json')
        if os.path.exists(images_info_path):
            with open(images_info_path, 'r') as f:
                self.images = json.load(f)
        self.create_widgets()

# 快速打印区域类定义
# Definition of the QuickPrintArea class
class QuickPrintArea(tk.Frame):
    # 构造函数
    # Constructor
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.config(borderwidth=2, relief="groove")
        self.buttons = []  # 存储按钮文本和打印内容 Store button text and print content
        entryFont = tkFont.Font(family="宋体", size=14)
        self.entry = tk.Entry(self, font=entryFont, width=20)
        self.entry.grid(row=0, column=0, padx=5)
        self.entry.insert(0, "在此输入天数,如10天")  # 设置默认文字 Set default text
        self.entry.bind("<FocusIn>", self.clear_default_text)  # 绑定事件处理函数 Bind event handler

        self.print_button = tk.Button(self, text="打印", command=self.print_custom_content, width=10, height=2)
        self.print_button.grid(row=0, column=1, padx=5)
        self.load_buttons_config()  # 加载现有按钮配置 Load existing button configurations
        self.create_buttons()

    # 当用户点击 Entry 时，如果文本是默认值，则清除它
    # When the user clicks the Entry, if the text is the default value, clear it
    def clear_default_text(self, event):
        if self.entry.get() == "在此输入天数,如10天":
            self.entry.delete(0, tk.END)

    # 创建快速打印按钮
    # Create quick print buttons
    def create_buttons(self):
        # 移除除了 entry 和 print_button 的所有组件
        # Remove all components except entry and print_button
        for widget in self.winfo_children():
            if widget not in (self.entry, self.print_button):
                widget.destroy()

        # 创建右键菜单
        # Create a context menu
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="修改", command=self.edit_button)
        self.context_menu.add_command(label="删除", command=self.remove_button)

        customFont = tkFont.Font(family="宋体", size=14, weight="bold")
        # 根据配置创建按钮
        # Create buttons based on the configuration
        for i, (text, content) in enumerate(self.buttons):
            btn = tk.Button(self, text=text, command=lambda c=content: print_content(c), font=customFont, height=3, width=12)
            btn._button_index = i  # 保存按钮索引 Save the button index
            btn.grid(row=1 + i // 8, column=i % 8, padx=5, pady=5)
            btn.bind('<Button-1>', lambda event, c=content: print_content(c))
            btn.bind('<Button-3>', self.popup_context_menu)

    # 弹出右键菜单
    # Popup the context menu
    def popup_context_menu(self, event):
        try:
            self.context_menu._button_index = event.widget._button_index
            self.context_menu.post(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    # 编辑按钮
    # Edit button
    def edit_button(self):
        index = self.context_menu._button_index
        if 0 <= index < len(self.buttons):
            text, _ = self.buttons[index]
            new_text = simpledialog.askstring("修改按钮文本", "请输入新的按钮文本:", initialvalue=text)
            new_content = simpledialog.askstring("修改打印内容", "请输入新的打印内容:", initialvalue=self.buttons[index][1])
            if new_text and new_content:
                self.buttons[index] = (new_text, new_content)
                self.create_buttons()
                self.save_buttons_config()

    # 删除按钮
    # Delete button
    def remove_button(self):
        index = self.context_menu._button_index
        if 0 <= index < len(self.buttons):
            del self.buttons[index]
            self.create_buttons()
            self.save_buttons_config()

    # 打印用户自定义内容
    # Print user-defined content
    def print_custom_content(self):
        content = self.entry.get()
        if content and content != "在此输入天数,如10天":
            print_content(content)
        else:
            messagebox.showerror("错误", "请输入有效的打印内容!")

    # 添加新按钮
    # Add a new button
    def add_button(self):
        text = simpledialog.askstring("按钮文本", "请输入按钮文本:")
        content = simpledialog.askstring("打印内容", "请输入打印内容:")
        if text and content:
            self.buttons.append((text, content))
            self.create_buttons()
            self.save_buttons_config()  # 保存更新后的配置 Save the updated configurations

    # 保存按钮配置到 JSON 文件
    # Save button configurations to a JSON file
    def save_buttons_config(self):
        # Save the button configurations to a JSON file
        with open('buttons_config.json', 'w') as f:
            json.dump(self.buttons, f)

    # 从 JSON 文件加载按钮配置
    # Load button configurations from a JSON file
    def load_buttons_config(self):
        # Load the button configurations from a JSON file if it exists
        try:
            with open('buttons_config.json', 'r') as f:
                self.buttons = json.load(f)
            self.create_buttons()
        except FileNotFoundError:
            pass  # 没有发现配置文件 No configuration file found

# 主应用程序类定义
# Definition of the main Application class
class Application(tk.Tk):
    # 构造函数
    # Constructor
    def __init__(self):
        super().__init__()
        self.title("标签打印程序")
        self.state('zoomed')  # Windows系统专用 Specific to Windows system
        self.image_print_area = ImagePrintArea(self)
        self.image_print_area.pack(side="top", fill="x", pady=10)
        self.quick_print_area = QuickPrintArea(self)
        self.quick_print_area.pack(side="bottom", fill="x", pady=10)

        # 创建菜单栏
        # Create menu bar
        menu_font = tkFont.Font(family="宋体", size=20, weight="bold")
        self.menu = tk.Menu(self, font=menu_font)
        self.config(menu=self.menu)
        edit_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="编辑", menu=edit_menu)
        edit_menu.add_command(label="添加图片", font=menu_font, command=self.image_print_area.add_image)
        edit_menu.add_command(label="添加快捷按钮", font=menu_font, command=self.quick_print_area.add_button)

        # 添加"设定日期"菜单项
        # Add "Set Date" menu item
        date_menu = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="设定日期", menu=date_menu)
        date_menu.add_command(label="选择日期", font=menu_font, command=self.show_date_picker)

        # 初始化用于显示日期的菜单项
        # Initialize the menu item for displaying the date
        self.selected_date_str = datetime.now().strftime("%Y年%m月%d日")
        self.menu.add_command(label="当前日期:"+self.selected_date_str, state='disabled')

    # 显示日期选择器
    # Show the date picker
    def show_date_picker(self):
        # 创建一个新的顶级窗口
        # Create a new top-level window
        top = tk.Toplevel(self)
        top.title("选择日期")

        # 在顶级窗口中添加一个 Calendar 控件
        # Add a Calendar widget to the top-level window
        cal = Calendar(top, selectmode='day', year=datetime.now().year, month=datetime.now().month, day=datetime.now().day, date_pattern='y年m月d日')
        cal.pack(pady=20)

        # 添加一个按钮设置日期
        # Add a button to set the date
        set_date_button = tk.Button(top, text="确定", command=lambda: self.set_date(cal.get_date(), top))
        set_date_button.pack()

    # 设置日期
    # Set the date
    def set_date(self, date, top_level_window):
        # 解析日期字符串为 datetime 对象并重新格式化
        # Parse the date string into a datetime object and reformat
        self.selected_date_str = datetime.strptime(date, '%Y年%m月%d日').strftime('%Y年%m月%d日')
        # 根据选定的日期更新菜单项
        # Update the menu item with the selected date
        self.menu.entryconfig(3, label="当前日期:"+self.selected_date_str)
        # 关闭日期选择器窗口
        # Close the date picker window
        top_level_window.destroy()

# 创建并运行应用程序实例
# Create and run the application instance
app = Application()
app.mainloop()