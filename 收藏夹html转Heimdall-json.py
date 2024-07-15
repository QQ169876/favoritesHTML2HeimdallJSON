import json
import random
import subprocess
import sys
import os
import webbrowser

# 自动安装所需库
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# 安装依赖库
try:
    from bs4 import BeautifulSoup
except ImportError:
    install('beautifulsoup4')
    from bs4 import BeautifulSoup

try:
    from tkinter import Tk, filedialog, messagebox, Listbox, MULTIPLE, Button, END, Checkbutton, IntVar, Label, Frame, LEFT, BOTH
except ImportError:
    install('tk')
    from tkinter import Tk, filedialog, messagebox, Listbox, MULTIPLE, Button, END, Checkbutton, IntVar, Label, Frame, LEFT, BOTH

def generate_random_color():
    """生成随机颜色"""
    return "#{:06x}".format(random.randint(0, 0xFFFFFF))

def is_edge_bookmark_file(soup):
    """判断是否为Edge浏览器导出的收藏夹文件"""
    return soup.find('h3') is not None

def convert_html_to_json(html_file_path, selected_folders, selected_titles, use_folder_color):
    """将HTML文件转换为JSON格式"""
    with open(html_file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

    bookmarks = []

    def extract_bookmarks(folder, folder_name, folder_color=None):
        for item in folder.find_all(['a', 'h3']):
            if item.name == 'h3':
                subfolder = item.find_next_sibling('dl')
                subfolder_color = generate_random_color() if use_folder_color else None
                if subfolder:
                    extract_bookmarks(subfolder, folder_name + '/' + item.get_text(), subfolder_color)
            elif item.name == 'a':
                bookmark_color = folder_color if use_folder_color else generate_random_color()
                bookmark = {
                    "title": item.get_text(),
                    "colour": bookmark_color,
                    "url": item.get('href'),
                    "description": None,
                    "appid": "null",
                    "appdescription": item.get_text()
                }
                if folder_name in selected_folders and (not selected_titles or bookmark["title"] in selected_titles):
                    bookmarks.append(bookmark)

    def extract_links():
        for item in soup.find_all('a'):
            bookmark = {
                "title": item.get_text(),
                "colour": generate_random_color(),
                "url": item.get('href'),
                "description": None,
                "appid": "null",
                "appdescription": item.get_text()
            }
            bookmarks.append(bookmark)

    if is_edge_bookmark_file(soup):
        def process_selected_folders():
            for folder_name in selected_folders:
                folder = soup.find('h3', text=folder_name)
                if folder:
                    extract_bookmarks(folder.find_next_sibling('dl'), folder_name, generate_random_color() if use_folder_color else None)
                else:
                    messagebox.showerror("错误", f"未找到目录 '{folder_name}'")

        process_selected_folders()
    else:
        extract_links()

    json_file_path = html_file_path.replace('.html', '.json')

    if os.path.exists(json_file_path):
        json_file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=(("JSON文件", "*.json"), ("所有文件", "*.*")),
            initialfile=os.path.basename(json_file_path)
        )
        if not json_file_path:
            return

    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(bookmarks, json_file, ensure_ascii=False, indent=2)

    print(f"HTML书签已转换为JSON格式，并保存为 {json_file_path}")
    
    # 打开JSON文件所在文件夹
    json_dir = os.path.dirname(os.path.abspath(json_file_path))
    webbrowser.open(f'file://{json_dir}')

def select_html_file():
    """选择HTML文件的对话框"""
    root = Tk()
    root.withdraw()  # 隐藏Tk窗口
    file_path = filedialog.askopenfilename(
        title="选择HTML文件",
        filetypes=(("HTML文件", "*.html"), ("所有文件", "*.*"))
    )
    root.destroy()  # 销毁隐藏的Tk窗口
    if file_path:
        display_folder_selection(file_path)

def display_folder_selection(file_path):
    """显示目录选择对话框"""
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

    folder_names = ["一般Html"] if not is_edge_bookmark_file(soup) else [h3.get_text() for h3 in soup.find_all('h3')]

    def on_folder_select(event):
        selected_indices = listbox_folders.curselection()
        if selected_indices:
            selected_folder_name = folder_names[selected_indices[0]]
            display_titles(selected_folder_name)

    def on_ok():
        selected_folder_indices = listbox_folders.curselection()
        selected_folders = [folder_names[i] for i in selected_folder_indices]
        selected_title_indices = listbox_titles.curselection()
        selected_titles = [listbox_titles.get(i) for i in selected_title_indices]
        use_folder_color = folder_color_var.get() == 1
        convert_html_to_json(file_path, selected_folders, selected_titles, use_folder_color)

    def select_all():
        listbox_titles.select_set(0, END)

    def deselect_all():
        listbox_titles.select_clear(0, END)

    def invert_selection():
        selected_indices = listbox_titles.curselection()
        for i in range(listbox_titles.size()):
            if i in selected_indices:
                listbox_titles.select_clear(i)
            else:
                listbox_titles.select_set(i)

    def display_titles(folder_name):
        listbox_titles.delete(0, END)
        if folder_name == "一般Html":
            for url in soup.find_all('a'):
                listbox_titles.insert(END, url.get_text())
        else:
            folder = soup.find('h3', text=folder_name)
            if folder:
                urls = folder.find_next_sibling('dl').find_all('a')
                for url in urls:
                    listbox_titles.insert(END, url.get_text())

    def browse_file():
        select_html_file()

    folder_selection = Tk()
    folder_selection.title("选择要转换的目录和网址")

    screen_width = folder_selection.winfo_screenwidth()
    screen_height = folder_selection.winfo_screenheight()
    window_width = 600
    window_height = 700
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    folder_selection.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    frame = Frame(folder_selection)
    frame.pack(pady=10, expand=True, fill=BOTH)

    listbox_folders = Listbox(frame, selectmode=MULTIPLE)
    listbox_folders.pack(pady=5, expand=True, fill=BOTH)
    listbox_folders.bind('<<ListboxSelect>>', on_folder_select)

    for name in folder_names:
        listbox_folders.insert(END, name)

    listbox_titles = Listbox(frame, selectmode=MULTIPLE)
    listbox_titles.pack(pady=5, expand=True, fill=BOTH)

    folder_color_var = IntVar(value=1)  # 默认选中
    folder_color_checkbutton = Checkbutton(frame, text="按目录随机颜色", variable=folder_color_var)
    folder_color_checkbutton.pack(pady=5)

    button_frame = Frame(frame)
    button_frame.pack(pady=10)

    browse_button = Button(button_frame, text="浏览", command=browse_file)
    browse_button.pack(side=LEFT, padx=5)

    ok_button = Button(button_frame, text="开始转换", command=on_ok)
    ok_button.pack(side=LEFT, padx=5)

    select_all_button = Button(button_frame, text="全选", command=select_all)
    select_all_button.pack(side=LEFT, padx=5)

    deselect_all_button = Button(button_frame, text="取消选择", command=deselect_all)
    deselect_all_button.pack(side=LEFT, padx=5)

    invert_selection_button = Button(button_frame, text="反选", command=invert_selection)
    invert_selection_button.pack(side=LEFT, padx=5)

    folder_selection.mainloop()

if __name__ == "__main__":
    select_html_file()