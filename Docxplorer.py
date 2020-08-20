# Docxplorer v0.3.1
import docx
import os
from tkinter import filedialog
import tkinter as tk

os_cwd = os.getcwd()
print_dir = str(os_cwd) + '\\'

screen = tk.Tk()
screen.title('Docxplorer')


def open_something():
	os.startfile(str(os_cwd) + '\\')


def ws_decor(ws_func, quantity):
    list_func = [_ for _ in ws_func]
    list_func.append(' ' * (quantity - int(len(list_func))))
    res = ''.join([_ for _ in list_func])
    return res



menu_bar = tk.Menu(screen)

file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Open", command="hello")
file_menu.add_command(label="Save", command="hello")
file_menu.add_separator()
file_menu.add_command(label="Exit", command=screen.quit)
menu_bar.add_cascade(label="File", menu=file_menu)

edit_menu = tk.Menu(menu_bar, tearoff=0)
edit_menu.add_command(label="Cut", command="hello")
edit_menu.add_command(label="Copy", command="hello")
edit_menu.add_command(label="Paste", command="hello")
menu_bar.add_cascade(label="Edit", menu=edit_menu)

view_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="View", menu=view_menu)
style_menu = tk.Menu(view_menu, tearoff=0)
style_menu.add_command(label="Darkmode", command="hello")
view_menu.add_cascade(label="Style", menu=style_menu)

help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="About", command="hello")
menu_bar.add_cascade(label="Help", menu=help_menu)

screen.config(menu=menu_bar)

dir_entry = tk.Entry(screen, width=45, borderwidth=4)


def search_dir():
    screen.update()
    choose_filename = filedialog.askdirectory()
    if choose_filename:
        print(choose_filename)
        os_cwd = dir_entry.get()
        dir_entry.delete(0, tk.END)
        dir_entry.insert(0, choose_filename)


dir_entry.insert(0, str(os_cwd))
dir_button = tk.Button(screen, text=' Browse ', width=9,  padx=3, command=search_dir)

search_entry = tk.Entry(screen, width=45, borderwidth=4)


def finder(f_cwd=dir_entry.get()):
	f_result_list = []
	f_result_find = search_entry.get().lower().encode()
	if not f_result_find:
		return ['Output screen']
	f_dir = os.listdir(f_cwd)
	for _ in f_dir:
		if _.endswith('.docx'):
			file_path = f'{f_cwd}\\{_}'
			subject_file = docx.Document(file_path)
			all_paragraphs = subject_file.paragraphs
			for line in range(len(all_paragraphs)):
				line_text = all_paragraphs[line].text.lower().encode(encoding='utf8')
				if f_result_find in line_text:
					for single_result in range(line_text.count(f_result_find)):
						f_result_list.append(file_path)
		if '.' not in _:
			try:
				f_result_list += finder(f'{f_cwd}\\{_}')
			except (NotADirectoryError, TypeError) as e:
				continue
	return f_result_list




output_text = tk.Text(screen, width=63, height=9, borderwidth=4)
output_text.insert(1.0, 'Output screen')
output_text.configure(state='disabled')


def output_find():
	result_list = []
	result_list += finder()
	result_len = len(result_list)
	found = result_len > 0

	output_text.configure(state='normal')
	output_text.delete(1.0, tk.END)
	if result_list != ['Output screen']:
		output_text.insert(1.0, f"{result_len if found else 'No'} match{'' if result_len == 1 else 'es'} found{':' if found else ''}\n")
		for _ in set(result_list):
			output_text.insert(tk.END, f" - {ws_decor(_.split('.')[0].replace(print_dir, ''), 36)} {result_list.count(_)} time{'s' if result_list.count(_) > 1 else ''}\n")
	else:
		output_text.insert(tk.END, 'Output screen')
	output_text.configure(state='disabled')


search_button = tk.Button(screen, text=' Search ', width=9, padx=3, command=output_find)
output_text_scrollbar = tk.Scrollbar(screen, command=output_text)

for _ in range(0, 11):
	tk.Label(screen, text='', width=6).grid(row=0, column=_)

empty_label1 = tk.Label(screen, text='').grid(row=1, column=0)

dir_entry.grid(row=2, column=2, columnspan=5, padx=3)
dir_button.grid(row=2, column=1)

empty_label2 = tk.Label(screen, text='').grid(row=3, column=0)

search_entry.grid(row=4, column=2, columnspan=5, padx=3)
search_button.grid(row=4, column=1)

for _ in range(5, 9):
	tk.Label(screen, text='', width=6).grid(row=_, column=0)

output_text.grid(row=7, column=1, columnspan=9)

screen.mainloop()
