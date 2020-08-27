# Docxplorer v0.3.2
# This code is used to find text match amount in docx format documents which are in code's directory
import docx
import os
from tkinter import filedialog
import tkinter as tk

os_cwd = os.getcwd()
print_dir = str(os_cwd) + '\\'

screen = tk.Tk()
screen.title('Docxplorer')


def ws_decor(ws_func, quantity):
	# this function improves readability of search result
    list_func = [_ for _ in ws_func]
    list_func.append(' ' * (quantity - int(len(list_func))))
    res = ''.join([_ for _ in list_func])
    return res


search_entry = tk.Entry(screen, width=45, borderwidth=4)


def finder():
	# main search function which also encodes search text and document text for better search results
	f_result_list = []
	f_result_find = search_entry.get().lower().encode()
	if not f_result_find:
		return ['Output screen']
	f_dir = os.listdir(os_cwd)
	for _ in f_dir:
		if _.endswith('.docx'):
			file_path = f'{os_cwd}\\{_}'
			subject_file = docx.Document(file_path)
			all_paragraphs = subject_file.paragraphs
			for line in range(len(all_paragraphs)):
				line_text = all_paragraphs[line].text.lower().encode(encoding='utf8')
				if f_result_find in line_text:
					for single_result in range(line_text.count(f_result_find)):
						f_result_list.append(file_path)
	return f_result_list




output_text = tk.Text(screen, width=63, height=9, borderwidth=4)
output_text.insert(1.0, 'Output screen')
output_text.configure(state='disabled')


def output_find():
	# output function for result display at the bottom of GUI
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

for _ in range(0, 11):
	tk.Label(screen, text='', width=6).grid(row=0, column=_)

empty_label1 = tk.Label(screen, text='').grid(row=3, column=0)

search_entry.grid(row=4, column=2, columnspan=5, padx=3)
search_button.grid(row=4, column=1)

for _ in range(5, 9):
	tk.Label(screen, text='', width=6).grid(row=_, column=0)

output_text.grid(row=7, column=1, columnspan=9)

screen.mainloop()
