# Docxplorer v0.1.1
import docx
import os


def ws_decor(ws_func, quantity):
    list_func = [_ for _ in ws_func]
    list_func.append(' ' * (quantity - int(len(list_func))))
    res = ''.join([_ for _ in list_func])
    return res


print('Docxplorer v0.1.1\n')
os_cwd = os.getcwd()
os_dir = os.listdir(os_cwd)
print(os_dir)
# Multi folder search naudojant os_cwd kaip root directory:
for _ in os_dir:
	if '.' not in _:
		print(_)
		print(os.listdir(f'{os_cwd}\\{_}'))
# WIP

while True:
	result_list = []
	result_find = input('Search: ')
	ed_result_find = result_find.lower().encode()	
	for _ in os_dir:
		if _.endswith('.docx'):
			subject_file = docx.Document(f'{os_cwd}\\{_}')
			all_paragraphs = subject_file.paragraphs
			for line in range(len(all_paragraphs)):
				line_text = all_paragraphs[line].text.lower().encode(encoding='utf8')
				if ed_result_find in line_text:
					result_list.append(_)
		else:
			continue

	result_len = len(result_list)
	found = result_len > 0
	
	print(f"{result_len if found else 'No'} match{'' if result_len == 1 else 'es'} found{':' if found else ''}")
	for _ in set(result_list):
		print(f" - {ws_decor(_.split('.')[0], 27)} {result_list.count(_)} time{'s' if result_list.count(_) > 1 else ''}")
	print('')
