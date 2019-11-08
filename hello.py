def find_partname(header, header_name):
	for item in header:
		if item.value == header_name:
			return item.column	
	return 0

def load_partnames_to_set(file_xls, page_name, header_name):
	from openpyxl import load_workbook
	import xlsxwriter 
	result = set()
	wb = load_workbook(filename = file_xls)
	ws = wb[page_name]
	row_header = ws[1]
	c = find_partname(row_header, header_name)
	if(c != 0):
		for cell in ws[xlsxwriter.utility.xl_col_to_name(c-1)]:			
			result.add(cell.value)		
	else:
		print("No Partname column")
	return result

def array_to_file(filename, array):
	with open(filename, "w") as file:
		for line in array:
			file.write(line+"\n")


def add_missing_to_array(set1,set2):
	list1 = list()
	for item in set1:
		if not (item in set2):
			list1.append(item)
	list1.sort()
	return list1


described = load_partnames_to_set("bd.xlsx", "Контактирующие", "Partname")
loaded = load_partnames_to_set("sockets.xlsx", "sockets_20190329", "Наименование")
not_online = add_missing_to_array(described, loaded)
array_to_file("not_online.txt", not_online)
array_to_file("wrong_online.txt", add_missing_to_array(loaded, described))

