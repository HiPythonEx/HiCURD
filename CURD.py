#coding=utf-8
import os
import sys
import codecs 
import xlrd

def script_path():
	import inspect, os
	caller_file = inspect.stack()[1][1]         # caller's filename
	path = os.path.abspath(os.path.dirname(caller_file))# path	
	return path
	
def ChangePath(path):
	return path.decode('gbk').encode('gbk')

def read_val(val):
	if isinstance(val, float):
		return str(int(val))
	if isinstance(val, int):
		return str(val)
	return val.rstrip()
	
def ReadTemplate(file):
	path = os.path.join(local.decode('gbk').encode('utf8')+"/tmp/", file)
	f = open(path)
	text = ""
	for i in f:
		text += i
	return text
	
def createAndOpenFile(name):
	path = os.path.join(local.decode('gbk').encode('utf8')+"/src/", name)
	if os.path.exists(path):
		os.remove(path)
	return codecs.open(path, 'w+', 'gbk')	
	
def export(file, text, script):
	csv=createAndOpenFile(file)
	csv.write(text)
	csv.close()
	print script
	
local = script_path()
print "local:" + local
xlsPath = os.path.join(local,"DBConfig.xls")
print "excel:" + xlsPath
xlsPath = ChangePath(xlsPath)
print "excel:" + xlsPath

try:
	book = xlrd.open_workbook(xlsPath)
except Exception , e:
	print "err:",
	print e

print "book:" + str(book)

Rows = []
tables = []
dictName = {}
dictIndex = {}

def read_json(row, dcIndex):
	text = ""	
	for column in range(len(dcIndex)):
		if text != "":
			text += ","
		text += '"' + dcIndex[column] + '":"' + read_val(row[column].value) + '"'
	return "{" + text + "}"

def dic2json(dic):
	text = ""
	for k in dic.keys():
		if text != "":
			text += ","
		text += '"' + k + '":"' + dic[k] + '"'
	if text != "":
		text = "{" + text + "}"
	return text
def colls2json(colls):	
	text = ""
	for i in colls:
		if text != "":
				text += ",\n"
		text += dic2json(i)
	if text != "":
		text = "[" + text + "]"
	return text
def coll2params(colls):
	text = ""
	for k in colls:
		tmp = ReadTemplate("param.xml")		
		text += tmp.replace("#name#", k["Tag"])
	return text
class Table:
	def __init__(self):
		self.cname = ""
		self.table = ""
		self.remark = ""
		self.saves = []
		self.search = []
	def to_string(self):
		return "name:" + self.cname + ";table:" + self.table + ";remark:" + self.remark + "\n\n\n\n" + colls2json(self.saves) + '\n\n\n\n' + colls2json(self.search)
	def to_json(self):
		text = ReadTemplate("json.xml")
		controls_search = self.read_controls_search()
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname).replace("#controls_search#", controls_search).replace("#Tag#", self.remark)
	def to_sql_xml(self):
		text = ReadTemplate("sql.xml")
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname).replace("#table#", self.table).replace("#Search#", coll2params(self.search)).replace("#Save#", coll2params(self.saves))
	def read_sheet(self, name):
		sheet_name = self.remark + "." + name
		for sheet in book.sheets():
			if sheet.name == sheet_name:
				text = ""	
				for i in range(2, sheet.nrows):
					if text != "":
						text += ",\n"
					text += read_json(sheet.row(i), dictIndex)
				return "[" + text + "]"
		return ""
	def read_controls(self):
		return colls2json(self.saves)
	def read_search(self):
		return colls2json(self.search)
	def read_form(self):
		text = ReadTemplate("json.xml")
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname)
	def read_controls_search(self):
		controls = self.read_controls()
		search = self.read_search()
		if controls == "" or search == "":
			return ""
		text = ReadTemplate("controls.xml")		
		return text.replace("#controls#", controls).replace("#search#", search).replace("#remark#", self.remark)
	def on_edit(self):
		text = ReadTemplate("Edit.cs")	
		return text.replace("#table#", self.vo_name())
	def on_edit_case(self):
		text = ReadTemplate("Edit_item.cs")	
		return text.replace("#table#", self.vo_name()).replace("#remark#", self.remark).replace("#operate#", "Edit")
	def on_delete_case(self):
		text = ReadTemplate("Edit_item.cs")	
		return text.replace("#table#", self.vo_name()).replace("#remark#", self.remark).replace("#operate#", "Delete")
	def on_delete(self):
		text = ReadTemplate("Delete.cs").decode('utf8')
		return text.replace("#table#", self.vo_name()).replace("#cname#", self.cname)
	def on_create_case(self):
		text = ReadTemplate("Control_Item.cs")	
		return text.replace("#table#", self.vo_name()).replace("#remark#", self.remark)
	def vo_name(self):
		return self.table.lower().capitalize()


for sheet in book.sheets():
	if sheet.name == "Table":			
		for i in range(2, sheet.nrows):			
			table = Table()
			table.cname = read_val(sheet.row(i)[0].value)
			table.table = read_val(sheet.row(i)[1].value)
			table.remark = read_val(sheet.row(i)[2].value)
			tables.append(table)
	if sheet.name == "ClolumnInfo":		
		colnames =  sheet.row_values(0) 
		for column in range(len(colnames)):
			dictName[colnames[column]] = column
			dictIndex[column] = colnames[column]
		
for t in tables:
	for sheet in book.sheets():
		if sheet.name == t.remark + ".Controls":	
			for r in range(2, sheet.nrows):	
				columnInfo = {}
				for column in range(len(dictIndex)):		
					if dictIndex[column] == "Tag":
						columnInfo[dictIndex[column]] = read_val(sheet.row(r)[column].value).capitalize()
					else:
						columnInfo[dictIndex[column]] = read_val(sheet.row(r)[column].value) 
				t.saves.append(columnInfo)
		if sheet.name == t.remark + ".Search":	
			for r in range(2, sheet.nrows):	
				columnInfo = {}
				for column in range(len(dictIndex)):		
					if dictIndex[column] == "Tag":
						columnInfo[dictIndex[column]] = read_val(sheet.row(r)[column].value).capitalize()
					else:
						columnInfo[dictIndex[column]] = read_val(sheet.row(r)[column].value) 
				t.search.append(columnInfo)
	
sql = ""
json = ""
cs = ""
del_fun = ""
edit_fun = ""
edit_case = ""
del_case = ""
create_case = ""
for i in tables:
	json += i.to_json()
	sql += i.to_sql_xml()
	edit_fun += i.on_edit()
	del_fun += i.on_delete()
	edit_case += i.on_edit_case()
	del_case += i.on_delete_case()
	create_case += i.on_create_case()

text = ReadTemplate("Control.cs")	
cs = text.replace("#create_form#", create_case).replace("#ondelete#", del_case).replace("#onedit#", edit_case).replace("#edit_funs#", edit_fun).replace("#del_funs#", del_fun)

export("json.xml", json, "json success")
export("sql.xml", sql, "sql success")
export("controls.cs", cs, "controls code success")