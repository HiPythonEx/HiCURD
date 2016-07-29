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
	return codecs.open(path, 'w+', 'utf8')	
	
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
		text += dic2json4control(i)
	if text != "":
		text = "[" + text + "]"
	return text
def dic2json4control(dic):
	text = ""	
	for column in range(len(dictIndex)):
		if text != "":
			text += ","
		text += '"' + dictIndex[column] + '":"' + dic[dictIndex[column]] + '"'
	return "{" + text + "}"
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
		self.DelConfirm = ""
		self.saves = []
		self.search = []
		self.PK = ""
		self.PKs = []
	def get_pks(self):
		if len(self.PKs) > 0:
			return  self.PKs
		arr = self.PK.split(";")
		if len(arr) > 0:
			for i in arr:
				arr2 = i.split("|")
				if len(arr2) > 1:
					self.PKs.append(arr2[1])
		for i in self.PKs:
			print i
		return self.PKs
	def to_string(self):
		return "name:" + self.cname + ";table:" + self.table + ";remark:" + self.remark + "\n\n\n\n" + colls2json(self.saves) + '\n\n\n\n' + colls2json(self.search)
	def to_json(self):
		text = ReadTemplate("json.xml")
		controls_search = self.read_controls_search()
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname).replace("#controls_search#", controls_search).replace("#Tag#", self.remark).replace("#PK#", self.PK).replace("#DelConfirm#", self.DelConfirm)
	def to_sql_xml(self):
		text = ReadTemplate("sql.xml")
		return text.replace("#remark#", self.remark).replace("#cname#", self.cname).replace("#table#", self.table).replace("#Search#", coll2params(self.search)).replace("#Save#", coll2params(self.saves)).replace("#condition#", self.get_condition()).replace("#condition_params#", self.get_condition_params())
	def get_condition(self):
		text = ""
		for i in self.get_pks():
			if text != "":
				text += " and "
			text += i + " = @" + i
		if text != "":
			text = " and " + text
		return text
	def get_condition_params(self):
		text = ""
		for i in self.get_pks():
			tmp = ReadTemplate("param.xml")		
			text += tmp.replace("#name#", i)
		return text
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
	def on_create_case(self):
		text = ReadTemplate("Control_Item.cs")	
		return text.replace("#table#", self.vo_name()).replace("#remark#", self.remark)
	def vo_name(self):
		return self.table.capitalize()


for sheet in book.sheets():
	if sheet.name == "Table":			
		for i in range(2, sheet.nrows):			
			table = Table()
			table.cname = read_val(sheet.row(i)[0].value)
			table.table = read_val(sheet.row(i)[1].value)
			table.remark = read_val(sheet.row(i)[2].value)
			table.PK = read_val(sheet.row(i)[3].value)
			table.DelConfirm = read_val(sheet.row(i)[4].value)
			tables.append(table)
	if sheet.name == "ClolumnInfo":		
		colnames =  sheet.row_values(0) 
		for column in range(len(colnames)):
			dictName[colnames[column]] = column
			dictIndex[column] = colnames[column]
def First_Upper(str):
	if len(str) <= 1:
		return str
	return str[0].upper() + str[1:]
for t in tables:
	for sheet in book.sheets():
		if sheet.name == t.remark + ".Controls":	
			for r in range(2, sheet.nrows):	
				columnInfo = {}
				for column in range(len(dictIndex)):		
					if dictIndex[column] == "Tag":
						columnInfo[dictIndex[column]] = First_Upper(read_val(sheet.row(r)[column].value))
					else:
						columnInfo[dictIndex[column]] = read_val(sheet.row(r)[column].value) 
				t.saves.append(columnInfo)
		if sheet.name == t.remark + ".Search":	
			for r in range(2, sheet.nrows):	
				columnInfo = {}
				for column in range(len(dictIndex)):		
					if dictIndex[column] == "Tag":
						columnInfo[dictIndex[column]] = First_Upper(read_val(sheet.row(r)[column].value))
					else:
						columnInfo[dictIndex[column]] = read_val(sheet.row(r)[column].value) 
				t.search.append(columnInfo)
	
sql = ""
create_case = ""
for i in tables:
	create_case += i.on_create_case()
	export("config/" + i.table + ".xml", i.to_json(), i.table  + " json config success")
	export("sql/" + i.table + ".xml", i.to_sql_xml(), i.table  + " sql success")

text = ReadTemplate("Control.cs")	
cs = text.replace("#create_form#", create_case)

export("ControlsEvent.cs", cs, "controls code success")