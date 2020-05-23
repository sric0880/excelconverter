# encoding: utf-8
import re
from openpyxl import load_workbook

def __parseDict(d):
	d = { k: v for k, v in d.items() if k is not None}
	ret = {}
	for k,v in d.items():
		i = k.find('<')
		if i != -1:
			k = k[0:i]
		found = k.find(':')
		if found != -1:
			subKey = k[0:found]
			secondaryKey = k[found+1:]
			subValue = ret.setdefault(subKey, {})
			subValue[secondaryKey] = '' if v is None else v
		else:
			ret[k] = '' if v == None else v

	for k,v in ret.items():
		if type(v) is dict:
			ret[k] = __parseDict(v)

	pattern = re.compile(r'([\w_]+)\[(\d+)\]')
	for k,v in ret.items():
		match = pattern.match(k)
		if match:
			subKey = match.group(1)
			index = int(match.group(2))
			ret.setdefault(subKey, []).insert(index, v)
			del ret[k]

	return ret

def parseObject(xlsxfile):
	wb = load_workbook(filename = xlsxfile)
	ws = wb.active
	rowsGenerator = ws.rows
	headerRow = next(rowsGenerator)
	header = [ cell.value for cell in headerRow ]
	if ws.title == "ObjectList":
		o = []
		for row in rowsGenerator:
			item = dict(zip(header, [ cell.value for cell in row ]))
			o.append(__parseDict(item))
	elif ws.title == "PrimitiveList":
		o = []
		for row in rowsGenerator:
			o.append(row[0].value)
	elif ws.title == "SingleObject":
		o = __parseDict(dict(zip(header, [cell.value for cell in next(rowsGenerator)])))
	else:
		raise Exception('Excel type not supported')
	return o
