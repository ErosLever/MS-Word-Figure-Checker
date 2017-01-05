import re, zipfile

def fix_xml_references(xml):
	bookmarks = re.findall('<w:bookmarkStart\W(.*?)<w:bookmarkEnd\W', xml, re.DOTALL)
	img_bookmarks = filter(lambda x:"> SEQ Figur" in x, bookmarks)
	counter = 0
	for bookmark in img_bookmarks:
		counter += 1
		bookmark_id = re.search(r' w:name="(_Ref.*?)"',bookmark).group(1)
		new_value = re.sub(r'<w:t>((?:Figur[ae] )?)\d+</w:t>',r'<w:t>\g<1>%d</w:t>' % counter,bookmark)
		xml = re.sub(re.escape(bookmark),new_value,xml,0,re.DOTALL)
		xml = re.sub(r'> REF %s (.*?)<w:t>((?:Figur[ae] )?)\d+</w:t>(\s*)</w:r>' % bookmark_id, r'> REF %s \1<w:t>\g<2>%d</w:t>\3</w:r>' % (bookmark_id, counter), xml, 0, re.DOTALL)
	return xml

def fix_docx_references(path):
	with zipfile.ZipFile(path) as zin:
		with zipfile.ZipFile(path+".refs.docx","w") as zout:
			for item in zin.infolist():
				content = zin.read(item.filename)
				if item == "word/document.xml":
					content = fix_xml_references(content)
				zout.writestr(item, content)

if __name__ == "__main__":
	import sys
	map(fix_docx_references,sys.argv[1:])