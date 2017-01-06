import re, zipfile
from collections import OrderedDict

def xml2txt(xml):
	return "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>',xml))

sequences = {}

class Bookmark(object):
	def __init__(self,id,name,xml):
		self.id = int(id)
		self.name = name
		self.xml = xml
		self.nested = bool(re.search(r'<w:bookmarkStart',xml))
		self.text = xml2txt(xml)
		
		field_xml = re.search(r'<w:fldChar w:fldCharType="begin"/>.*?<w:fldChar w:fldCharType="end"/>',xml,re.DOTALL)
		if field_xml:
			instrText = re.search(r'<w:instrText[^>]*>(.*?)</w:instrText>',field_xml.group(0),re.DOTALL)
			if instrText:
				seq_name = re.search(r'SEQ (\w+)',instrText.group(1),re.DOTALL)
				if seq_name:
					self.seq_name = seq_name.group(1)
					sequence = sequences.get(self.seq_name,[])
					sequence.append(self)
					self.seq_num = len(sequence)
					sequences[self.seq_name] = sequence
					self.correct = self.text == "%s %d" % (self.seq_name,self.seq_num)
		if not hasattr(self,"seq_name"):
			self.seq_name = None
		if not hasattr(self,"seq_num"):
			self.seq_num = 0
		if not hasattr(self,"correct"):
			self.correct = True

	def __repr__(self):
		return "<bookmark id=%d name=%s nested=%s seq_name=%s seq_num=%s correct=%s>%s</bookmark>" % (self.id,self.name,self.nested,self.seq_name,self.seq_num,self.correct,self.text)

def find_all_bookmarks(xml):
	bookmark_ids = re.findall(r'<w:bookmarkStart w:id="(\d+)"', xml, re.DOTALL)
	bookmarks = []
	for bookmark_id in bookmark_ids:
		regex = r'<w:bookmarkStart w:id="(%s)" w:name="([^"]+)"[^>]*>(.*?)<w:bookmarkEnd w:id="%s"/>' % (bookmark_id,bookmark_id)
		bookmark_xml = re.search(regex, xml, re.DOTALL)
		bookmark = Bookmark(*bookmark_xml.groups())
		bookmarks.append(bookmark)
	return bookmarks

def find_all_sequence_bookmarks(xml):
	return filter(lambda x:x.seq_name,find_all_bookmarks(xml))

def find_references(xml,)

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


'''
      <w:bookmarkStart w:id="141" w:name="_Ref470181140"/>
      <w:proofErr w:type="spellStart"/>
      <w:r w:rsidRPr="007D4707">
        <w:t>Figura</w:t>
      </w:r>
      <w:proofErr w:type="spellEnd"/>
      <w:r w:rsidRPr="007D4707">
        <w:t xml:space="preserve"> </w:t>
      </w:r>
      <w:r w:rsidRPr="007D4707">
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r w:rsidRPr="007D4707">
        <w:instrText xml:space="preserve"> SEQ Figura \* ARABIC </w:instrText>
      </w:r>
      <w:r w:rsidRPr="007D4707">
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r w:rsidR="000760DB">
        <w:rPr>
          <w:noProof/>
        </w:rPr>
        <w:t>1</w:t>
      </w:r>
      <w:r w:rsidRPr="007D4707">
        <w:fldChar w:fldCharType="end"/>
      </w:r>
      <w:bookmarkEnd w:id="141"/>
'''