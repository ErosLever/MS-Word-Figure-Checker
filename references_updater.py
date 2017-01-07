import re, zipfile
from collections import OrderedDict
from functools import reduce

def xml2txt(xml,br=False):
	if br:
		xml = re.sub(r'<w:br/>','\n',xml)
	return "".join(re.findall(r'<w:t[^\w>]*>(.*?)</w:t>',xml))

class EqualityChecker(object):
	def __eq__(self,other):
		if isinstance(other, self.__class__):
			return self.__dict__ == other.__dict__
		return NotImplemented

	def __ne__(self, other):
		if isinstance(other, self.__class__):
			return not self.__eq__(other)
		return NotImplemented

	def __hash__(self):
		return hash(tuple(sorted(self.__dict__.items())))

class Document(EqualityChecker):
	def __init__(self,path):
		with zipfile.ZipFile(path) as zin:
			for item in zin.infolist():
				if item.filename == "word/document.xml":
					self.xml = zin.read(item.filename)
					self.__find_all_fields()
					self.__find_all_bookmarks()
					return

	def __find_all_bookmarks(self):
		bookmark_ids = re.findall(r'<w:bookmarkStart w:id="(\d+)"', self.xml, re.DOTALL)
		self.bookmarks = {}
		for bookmark_id in bookmark_ids:
			regex = r'<w:bookmarkStart w:id="(%s)" w:name="([^"]+)"[^>]*>(.*?)<w:bookmarkEnd w:id="%s"/>' % (bookmark_id,bookmark_id)
			bookmark_xml = re.search(regex, self.xml, re.DOTALL)
			bookmark = Bookmark(*bookmark_xml.groups())
			self.bookmarks[bookmark.name] = bookmark

	def __find_all_fields(self):
		fields = re.findall(r'<w:fldChar w:fldCharType="begin"/>.*?(?=<w:fldChar w:fldCharType="(?:end|begin)"/>)',self.xml,re.DOTALL)
		self.sequences = {}
		self.references = OrderedDict()
		for field in fields:
			field = Field(field)
			if field.sequence:
				sequence = self.sequences.get(field.seq_name,[])
				sequence.append(field)
				self.sequences[field.seq_name] = sequence
			elif field.reference:
				reference = self.references.get(field.ref_name,[])
				reference.append(field)
				self.references[field.ref_name] = reference

	def check_fields_and_bookmarks(self):
		if len(self.sequences) == 1:
			figure_seq = self.sequences.keys()[0]
		else:
			seq_count = sorted([(x,len(y)) for x,y in self.sequences.items()],key=lambda (x,y):-y)
			seq_count = OrderedDict(seq_count).keys()
			figure_seq = seq_count[0]
			print "Most used sequence is '%s', others have been found: %s" % (seq_count[0],", ".join(seq_count[1:]))

		expected_nums = {}
		for idx,field in enumerate(self.sequences[figure_seq]):
			expected_nums[field] = str(idx+1)

		field_to_bookmark = {}
		field_references = {}
		for bookmark in self.bookmarks.values():
			if bookmark.field:
				field_to_bookmark[bookmark.field] = bookmark.name
				if bookmark.name in self.references:
					refs = self.references[bookmark.name]
					field_references[bookmark.field] = refs

		fig_no_bookmark = set(self.sequences[figure_seq]) - set(field_references.keys())
		if fig_no_bookmark:
			print "These Figures appear not to be mentioned: %s" % "\n\t".join(map(str,fig_no_bookmark))

		for field in self.sequences[figure_seq]:
			if not field in field_to_bookmark:
				print "Can't find corresponding bookmark for reference %s" % (field)
			if field.text != expected_nums[field]:
				print "Wrong field declaration, should be %s: %s" % (expected_nums[field],field)

		for name, refs in self.references.items():
			if not name in field_to_bookmark.values():
				continue
			field = self.bookmarks[name].field
			for ref in refs:
				if ref.text.replace("%s " % figure_seq,"") != field.text.replace("%s " % figure_seq,""):
					print "Wrong reference %s should be %s" % (ref,field.text)

		text = re.sub(r'<w:fldChar w:fldCharType="begin"/>.*?<w:fldChar w:fldCharType="end"/>','',self.xml,0,re.DOTALL)
		text = xml2txt(text,True)
		for fig_num in expected_nums.values():
			fig_name = "%s %s" % (figure_seq,fig_num)
			if fig_name in text:
				print "Reference to %s found in plain text" % fig_name


class Bookmark(EqualityChecker):
	def __init__(self,id,name,xml):
		self.id = int(id)
		self.name = name
		self.xml = xml
		self.nested = bool(re.search(r'<w:bookmarkStart',xml))
		self.text = xml2txt(xml)
		
		field_xml = re.search(r'<w:fldChar w:fldCharType="begin"/>.*?(?=<w:fldChar w:fldCharType="(?:end|begin)"/>)',xml,re.DOTALL)
		if field_xml:
			self.field = Field(field_xml.group(0))
		else:
			self.field = None

	def __repr__(self):
		field = str(self.field) if self.field else ""
		return "<bookmark id=%d name=%s>%s%s</bookmark>" % (self.id,self.name,field,self.text)

class Field(EqualityChecker):
	def __init__(self,xml):
		self.text = xml2txt(xml)
		instrText = re.search(r'<w:instrText[^>]*>(.*?)</w:instrText>',xml,re.DOTALL)
		if instrText:
			seq_name = re.search(r'SEQ (\w+)',instrText.group(1),re.DOTALL)
			if seq_name:
				self.sequence = True
				self.reference = False
				self.seq_name = seq_name.group(1)
				return
			self.sequence = False
			bookmark_name = re.search(r'REF (\w+)',instrText.group(1),re.DOTALL)
			if bookmark_name:
				self.reference = True
				self.ref_name = bookmark_name.group(1)
			else:
				self.reference = None

	def __repr__(self):
		if self.sequence:
			name = self.seq_name
		elif self.reference:
			name = self.ref_name
		else:
			name = None
		return "<field sequence=%s reference=%s name=%s>%s</field>" % (self.sequence,self.reference,name,self.text)

if __name__ == "__main__":
	import sys
	for docx in sys.argv[1:]:
		print "Checking file: %s" % docx
		d = Document(docx)
		d.check_fields_and_bookmarks()
		print
