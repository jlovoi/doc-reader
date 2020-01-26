from lxml import etree
import zipfile

ooXMLns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# docxZip = zipfile.ZipFile('tpsReport_v69_420_vape.docx')
docxZip = zipfile.ZipFile('joe_edit.docx')

# print(docxZip.infolist())

# commentsXML = docxZip.read('word/comments.xml')

commentsXML = docxZip.read('word/comments.xml')
et = etree.XML(commentsXML)

# './w:ins//w:p' for inserted lines

print("COMMENTS")
# comments = et.xpath('//w:comment', namespaces=ooXMLns)
comments = et.xpath('//w:comment', namespaces=ooXMLns)
for c in comments:
    # attributes:
    print(c.xpath('@w:author', namespaces=ooXMLns))
    print(c.xpath('@w:date', namespaces=ooXMLns))
    # string value of the comment:
    print(c.xpath('string(.)', namespaces=ooXMLns))


insertsXML = docxZip.read('word/document.xml')
te = etree.XML(insertsXML)

# './w:ins//w:p' for inserted lines

print("\n\nINSERTS\n\n")
# comments = et.xpath('//w:comment', namespaces=ooXMLns)
inserts = te.xpath('//w:p', namespaces=ooXMLns)
for i in inserts:
    if (i.xpath('w:ins', namespaces=ooXMLns)):

        # print(i.xpath('@w:author', namespaces=ooXMLns))
        # print(i.xpath('@w:date', namespaces=ooXMLns))
        print(i.xpath('string(.)', namespaces=ooXMLns))
        ins = i.xpath('w:ins', namespaces=ooXMLns)
        for ting in ins:
            print(ting.xpath('@w:author', namespaces=ooXMLns))
            print(ting.xpath('@w:date', namespaces=ooXMLns))
            print(ting.xpath('string(.)', namespaces=ooXMLns))
        print("\n\nBREAK\n\n")
        # print(i.xpath('@w:rsidR', namespaces=ooXMLns))
