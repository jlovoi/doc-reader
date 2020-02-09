import os
import json

from lxml import etree
import zipfile

fileName = 'joe_edit.docx'

ooXMLns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
docxZip = zipfile.ZipFile(fileName)
commentsXML = docxZip.read('word/comments.xml')
et = etree.XML(commentsXML)

extraction = {'comments': [], 'inserts': []}

comments = et.xpath('//w:comment', namespaces=ooXMLns)
for c in comments:
    authors = c.xpath('@w:author', namespaces=ooXMLns)
    date = c.xpath('@w:date', namespaces=ooXMLns)
    content = c.xpath('string(.)', namespaces=ooXMLns)
    extraction['comments'].append({
        'authors': authors,
        'date': date,
        'content': content
    })

insertsXML = docxZip.read('word/document.xml')
te = etree.XML(insertsXML)

inserts = te.xpath('//w:p', namespaces=ooXMLns)
x = 0
for i in inserts:
    if (i.xpath('w:ins', namespaces=ooXMLns)):
        line = i.xpath('string(.)', namespaces=ooXMLns)
        extraction['inserts'].append({
            'line': line
        })
        ins = i.xpath('w:r|w:ins', namespaces=ooXMLns)
        inserted = []
        for ting in ins:
            author = ting.xpath('@w:author', namespaces=ooXMLns)
            date = ting.xpath('@w:date', namespaces=ooXMLns)
            content = ting.xpath('string(.)', namespaces=ooXMLns)
            inserted.append({
                'author': author,
                'date': date,
                'content': content
            })
        extraction['inserts'][x]['inserted'] = inserted
        x += 1

print(extraction)

jsonObj = json.dumps(extraction)

writeTo = fileName.split('.')[0] + '.json'

f = open(writeTo, 'w')
f.write(jsonObj)
f.close()
