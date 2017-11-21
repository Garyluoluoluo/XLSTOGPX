# -*- coding: utf-8 -*-
import xml.dom.minidom
import xlrd
def readexl():
    data = xlrd.open_workbook('toGPX.xls')
    table = data.sheet_by_name('Sheet1')
    nRow = table.nrows
    nCol = table.ncols
    list1 = []

    for i in range(nRow):
        point = {}
        for j in range(nCol):
            value = table.row_values(i)[j]
            if j == 0:
                point['name'] = str(value)
            if j == 1:
                point['lat'] = str(value)
            if j == 2:
                point['lon'] = str(value)
        list1.append(point)
    return list1

def toxml(checklist):
    #在内存中创建一个空的文档
    doc = xml.dom.minidom.Document()
    #创建一个根节点Managers对象
    root = doc.createElement('gpx')
    #设置根节点的属性
    root.setAttribute('xmlns', 'http://www.topografix.com/GPX/1/1')
    root.setAttribute('creator', 'Garmin Desktop App')
    root.setAttribute('version', '1.1')
    root.setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')
    root.setAttribute('xsi:schemaLocation', 'http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd')
    #将根节点添加到文档对象中
    doc.appendChild(root)

    #managerList = [{'name' : 'dian1',  'lat' : '30.513906525447965', 'lon' : '104.05195640400052'},{'name' : 'dian2', 'lat' : '30.512208770960569', 'lon' : '104.06846079044044'},{'name' : 'dian3', 'lat' : '30.510888202115893', 'lon' : '104.0592603944242'}]
    noderte = doc.createElement('rte')
    nodertename = doc.createElement('name')
    nodertename.appendChild(doc.createTextNode('Line'))
    noderte.appendChild(nodertename)
    for i in checklist :
        nodertept = doc.createElement('rtept')
        nodertept.setAttribute('lat', str(i['lat']))
        nodertept.setAttribute('lon', str(i['lon']))
        #给叶子节点name设置一个文本节点，用于显示文本内容
        noderteptname = doc.createElement("name")
        noderteptname.appendChild(doc.createTextNode(str(i['name'])))
        nodedesc = doc.createElement("desc")
        nodedesc.appendChild(doc.createTextNode(''))

        nodesym = doc.createElement("sym")
        nodesym.appendChild(doc.createTextNode('unistrong:104'))


        nodertept.appendChild(noderteptname)
        nodertept.appendChild(nodedesc)
        nodertept.appendChild(nodesym)
        noderte.appendChild(nodertept)
    root.appendChild(noderte)
#开始写xml文档
    fp = open('GPXresult.gpx', 'w')
    doc.writexml(fp, indent='\t', addindent='\t', newl='\n', encoding="utf-8")







if __name__ == '__main__':
    managerList = readexl()
    toxml(managerList)
 