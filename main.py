# -*- coding: utf-8 -*-
import xml.dom.minidom
import xlrd
import datetime
def readexl():
    data = xlrd.open_workbook('时间位置.xlsx')
    table = data.sheet_by_name('Sheet1')
    nRow = table.nrows
    nCol = table.ncols
    list1 = []

    for i in range(1,nRow):
        point = {}
        for j in range(nCol):
            value = table.row_values(i)[j]
            if j == 0:
                timecon = str(value)
                timecon =str(datetime.datetime.now().year)+ '-'+timecon[-12:-10]+'-'+timecon[-10:-8] +'T'+ timecon[-8:-6]+':'+ timecon[-6:-4] + ':'+ timecon[-4:-2] + 'Z'

                point['time'] = timecon
            if j == 3:
                lat = float(value)/3600
                point['lat'] = str(lat)
            if j == 2:
                lon = float(value)/3600
                point['lon'] = str(lon)
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
    nodetrk = doc.createElement('trk')
    nodetrkname = doc.createElement('name')
    nodetrkname.appendChild(doc.createTextNode('Track'))
    nodetrk.appendChild(nodetrkname)
    nodetrkseg = doc.createElement('trkseg')
    for i in checklist :
        nodetrkpt = doc.createElement('trkpt')
        nodetrkpt.setAttribute('lat', str(i['lat']))
        nodetrkpt.setAttribute('lon', str(i['lon']))
        #给叶子节点name设置一个文本节点，用于显示文本内容
        nodeele = doc.createElement('ele')
        nodeele.appendChild(doc.createTextNode('0'))
        nodetrkpt.appendChild(nodeele)
        nodetime = doc.createElement('time')
        nodetime.appendChild(doc.createTextNode(i['time']))
        nodetrkpt.appendChild(nodetime)
        nodetrkseg.appendChild(nodetrkpt)
        print(i['time']+'    '+i['lat']+'    '+i['lon'])



    nodetrk.appendChild(nodetrkseg)
    root.appendChild(nodetrk)
#开始写xml文档
    fp = open('GPXresult.gpx', 'w')
    doc.writexml(fp, indent='\t', addindent='\t', newl='\n', encoding="utf-8")







if __name__ == '__main__':
    managerList = readexl()
    toxml(managerList)
    close =input('转换成功，点击任意按键关闭窗口')
    if close is not None:
        exit()

 