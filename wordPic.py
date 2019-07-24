# _*_ coding: UTF-8 _*_
'this is a file tools'
__author__ = 'TB'
import json
import re
import datetime
from PyQt5.QtWidgets import *
from From import Ui_Form
import sys
import xml.etree.ElementTree as ET
import os
from win32com.client import Dispatch
class setUi(QWidget,Ui_Form):
    def __init__(self,parent = None):
        super(setUi, self).__init__(parent)
        self.setupUi(self)
        self.initUi()
        self.picW = '0pt'  # 设置图片的宽度
        self.picH = '0pt'  # 设置图片高度
        self.picSize = ''   # 要替换图片的属性
        self.regw = ''  # 用正则匹配出宽度数据
        self.regh = ''  # 用正则匹配出的高度数据
        self.mixw = '0pt'  # 设置需要替换的最小宽度
        self.maxw = '0pt'  # 设置需要替换的最大宽度
        self.mixh = '0pt'  # 设置需要替换的最小高度
        self.maxh = '0pt'  # 设置需要替换的最大高度
        self.getval = ''    # 获取的章节参数
        self.xml1 = ''
        self.setAcceptDrops(True)
    def dragEnterEvent(self, evn):
        links = evn.mimeData().text()
        a = r'(?<=file:///).+'
        link = re.search(a,links).group()
        self.lineEdit.setText(link)

        evn.accept()

    def initUi(self):
        self.pushButton.clicked.connect(self.getfile)   #设置文件选择信号连接
        self.pushButton_2.clicked.connect(self.main)     #设置开始信号连接
        self.pushButton_3.clicked.connect(self.saveBin)
        self.pushButton_4.clicked.connect(self.loadBin)
        self.lineEdit_2.textChanged.connect(self.getvalue)  #设置章节获取信号连接
        self.doubleSpinBox.valueChanged.connect(self.valuechangeW)  #设置 图片宽度信号
        self.doubleSpinBox_2.valueChanged.connect(self.valuechangeH)    #设置图片宽度信号
        self.doubleSpinBox_2.valueChanged.connect(self.valuechangeH)
        self.doubleSpinBox_minw.valueChanged.connect(self.mix_w)
        self.doubleSpinBox_maxw.valueChanged.connect(self.max_w)
        self.doubleSpinBox_mixh.valueChanged.connect(self.mix_h)
        self.doubleSpinBox_maxh.valueChanged.connect(self.max_h)

    def saveBin(self):
        one = {'self.picW': self.doubleSpinBox.value(),
               'self.picH': self.doubleSpinBox_2.value(),
               'self.mixw': self.doubleSpinBox_minw.value(),
               'self.maxw': self.doubleSpinBox_maxw.value(),
               'self.mixh': self.doubleSpinBox_mixh.value(),
               'self.maxh': self.doubleSpinBox_maxh.value(),
               'self.getval': self.lineEdit_2.text()
               }
        date = json.dumps(one)
        with open('Bin.json','w') as f:
            f.write(date)
            f.close()
        self.label_18.setText('保存成功')


    def loadBin(self):
        with open('Bin.json','r') as g:
            temp = json.loads(g.read())
            self.doubleSpinBox.setValue(temp['self.picW'])
            self.doubleSpinBox_2.setValue(temp['self.picH'])
            self.doubleSpinBox_minw.setValue(temp['self.mixw'])
            self.doubleSpinBox_maxw.setValue(temp['self.maxw'])
            self.doubleSpinBox_mixh.setValue(temp['self.mixh'])
            self.doubleSpinBox_maxh.setValue(temp['self.maxh'])
            self.lineEdit_2.setText(temp['self.getval'])









        self.label_16.setText('读取成功')


    def mix_w(self):
        self.mixw = str(self.doubleSpinBox_minw.value()*2.832861189)
    def max_w(self):
        self.maxw = str(self.doubleSpinBox_maxw.value()*2.832861189)
    def mix_h(self):
        self.mixh = str(self.doubleSpinBox_mixh.value()*2.832861189)
    def max_h(self):
        self.maxh = str(self.doubleSpinBox_maxh.value()*2.832861189)
    def getfile(self):
        self.choosefile = QFileDialog.getOpenFileName(self,'请选择word文档',r'C:\Users\My\Desktop','office file (*.doc *.docx *.xml)')
        self.lineEdit.setText(self.choosefile[0])
        print('获取文件',self.lineEdit.text())
        self.textBrowser.append('获取文件成功')
    def getvalue(self):
        self.getval = self.lineEdit_2.text()
        print('获取章节名称',self.lineEdit_2.text())
        self.textBrowser.append('获取章节名称成功')
    def valuechangeW(self):
        self.picW = '{}pt'.format(str((self.doubleSpinBox.value())*2.832861189))



    def valuechangeH(self):
        self.picH = '{}pt'.format(str((self.doubleSpinBox_2.value())*2.832861189))
        print('设置图片高度为',self.picH)

    def converXml(self,dirfiles):

        time1 = datetime.datetime.now()
        try:
            if os.path.splitext(dirfiles)[1] in ['.doc', '.docx']:

                app = Dispatch('Word.Application')
                self.textBrowser.append('进程创建...')
                doc = app.Documents.Open(dirfiles)
                self.textBrowser.append('打开成功...')
                doc.SaveAs(os.path.splitext(dirfiles)[0] + '.xml', 11)
                self.textBrowser.append('保存成xml成功...')
                doc.Close()
                app.Quit()
                time2 = datetime.datetime.now()

                return (os.path.splitext(dirfiles)[0] + '.xml')
            elif os.path.splitext(dirfiles)[1] in ['.xml']:
                # print(os.path.splitext(dirfiles)[1])
                app = Dispatch('Word.Application')
                doc = app.Documents.Open(dirfiles)
                doc.SaveAs(os.path.splitext(dirfiles)[0] + '.docx', 12)
                doc.Close()
                app.Quit()
                os.remove(dirfiles)
                time2 = datetime.datetime.now()
                self.textBrowser.append('转换用时' + str(time2 - time1))
                self.textBrowser.append('完成')
                return (os.path.splitext(dirfiles)[0] + '.docx')
        except:
            self.textBrowser.append('不支持该文件')
        else:
            self.textBrowser.append('请选择word文档')

    def wordEdit(self,dirs):
        try:
            picWW = re.compile('(?<=width:).*?(pt|in)')
            picHH = re.compile('(?<=height:).*?(pt|in)')
            wx_sub_section = '{http://schemas.microsoft.com/office/word/2003/auxHint}sub-section'
            w_t = '{http://schemas.microsoft.com/office/word/2003/wordml}t'
            pic_att = '{urn:schemas-microsoft-com:vml}shape'
            tree = ET.parse(dirs)
            root = tree.getroot()
            for a in root.iter(wx_sub_section):
                for b in a.iter(w_t):
                    if b.text == r'—{}'.format(self.getval) or b.text == r'、{}'.format(
                            self.getval) or b.text == r'{}'.format(self.getval):
                        print('找到了')
                        for pic in a.iter(pic_att):
                            print('原始', pic.attrib['style'])
                            self.picSize = pic.attrib['style']
                            # self.textBrowser.append('找到图片属性', self.picSize)
                            self.regw = picWW.search(self.picSize).group()
                            print('匹配到宽', self.regw)
                            # self.textBrowser.append('匹配到宽度', self.regw)
                            self.regh = picHH.search(self.picSize).group()
                            print('匹配到高', self.regh)
                            # self.textBrowser.append('匹配到高度', self.regh)
                            if (self.regw >= self.mixw or self.regw <= self.maxw) and \
                                    (self.regh >= self.mixh or self.regh <= self.maxh):
                                # print('待替换宽',self.picW)
                                resultw = re.sub(self.regw, self.picW, self.picSize)
                                print('替换成功宽度', resultw)
                                # print('待替换高度',self.picH)
                                resultwh = re.sub(self.regh, self.picH, resultw)
                                print('替换成功高度', resultwh)
                                pic.set('style', resultwh)
                                print('新的属性是', pic.attrib['style'])
            tree.write(os.path.splitext(self.lineEdit.text())[0] + '.xml')
            self.textBrowser.append('word操作完成\n正在保存请稍等...')
            return dirs

        except:
            print('word操作失败')
            self.textBrowser.append('word操作失败请重试')


    def main(self):
        s = re.sub('\/', '\\\\', self.lineEdit.text())

        self.converXml(self.wordEdit(self.converXml(s)))














if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = setUi()
    w.show()
    sys.exit(app.exec())

