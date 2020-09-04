# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'F:\MyPython\MODEL.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

import openpyxl 
import os
import pandas
from pandas import DataFrame
import json

import pymysql

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *

from PIL import Image

import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter   # 设置坐标轴显示方式
## 正常显示中文设置
from pylab import *
mpl.rcParams['font.sans-serif'] = ['SimHei']



class Spindle_Wnd(QWidget):
    '''荣威名爵品牌残值程序'''
    def __init__(self,parent = None):
        super(Spindle_Wnd,self).__init__(parent)
        # 从EXCEL初始化需要匹配的数据
        self.initUi()
        self.resultDict = {}
        self.fileNameList=[]

    def initUi(self):
        mainlayout = QVBoxLayout(self)
        self.scroll = QScrollArea()

        self.topcontainer = QWidget(self.scroll)
        self.topcontainer.setMinimumSize(850,600)
        self.topcontainer.setSizePolicy(QSizePolicy.Expanding,QSizePolicy.Expanding)
        toplayout = QGridLayout(self.topcontainer)

        # 1）NTF文件选择
        self.aboutBtn = QPushButton("程序介绍")
        self.aboutBtn.clicked.connect(self.aboutDlg)
        toplayout.addWidget(self.aboutBtn,0,6,1,2)

        
        toplayout.addWidget(QLabel("导入EXCEL数据"),1,0,1,1)
        self.compressCheck = QCheckBox()
        self.compressCheck.setChecked(True)
        self.compressCheck.setText("多条曲线")
        toplayout.addWidget(self.compressCheck,5,3,1,1)
        self.jpgPath = QLineEdit()
        toplayout.addWidget(self.jpgPath,1,2,1,4)
        self.jpgButton = QPushButton("选择EXCEL数据文件")
        self.jpgButton.clicked.connect(self.SelectExcelFile)
        toplayout.addWidget(self.jpgButton,1,6,1,2)

        # 42) 加载数据曲线
        toplayout.addWidget(QLabel(""),2,0,1,3)
        toplayout.addWidget(QLabel("纵梁传递曲线显示"),3,0,1,3)
        self.projectCombo = QComboBox()
        label = QLabel("项目:")
        label.setAlignment(Qt.AlignRight)
        toplayout.addWidget(label,4,0,1,1)
        toplayout.addWidget(self.projectCombo,4,1,1,1)
        label = QLabel("阶段:")
        label.setAlignment(Qt.AlignRight)
        toplayout.addWidget(label,4,2,1,1)
        self.stageCombo = QComboBox()
        toplayout.addWidget(self.stageCombo,4,3,1,1)
        self.locationCombo = QComboBox()
        self.locationCombo.addItems(["","左前纵梁","右前纵梁","左后纵梁","右后纵梁"])
        self.locationCombo.currentTextChanged.connect(self.selectLocation)
        label = QLabel("位置:")
        label.setAlignment(Qt.AlignRight)
        toplayout.addWidget(label,4,4,1,1)
        toplayout.addWidget(self.locationCombo,4,5,1,1)
        label = QLabel("占位符:")
        label.setAlignment(Qt.AlignRight)
        toplayout.addWidget(label,4,6,1,1)
        self.dllxCombo = QComboBox()
        toplayout.addWidget(self.dllxCombo,4,7,1,1)
        label = QLabel("选择方向:")
        label.setAlignment(Qt.AlignRight)
        toplayout.addWidget(label,5,0,1,1)
        self.directionCombo = QComboBox()
        toplayout.addWidget(self.directionCombo,5,1,1,2)


        self.saveImgBtn = QPushButton("保存图例")
        self.saveImgBtn.clicked.connect(self.saveImg)
        toplayout.addWidget(self.saveImgBtn,5,5,1,1)
        self.showBtn = QPushButton("查看")
        self.showBtn.clicked.connect(self.showCurve)
        toplayout.addWidget(self.showBtn,5,6,1,2)
        

        # 4.1）显示纵梁模型
        self.curveWidget = QWidget()
        self.curveWidget.setMinimumSize(800,300)
        self.curveLayout = QHBoxLayout()
        self.curveWidget.setLayout(self.curveLayout)
        toplayout.addWidget(self.curveWidget,6,0,1,8)
        self.curvefig = plt.figure()
        # plt.axis("off")
        self.curveAxs = self.curvefig.add_subplot(111)
        self.curveCanvas = FigureCanvas(self.curvefig)
        self.curveCanvas.draw()
        self.curveLayout.addWidget(self.curveCanvas)

        self.scroll.setWidget(self.topcontainer)
        self.scroll.show()
        mainlayout.addWidget(self.scroll)
        
        # 初始化项目选择项
        # 1.1) 连接数据库
        conn = pymysql.connect(host="127.0.0.1",port=3306,user="root",password="123456",db="sd")
        if not conn:
            raise(NameError, "连接数据库失败！")
        cursor = conn.cursor()
        sql = "select distinct project from sd.zlcd_result"
        cursor.execute(sql)
        row = cursor.fetchall()
        rowList = [x[0] for x in row]
        self.projectCombo.addItems(rowList)
        #关闭游标
        cursor.close()
        #关闭连接
        conn.close() 


        self.resize(920,750)
    

    def aboutDlg(self):
        '''程序说明'''
        about = '''荣威名爵品牌残值处理程序\n1.请将EXCEL文件数据放在指定的同一个文件夹下，深层文件夹里的EXCEL数据不会被读取\n
        2.数据读取前请不要打开文件夹下的将要读取的EXCEL数据\n
        3.EXCEL命名规划为下划线“_” + 精真估月份\n
        4.一个EXCEL文件中包含两个表格数据，Sheet1面向商家，Sheet2面向客户，不要弄混\n
        5.如果需要多条曲线显示，请勾选多条显示方框\n
        '''
        reply = QMessageBox.information(self,"关于",about,QMessageBox.Close)
    
    def selectLocation(self,str):
        pro = self.projectCombo.currentText()
        location = "%" + str[0:2] + "%"

        # 1.1) 连接数据库
        conn = pymysql.connect(host="127.0.0.1",port=3306,user="root",password="123456",db="sd")
        if not conn:
            raise(NameError, "连接数据库失败！")
        cursor = conn.cursor()

        sql = "select channel from sd.zlcd_result where project = '%s' and channel like '%s'" %(pro,location)
        # print(sql)
        cursor.execute(sql)
        rowall = cursor.fetchall()
        #关闭游标
        cursor.close()
        #关闭连接
        conn.close() 

        channel = [x[0] for x in rowall]
        self.directionCombo.clear()
        self.directionCombo.addItems(channel)


    def saveImg(self):
        self.curvefig.savefig("saveimg.png")

    def SelectExcelFile(self):
        '''选择Excel文件'''
        # 1) 选择图片文件夹
        path = self.jpgPath.text()

        try:
            fName,_ = QFileDialog.getOpenFileName(self,"选择纵梁传递数据库文件",path,"Excel files(*.xls *.xlsx)")
        except ValueError:
            return

        if not fName:
            exit(0)

        # 1.1) 连接数据库
        conn = pymysql.connect(host="127.0.0.1",port=3306,user="root",password="123456",db="sd")
        if not conn:
            raise(NameError, "连接数据库失败！")
        cursor = conn.cursor()

        # 2) 读取EXCEL表格
        book = openpyxl.load_workbook(fName)
        # 2.1）遍历工作表
        projectList = []
        for sht in book.worksheets:
            shtName = sht.title
            print(shtName)
            # 2.2） 判断表格是否为数据
            if "纵梁传递导纳" not in shtName:
                continue
            project = shtName[:shtName.find("纵梁传递")]
            projectList.append(project)
            # 2.3） 判断数据是否统一
            if sht["B5"].value != "(m/s)/N":
                print("{}数据单位不对，请核实".format(shtName))
                continue
            if sht["A8"].value != 2:
                print("{}数据单位不对，请核实".format(shtName))  
                continue             
            if sht["A107"].value != 200 or sht["A108"].value != None:
                print("{}数据单位不对，请核实".format(shtName))
                continue
            # print(sht["A108"].value)
            # 2.4) 开始读取数据
            X = []
            vehno = ""
            wj = None
            X.append(project)
            X.append(vehno)
            X.append(wj)

            data = list(sht.columns)
            X.append("X")
            Xdata = [x.value for x in data[0][6:107]]
            X.append(json.dumps(Xdata,ensure_ascii=False))
            # print(X)
            # 2.5）插入X数据
            # 2.5.1） 先查询X数据是否存在
            sql = "select * from sd.zlcd_result where project = '%s' and channel = '%s'" %(X[0],X[3])
            # print(sql)
            cursor.execute(sql)
            rowall = cursor.fetchall()
            if rowall:
                pass
            else:
                sql="INSERT INTO sd.zlcd_result(project,vehno,wj,channel,data) VALUES(%s,%s,%s,%s,%s)" 
                cursor.execute(sql,X)
                conn.commit()
            
            # print(len(data))
            for id in range(1,25):
                Y = []
                Y.append(project)
                Y.append(vehno)
                Y.append(wj)
                Ycurve = {}
                # 判断是否有数据
                if data[id][6].value == None:
                    continue 
                YdataName = data[id][3].value
                Y.append(YdataName)            
                Ydata = [x.value for x in data[id][6:107]]
                Ycurve["Unit"] = data[id][4].value
                Ycurve["Linear"] = data[id][5].value
                Ycurve["data"] = Ydata
                Yvaule = json.dumps(Ycurve,ensure_ascii=False)
                Y.append(Yvaule)
                
                # 将数据存入数据库
                # 2.5.1） 先查询X数据是否存在
                sql = "select * from sd.zlcd_result where project = '%s' and channel = '%s'" %(Y[0],Y[3])
                cursor.execute(sql)
                rowall = cursor.fetchall()
                if rowall:
                    pass
                else:
                    sql="INSERT INTO sd.zlcd_result(project,vehno,wj,channel,data) VALUES(%s,%s,%s,%s,%s)" 
                    cursor.execute(sql,Y)
                    conn.commit()

            print("完成{}纵梁数据写入".format(project))

        #关闭游标
        cursor.close()
        #关闭连接
        conn.close()  

        # 初始化项目选择项
        self.projectCombo.insertItems(0,projectList)
                      

    def showCurve(self):
        pro = self.projectCombo.currentText()
        channel = self.directionCombo.currentText()

        #  查询数据并获取曲线
        # 1.1) 连接数据库
        conn = pymysql.connect(host="127.0.0.1",port=3306,user="root",password="123456",db="sd")
        if not conn:
            raise(NameError, "连接数据库失败！")
        cursor = conn.cursor()

        # 1.2) 获取Y轴数据
        sql = "select data from sd.zlcd_result where project = '%s' and channel = '%s'" %(pro,channel)
        # print(sql)
        cursor.execute(sql)
        row = cursor.fetchone()
        # 1.3） 提取Y轴数据
        if row:
            rowDict = json.loads(row[0])
            Ycurve = rowDict["data"]
        else:
            return

        # 1.4) 获取X轴数据
        sql = "select data from sd.zlcd_result where project = '%s' and channel = 'X'" %(pro)
        cursor.execute(sql)
        row = cursor.fetchone()
        if row:
            Xcurve = json.loads(row[0])
        else:
            return

        #关闭游标
        cursor.close()
        #关闭连接
        conn.close() 

        # 如果多条曲线显示，就是重绘曲线
        check_state = self.compressCheck.checkState()
        if check_state == Qt.Checked:
            pass
        elif check_state == Qt.Unchecked:
            self.curveAxs.cla()
            # plt.cla()

        self.curveAxs.plot(Xcurve,Ycurve,label = pro + ":" + channel)
        self.curveAxs.set_ylabel("(m/s)/N\n Amplitude")
        self.curveAxs.set_xlabel("Hz")
        plt.legend()

        self.curveCanvas.draw()
        


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    widget = QtWidgets.QWidget()
    widget.setWindowTitle("纵梁传递数据库文件读取")
    ui = Spindle_Wnd(widget)
    widget.show()
    # ui.setupUi(widget)
    # widget.show()
    
    sys.exit(app.exec_())
 


   