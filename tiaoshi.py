import socket#服务端server
import time
import xlwt#只能写Excel
import xlrd#只能读Excel
import xlutils#修改Excel，在原来的基础上修改
import patterns as patterns
import pymysql
import wx
import threading
import sys
def RedExcel(i,j):#读取对应单元格内数据的执行函数
      book=xlrd.open_workbook("测试.xls")  #打开对应Excel
      #book.sheet_names()
      #查看文件中包含sheet的名称
      #得到第一个工作表，或者通过索引顺序或工作表名称
      #sheet = book.sheets()[0]
      #sheet = book.sheet_by_index(0)
      sheet = book.sheet_by_name(u'sheet1')
      a = sheet.cell(i,j).value#获取当前单元格的数据
      a = a+time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))+"\n"#将当前数据与当前单元格的数据进行追加写入
      book=xlwt.Workbook("测试.xls")#写入
      return a

def mysql_add(stu):#MySQL数据库添加数据函数
      global db,cursor
      #sql_add="insert into demo(begin,name,major,class,student_number,help_time,warn_time,end_time,time_span,over_time ) values(stu[0],stu[1],stu[2],stu[3],stu[4],stu[5],stu[6],stu[7],stu[8],stu[9])"
      sql_add='insert into demo values("{}","{}","{}","{}","{}","{}","{}","{}","{}","{}");'.format(stu[0],stu[1],stu[2],stu[3],stu[4],stu[5],stu[6],stu[7],stu[8],stu[9])
      try:
            # 执行SQL
            cursor.execute(sql_add)
            #事务提交
            db.commit()
            print("MySQL添加数据成功", file=contents)
            print("本次实验已记录，准备进行下一次实验，正在连接下位机......", file=contents)
      except Exception as err:
            #事务回滚
            db.rollback()
            print("MySQL添加数据失败！原因：",err, file=contents)

def data_acquisition(x):
      global sheet,book,s,s1,addr,stu
      while True:#死循环，一直接收读取
            #revice接收数据
            data=s1.recv(1024)
            #设定一次可以接收1024字节大小
            #打印接收到的数据
            data=format(data.decode('gbk'))
            print(type(data), file=contents)
            print(data, file=contents)
            #写入学生信息
            if data[2:5]=="姓名：":
                  print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), file=contents)
                  # 打印时间戳print(time.time())
                  sheet.write(x,0,time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())),style1)
                  pp=time.time()
                  #记录当前时间戳
                  print(data[5:8], file=contents)
                  sheet.write(x,1,data[5:8],style1)
                  print(data[13:16], file=contents)
                  sheet.write(x,2,data[13:16],style1)
                  print(data[21:25], file=contents)
                  sheet.write(x,3,data[21:25],style1)
                  print(data[30:], file=contents)
                  sheet.write(x,4,data[30:],style1)
                  sheet.write(x,5,"")
                  sheet.write(x,6,"")
                  sheet.write(x,7,"")
                  sheet.write(x,8,"")
                  stu[0]=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
                  stu[1]=data[5:8]
                  stu[2]=data[13:16]
                  stu[3]=data[21:25]
                  stu[4]=data[30:]
                  stu[5] = ""
                  stu[6]= ""
                  flag=True#标志位，代表写入了学生信息，才能写入后面的按键信息
                  book.save('测试.xls')#保存内容，放在最后，确保每次单元格内写入的内容有保存，然后再执行读取该单元格内容的函数，否则会报错
            elif data[4:6]=="帮助"and flag==True:
                  print(data[4:6], file=contents)
                  aa=RedExcel(x,5)#读取当前单元格数据并进行追加写的执行函数
                  #注意：本函数内的sheet.cell(i,j).value必须保证该单元格内有内容，否则会报错
                  #注：sheet.write(i,j,"数值")执行后光标会自动向后移动一个单元格，所以对于下一个单元格采用sheet.cell(i+1,j).value不会报错
                  sheet.write(x,5,aa,style1)#写入按下帮助按键的时间
                  book.save('测试.xls')#保存内容，放在最后，确保每次单元格内写入的内容有保存，然后再执行读取该单元格内容的函数，否则会报错
                  #第一次写入
                  if stu[5]=='-':
                        stu[5]=aa
                  else:
                        stu[5]=stu[5]+'  '+aa
            elif data[4:6]=="紧急"and flag==True:
                  print(data[4:6], file=contents)
                  bb=RedExcel(x,6)                  
                  sheet.write(x,6,bb,style1)#写入按下紧急按键的时间
                  book.save('测试.xls')#保存内容
                  #第一次写入
                  if stu[6]=='-':
                        stu[6]=bb
                  else:
                        stu[6]=stu[6]+'  '+bb
            elif data[3:5]=="完成" and flag==True:
                  print(data[3:5], file=contents)
                  print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), file=contents)
                  sheet.write(x,7,time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())),style1)#写入实验完成时间到单元格中
                  stu[7]=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
                  cc=time.time()-pp#时间戳相减，得到实验时长（结果为秒）
                  print(cc, file=contents)#输出当前秒数
                  if cc/60>45:
                        #超时
                        sheet.write(x,8,"时长：{}分{}秒（已超时）".format(int(cc/60),int(cc%60)),style1)
                        book.save('测试.xls')#保存内容
                        stu[8]="时长：{}分{}秒（已超时）".format(int(cc/60),int(cc%60))
                        stu[9]='YES'
                  else:
                        #未超时
                        sheet.write(x,8,"时长：{}分{}秒（未超时）".format(int(cc/60),int(cc%60)),style1)
                        book.save('测试.xls')#保存内容
                        stu[8]="时长：{}分{}秒（未超时）".format(int(cc/60),int(cc%60))
                        stu[9]='NOT'
                        mysql_add(stu)
                        #注意：要保存成结尾时.xls的文件，.xlsx用微软的文件打不开，只能用WPS的打开
                        s=socket.socket(socket.AF_INET,socket.SOCK_STREAM)#创建TCP Socket
                        #需要自己绑定一个ip地址和端口号
                        s.bind( ('192.168.0.107',8080) )#注：每次连接WIFI后IP地址可能发生改变
                        #将套接字绑定到地址, 在AF_INET下,以元组（host,port）的形式表示地址
                        #服务端监听操作时刻注意是否有客户端请求发来
                        s.listen(3)#可以同时监听3个，但是这里只有一个客户请求，因为没有写多线程
                        #开始监听TCP传入连接。backlog指定在拒绝连接之前，操作系统可以挂起的最大连接数量。该值至少为1，大部分应用程序设为5就可以了。
                        #同意连接请求
                        s1,addr=s.accept()#接受TCP连接并返回(s1,address),其中s1是新的套接字对象,可以用来接收和发送数据。address是连接客户端的地址。
                        #s 是服务端的socket对象s1是接入的客户端socket对象
                        print(addr)#显示连接客户端的地址，即下位机的ESP8266模块的IP地址
                  x=x+1#进入下一行
                  flag=False#学生信息采集成功标志位置为0
                  print("连接下位机成功！！！", file=contents)
                  print("当前光标写入行数：",x, file=contents)
                  #传过来的字节流需要用decode()解码，然后格式化format
                  book.save('测试.xls')#保存内容
                  
                  
                  

                  
def Font_Style_DIY():
      global font,style,style1#声明在本函数中的这些变量都是全局变量
      # 为样式创建字体
      font = xlwt.Font()
      # 字体类型
      font.name = 'name Times New Roman'
      # 字体颜色
      font.colour_index = 64
      # 字体大小，11为字号，20为衡量单位
      font.height = 20*11
      # 字体加粗
      font.bold = True
      # 下划线
      font.underline = True
      # 斜体字
      font.italic = False
      # 设置单元格对齐方式
      alignment = xlwt.Alignment()
      # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
      alignment.horz = 0x02
      # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
      alignment.vert = 0x01
      # 设置自动换行
      # alignment.wrap = 1
      # 设置边框
      borders = xlwt.Borders()
      # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
      # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
      borders.left = 4
      borders.right = 4
      borders.top = 4
      borders.bottom =4
      borders.left_colour =4
      borders.right_colour = 4
      borders.top_colour = 4
      borders.bottom_colour =4
      # 设置列宽，一个中文等于两个英文等于两个字符，11为字符数，256为衡量单位
      sheet.col(0).width = 18 * 256
      sheet.col(1).width = 8 * 256
      sheet.col(2).width = 8 * 256
      sheet.col(3).width = 8 * 256
      sheet.col(4).width = 15 * 256
      sheet.col(5).width = 20 * 256
      sheet.col(6).width = 20 * 256
      sheet.col(7).width = 20 * 256
      sheet.col(8).width = 22 * 256
      # 设置背景颜色
      pattern = xlwt.Pattern()
      # 设置背景颜色的模式
      pattern.pattern = xlwt.Pattern.SOLID_PATTERN
      # 背景颜色
      pattern.pattern_fore_colour = 40
      # 初始化样式
      style= xlwt.XFStyle()
      style.font = font
      style.pattern = pattern
      style.alignment = alignment
      style.borders = borders
      style1 = xlwt.XFStyle()
      style1.alignment = alignment#设置居中风格


class TransparentStaticText(wx.StaticText):
      #定义类，该类用于重构wx.StaticText，实现字体背景透明化
      """
      重写StaticText控件
      """
      #定义构造函数
      def __init__(self, parent, id=wx.ID_ANY, label='', pos=(400,50), size=(250,30),style=wx.TRANSPARENT_WINDOW, name='TransparentStaticText'):
            wx.StaticText.__init__(self, parent, id, label, pos, size, style, name)
            self.Bind(wx.EVT_PAINT, self.OnPaint)#调用下面的函数
            self.Bind(wx.EVT_ERASE_BACKGROUND, lambda event: None)#擦除背景
            self.Bind(wx.EVT_SIZE, self.OnSize)#调用下面的函数
 
      def OnPaint(self, event):
            bdc = wx.PaintDC(self)#重构wx.PaintDC
            dc = wx.GCDC(bdc)
            font1 = wx.Font(16, wx.SCRIPT, wx.SLANT, wx.LIGHT, False)
            dc.SetFont(font1)
            dc.SetTextForeground("Blue")
            dc.DrawText(self.GetLabel(), 0, 0)
 
      def OnSize(self, event):
            self.Refresh()
            event.Skip()




#事件处理函数
def OpenProgramming():
      global sheet,book,s,s1,addr,db,cursor,stu
      global button_flag
      button_flag=False
      while True:
            s1,addr=s.accept()#接受TCP连接并返回(s1,address),其中s1是新的套接字对象,可以用来接收和发送数据。address是连接客户端的地址。
            #s 是服务端的socket对象s1是接入的客户端socket对象
            print("正在连接下位机，请稍候......", file=contents)
            print("显示ESP88266的IP地址：",addr, file=contents)#显示连接客户端的地址，即下位机的ESP8266模块的IP地址
            font.num_format_str = '#,##0.00'
            sheet.write(0,0,u'实验开始时间',style)#指定行和列，写内容
            sheet.write(0,1,u'姓名',style)
            sheet.write(0,2,u'专业',style)
            sheet.write(0,3,u'班级',style)
            sheet.write(0,4,u'学号',style)
            sheet.write(0,5,u'是否按下帮助按键',style)
            sheet.write(0,6,u'是否按下紧急按键',style)
            sheet.write(0,7,'实验完成时间',style)
            sheet.write(0,8,'是否超时',style)
            x=1
            stu=['-','-','-','-','-','-','-','-','-','-']#设定存放列表
            print("当前光标写入行数：",x, file=contents)
            data_acquisition(x)#调用实验信息数据采集函数
            book.save('测试.xls')#保存内容
            s1.close()#关闭Excel表格，结束本条记录
            # 关闭游标
            cursor.close()
            # 关闭连接
            db.close()
            print("结束程序", file=contents)
            break
      
def OpenProgram(event):
      thread=threading.Thread(target=OpenProgramming)
      thread.start()#启动线程
    
def ExitProgramming():
      print("退出程序", file=contents)
      sys.exit(0)
      

      
def ExitProgram(event):
      thread2=threading.Thread(target=ExitProgramming)
      thread2.start()#启动线程

def main():
      global sheet,book,s,s1,addr,db,cursor,stu,contents#声明在本函数中的这些变量都是全局变量
      
      book=xlwt.Workbook()#新建一个Excel
      sheet=book.add_sheet('sheet1',cell_overwrite_ok=True)#建一个sheet页，设定可对单元格内容重复写入
      #创建socket对象
      s=socket.socket(socket.AF_INET,socket.SOCK_STREAM)#创建TCP Socket
      #需要自己绑定一个ip地址和端口号
      s.bind( ('192.168.0.107',8080) )#注：每次连接WIFI后IP地址可能发生改变
      #将套接字绑定到地址, 在AF_INET下,以元组（host,port）的形式表示地址
      #服务端监听操作时刻注意是否有客户端请求发来
      s.listen(3)#可以同时监听3个，但是这里只有一个客户请求，因为没有写多线程
      #开始监听TCP传入连接。backlog指定在拒绝连接之前，操作系统可以挂起的最大连接数量。该值至少为1，大部分应用程序设为5就可以了。
      #同意连接请求
      Font_Style_DIY()
      #获取数据库连接
      db=pymysql.connect(host='localhost', port=3306, user='root', passwd='', db='student_information', charset='utf8')
      #创建游标对象
      cursor = db.cursor()
      
      app = wx.App()
      win = wx.Frame(None,title="智能实验室管理系统",size=(1024,600)) #创建一个单独的窗口
      '''
      text = wx.StaticText(win,label="智能实验室管理系统",pos=(400,50),size=(250,30), style=wx.TRANSPARENT_WINDOW, name="staticText")
      font1 = wx.Font(16, wx.SCRIPT, wx.SLANT, wx.LIGHT, False)
      text.SetFont(font1)
      text.SetForegroundColour("Blue")
      '''
      text = TransparentStaticText(win,label="智能实验室管理系统")
      to_bmp_image = wx.Image( 'zz.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap()#导入背景图
     
      OpenButton = wx.Button(win,label="打开",pos=(800,180),size=(100,30))#定义按键显示位置
      ExitButton = wx.Button(win,label="退出",pos=(800,280),size=(100,30))
     
      OpenButton.Bind(wx.EVT_BUTTON,OpenProgram)#按键绑定事件
      ExitButton.Bind(wx.EVT_BUTTON,ExitProgram)
     
      contents=wx.TextCtrl(win,pos=(250,100),size=(500,300),style=wx.TE_MULTILINE )#显示框
      win.Show()#显示按钮
      wx.StaticBitmap(win, -1, to_bmp_image, (0, 0))#图片起始坐标地址（0,0）
      app.MainLoop() #进入应用程序事件主循环
      

if __name__== '__main__':

      main()



