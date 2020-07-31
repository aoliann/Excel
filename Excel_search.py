import xlwings as xw #引入xlwings处理excel
from tkinter import * #引入tkinter处理界面
import tkinter.filedialog #调用上传文件窗口类
import math
import threading #多线程
import string
 
#基本变量
files=[]
result=[]
 
def search(file,word):
    app = xw.App(visible=False,add_book=False) #初始化    visible=False表示不在用户界面打开该excel   add_book=False只打开不新建工作薄
    wb = app.books.open(file) #打开已有excel文件
    sheets = wb.sheets #读取工作簿
    for sht in sheets:
        rng = sht.range('a1').expand('table')
        nrows = rng.rows.count #获取行数
        ncols = rng.columns.count #获取列数
        print("正在比对：%s"%(file))
        #文本转字符
        A=math.floor(ncols/26) #首字母
        A=chr((A-1)+97) if (A-1)>=0 else "" #如26以内首字母为空   27首字母为A  55首字母为B
        B=math.floor(ncols%26) #第二个字母 取26的余                按这个算法 B为第2列  Z为第26列    AA 为第27列    AB为第28列    BA为第53列
        B=chr(B+97) if B>=0 else ""
        col=("%s%s"%(A,B)) #合并
        val=sht.range('a1:%s%d'%(col,nrows)).value #获取区域内容
        for ii,i in enumerate(val): #遍历区域内容
            for jj,j in enumerate(i):
                if str(j).find(str(word))!=-1: #搜索该字符串
                    mess="成功找到数据位于表：%s 工作簿：%s 行：%d 列：%d 内容：%s"%(file,sht,ii+1,jj+1,j)
                    result.append(mess)
                    print(mess)
    app.quit() #退出excel
 
def MyThreading(num1,num2,searchText): #线程执行
    for i in range(num1,num2):
        print("打开文件%s"%(files[i]))
        search(files[i],searchText)
    return 0
 
def FileSelect():#文件选择
    files.clear()
    filename = tkinter.filedialog.askopenfilenames() #弹出多文件选择窗
    if(len(filename)==0):
        textbox.insert('end', "未选择任何文件！\n")
        root.update()
        print("未选择任何文件！\n")
    else:
        textbox.insert('end', "已选择文件\n")
        for i in filename:
            files.append(i)
            print("\t\t%s\n"%(i))
            textbox.insert('end', "%s\n"%(i))
            textbox.see(tkinter.END)
            root.update()
    return 0
 
def start():
    searchText=e.get() #从钩子获取内容
    if len(files)==0:
        print("未上传文件")
        textbox.insert('end', "错误：未上传文件\n")
        root.update()
        return 0
    if (not searchText) or searchText=="请输入需要搜索的字符串":
        print("未输入需查询内容")
        textbox.insert('end', "错误：未输入需查询内容\n")
        root.update()
        return 0
    print("正在执行")
    textbox.insert('end', "正在执行,页面可能稍有卡顿，请耐心等待\n")
    root.update()
 
    ThreadLen=10 if len(files)>10 else 1 #文件数量若小于10则只开一个线程
    ThreadArr=[]
    if ThreadLen>1:
        for i in range(ThreadLen):#追加十个线程（若需要）
            bei=math.floor(len(files)/ThreadLen)
            print("执行线程%d至%d"%((i*bei),(i*bei+bei)))
            t=threading.Thread(target=MyThreading,args=((i*bei),(i*bei+bei),searchText))
            ThreadArr.append(t)
 
    #追加线程执行剩余文件搜索
    print("执行线程%d至%d"%(0,len(files)))
    t=threading.Thread(target=MyThreading,args=(0,len(files),searchText))
    ThreadArr.append(t)
    for i in ThreadArr:
        i.start() #循环执行所有线程
    for i in ThreadArr:
        i.join() #循环等待所有线程执行结束
 
    print("执行结束")
 
    if len(result)==0:
            textbox.insert('end', "没有找到该字符串\n")
    else:
        textbox.insert('end', "执行结束，执行结果如下：\n")
    for i in result:
        textbox.insert('end', "%s\n"%(i))
    result.clear()
    textbox.see(tkinter.END) #查看文本框底部
    root.update() #更新文本框
 
 
 
 
#构建用户界面
root=Tk()
root.title("Excel关键字搜索") #窗口标题
root.geometry('450x300+585+265') #初始化窗口大小和位置
root.resizable(0, 0) #设置窗口大小不可更改
 
#输入框
e = StringVar()#字符串钩子
enrty=Entry(root,width=20,textvariable=e).grid(row=0,column=0,columnspan=2,padx=40,pady=5)
e.set("请输入需要搜索的字符串")
 
#上传文件按钮
getFile = Button(root,text="点击选择文件",command= FileSelect).grid(row=1,column=0,columnspan=2,padx=40,pady=5)
 
#开始按钮
startBtn = Button(root,text="开始查找",command= start).grid(row=2,column=0,columnspan=2,padx=40,pady=5)
 
#消息框
textbox = Text(root,width=50,height=10)
textbox.grid(row=3,column=0,columnspan=2,padx=40,pady=5)
textbox.insert("end", "日志消息：\n")
 
root.mainloop() #执行
