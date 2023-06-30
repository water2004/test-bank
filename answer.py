from tkinter import *
import tkinter
# from tkinter.ttk import *
import pandas as pd
import openpyxl
import ctypes

err=[]
onlyerr=False
flag=False
index=-1
id=0
multi=[0,0,0,0]
index=0

def write_err():
    f=open('err.txt','a')
    f.write(str(id)+'\n')
    f.close()

def next_problem():
    global id
    global flag
    global onlyerr
    global index
    global multi
    multi=[0,0,0,0]
    flag=False
    if onlyerr==False:
        id+=1
    else:
        index+=1
        id=err[index]
    state['text']="第"+str(id)+'题,'
    btn1['background']='#EEEEEE'
    btn2['background']='#EEEEEE'
    btn3['background']='#EEEEEE'
    btn4['background']='#EEEEEE'
    if problems['题型'][id]=='单选题':
        state['text']+='单选题'
        show_single()
        btn1['text']=problems["A"][id]
        btn2['text']=problems["B"][id]
        btn3['text']=problems["C"][id]
        btn4['text']=problems["D"][id]
        btn1['command']=chk1
        btn2['command']=chk2
        btn3['command']=chk3
        btn4['command']=chk4
        btn5['command']=show_ans
        btn5['text']='查看答案'
    if problems['题型'][id]=='填空题':
        ent.delete('1.0','end')
        state['text']+='填空题'
        show_text()
        btn5['command']=chk5
        btn5['text']='确定'
        btn6['command']=show_ans
        btn6['text']='查看答案'
    if problems['题型'][id]=='判断题':
        state['text']+='判断题'
        show_tf()
        btn1['command']=chk1
        btn2['command']=chk2
        btn3['command']=chk3
        btn4['command']=chk4
        btn1['text']=problems["A"][id]
        btn2['text']=problems["B"][id]
        btn5['command']=show_ans
        btn5['text']='查看答案'
    if problems['题型'][id]=='多选题':
        state['text']+='多选题'
        btn1['command']=multi1
        btn2['command']=multi2
        btn3['command']=multi3
        btn4['command']=multi4
        btn1['text']=problems["A"][id]
        btn2['text']=problems["B"][id]
        btn3['text']=problems["C"][id]
        btn4['text']=problems["D"][id]
        btn5['command']=check_muti
        btn5['text']='确定'
        btn6['command']=show_ans
        btn6['text']='查看答案'
        show_multi()
    lab["text"]=problems["题干"][id]

def show_single():
    ent.pack_forget()
    btn5.pack_forget()
    btn6.pack_forget()
    btn1["height"]=3
    btn2["height"]=3
    btn1.pack(fill='both')
    btn2.pack(fill='both')
    btn3.pack(fill='both')
    btn4.pack(fill='both')
    btn5.pack(side='left',fill='both')

def show_text():
    global btn1,btn2,btn3,btn4,btn5
    btn1.pack_forget()
    btn2.pack_forget()
    btn3.pack_forget()
    btn4.pack_forget()
    btn5.pack_forget()
    btn6.pack_forget()
    ent.pack(fill='both')
    btn5.pack(side='left',fill='both')
    btn6.pack(side='left',fill='both')

def show_tf():
    ent.pack_forget()
    btn1.pack_forget()
    btn2.pack_forget()
    btn3.pack_forget()
    btn4.pack_forget()
    btn5.pack_forget()
    btn6.pack_forget()
    btn1["height"]=6
    btn2["height"]=6
    btn1.pack(fill='both')
    btn2.pack(fill='both')
    btn5.pack(side='left',fill='both')

def show_multi():
    ent.pack_forget()
    btn5.pack_forget()
    btn6.pack_forget()
    btn1["height"]=3
    btn2["height"]=3
    btn1.pack(fill='both')
    btn2.pack(fill='both')
    btn3.pack(fill='both')
    btn4.pack(fill='both')
    btn5.pack(side='left',fill='both')
    btn6.pack(side='left',fill='both')

def chk1():
    global flag
    if problems['正确答案'][id]=='A':
        next_problem()
    elif flag==False:
        flag=True
        write_err()

def chk2():
    global flag
    if problems['正确答案'][id]=='B':
        next_problem()
    elif flag==False:
        flag=True
        write_err()

def chk3():
    global flag
    if problems['正确答案'][id]=='C':
        next_problem()

def chk4():
    global flag
    if problems['正确答案'][id]=='D':
        next_problem()
    elif flag==False:
        flag=True
        write_err()

def chk5():
    global flag
    if problems['A'][id]+'\n'==ent.get('0.0',END):
        next_problem()
    elif flag==False:
        flag=True
        write_err()

def multi1():
    if multi[0]:
        btn1['background']='#EEEEEE'
        multi[0]=0
    else:
        btn1['background']='deepskyblue'
        multi[0]=1

def multi2():
    if multi[1]:
        btn2['background']='#EEEEEE'
        multi[1]=0
    else:
        btn2['background']='deepskyblue'
        multi[1]=1

def multi3():
    if multi[2]:
        btn3['background']='#EEEEEE'
        multi[2]=0
    else:
        btn3['background']='deepskyblue'
        multi[2]=1

def multi4():
    if multi[3]:
        btn4['background']='#EEEEEE'
        multi[3]=0
    else:
        btn4['background']='deepskyblue'
        multi[3]=1

def check_muti():
    global flag
    ans=''
    if multi[0]:
        ans+='A'
    if multi[1]:
        ans+='B'
    if multi[2]:
        ans+='C'
    if multi[3]:
        ans+='D'
    if ans==problems['正确答案'][id]:
        next_problem()
    elif flag==False:
        flag=True
        write_err()

def show_ans():
    global flag
    if (problems['题型'][id]=='单选题') or (problems['题型'][id]=='判断题') or (problems['题型'][id]=='多选题'):
        if 'A' in problems['正确答案'][id]:
            btn1['background']='aquamarine'
        if 'B' in problems['正确答案'][id]:
            btn2['background']='aquamarine'
        if 'C' in problems['正确答案'][id]:
            btn3['background']='aquamarine'
        if 'D' in problems['正确答案'][id]:
            btn4['background']='aquamarine'
    if problems['题型'][id]=='填空题':
        ent.delete('1.0','end')
        ent.insert('end',problems['A'][id])
    btn5['command']=next_problem
    btn5['text']='下一题'
    if flag==False:
        flag=True
        write_err()

def jump():
    global id
    id=int(go.get())-1
    next_problem()

def use_err():
    global onlyerr
    onlyerr=True
    next_problem()

f = open("err.txt")
for line in f:
    err.append(int(line))
f.close()

err=list(set(err))

f=open('err.txt','w')
f.close()


root = Tk()                     # 创建窗口对象的背景色
root.geometry("2000x1200")
#调用api设置成由应用程序缩放
ctypes.windll.shcore.SetProcessDpiAwareness(1)
#调用api获得当前的缩放因子
ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
#设置缩放因子
root.tk.call('tk', 'scaling', ScaleFactor/75)
fr1=Frame(root,height=1)
go=Entry(fr1,width='1')
gobtn=Button(fr1,text='跳转',width='1',command=jump)
errbtn=Button(fr1,text='错题',width='1',command=use_err)
lab=Label(root,text="Hello RUNOOB!",height=10,wraplength = 1800,font=("等线", 18))
state=Label(fr1,height=1,width=100)
ent=Text(root,height=15,font=("等线", 14))
btn1=Button(root,height=3,wraplength = 1800,command=chk1,font=("等线", 14))
btn2=Button(root,height=3,wraplength = 1800,command=chk2,font=("等线", 14))
btn3=Button(root,height=3,wraplength = 1800,command=chk3,font=("等线", 14))
btn4=Button(root,height=3,wraplength = 1800,command=chk4,font=("等线", 14))
btn5=Button(root,height=3,width=10,text="下一题",command=next_problem)
btn6=Button(root,height=3,width=10,text="查看答案",command=show_ans)
problems=pd.read_excel("./problems.xlsx")
fr1.pack(fill='x',side='top',expand=TRUE)
state.pack(fill='x',side='left',expand=TRUE)
go.pack(fill='x',side='left',expand=TRUE)
gobtn.pack(fill='x',side='left',expand=TRUE)
errbtn.pack(fill='x',side='left',expand=TRUE)
lab.pack(fill='both',side=tkinter.TOP,expand=TRUE)
next_problem()
root.mainloop()                 # 进入消息循环
