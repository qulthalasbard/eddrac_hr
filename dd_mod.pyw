# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf8')
from Tkinter import *
import xlwt
import Tkinter as tk
import openpyxl
import MySQLdb as mydb
import datetime
import subprocess
import tkFont
import ttk
import gc



#表格创建
top=tk.Tk()
top.title("岗位\\宿舍调整")
#字体
ft1 = tkFont.Font(family='msyhbd', size=20, weight=tkFont.BOLD)
ft2 = tkFont.Font(family='msyh', size=15)
ft3 = tkFont.Font(family='msyh', size=12)
ft4 = tkFont.Font(family='msyh', size=10)
ft5 = tkFont.Font(family='msyhbd', size=6)
ft6 = tkFont.Font(family='msyhbd', size=20,weight="bold")
style = ttk.Style()

style = ttk.Style()
style.map("TCombobox",
    fieldbackground=[('readonly', '#aaffee')],
    selectbackground=[('!focus', '#aaffee')]
    )


#数据库变量

db_user=""
db_passwd=""
db_host="cdb-pdylikw0.bj.tencentcdb.com"
db_base="eddrac_employees"
db_port=10005



#列出框架
tk.Label(top,relief=GROOVE,borderwidth=1,bg="#eeeeee",text="岗位宿舍调动登记",font=ft6).grid(row=0,column=0,sticky=N+S+W+E)
fm1=Frame(top)
fm1.grid(row=1,column=0,sticky=N+S+W+E)
fm2=Frame(fm1,relief=GROOVE,borderwidth=1)
fm2.grid(row=0,column=0,sticky=N+W)

#日期检测函数
def date_chk(str_x):
 if len(str_x)==0:
  return 0
 if len(str_x)==10:
  if str_x[:4].isdigit() and str_x[5:7].isdigit() and str_x[-2:].isdigit():
    if int(str_x[:4])>2015 and int(str_x[:4]) < 2020 and int(str_x[5:7])<13 and int(str_x[5:7])>0 and  int(str_x[-2:])<32 and int(str_x[-2:])>0 and str_x[4:5]==str_x[-3:-2]=="-":
     return 1
 return -1 

#与当前日期差值
def date_diff(x):
  x=str(x)+" 00:00:00"
  y=str(datetime.datetime.now().date())+" 00:00:00"
  x=datetime.datetime.strptime(x,'%Y-%m-%d %H:%M:%S')
  y=datetime.datetime.strptime(y,'%Y-%m-%d %H:%M:%S')
  z=(x-y)
  if len(str(z))<15:
   return 0
  return int(str(z)[:-13])

global d_click
global do_click
d_click=0
do_click=0
#岗位调动函数确认函数
def import_dd():
    global in_sql,q_sql,text_done,d_click,fm3,up_sql
    if d_click%2==0:
        
        dd_date=en_date.get()
        if date_chk(dd_date)!=1 or date_diff(dd_date)<-3 or date_diff(dd_date)>4 or len(de_com.get())==0 or len(pos_com.get())==0:
            en_date.delete(0,END)
            en_date.insert(END,"日期或职位部门有误")
            return
        else:
            d_sql="select id_d from department where 部门=\""+de_com.get()+"\""
            p_sql="select id_p from position where 职位=\""+pos_com.get()+"\""
            cur.execute(d_sql)
            d_sql=cur.fetchall()
            
          
            dep=str(d_sql[0][0])
            cur.execute(p_sql)
            p_sql=cur.fetchall()
         
            pos=str(p_sql[0][0])

            
            in_sql="insert into dd_mod(工号,调整时间,之后部门,之后职位,last) values(\""+str(result[0][5])+"\",\""+str(dd_date)+"\",\""+dep+"\",\""+pos+"\",\""+str(result[0][4])+"\")"
            up_sql="update all_employees set 部门="+dep+",职位="+pos+"where 工号="+str(result[0][5])
            q_sql="update dd_mod set dd_mod.next=(select dd_mod.id from(select id from dd_mod where last="+str(result[0][4])+")dd_mod) where id="+str(result[0][4])
            text_done=result[0][0]+"于"+str(dd_date)+"日从"+result[0][1]+result[0][2]+"岗位\n调动到"+de_com.get()+pos_com.get()+"岗位"
            fm3.destroy()
            fm3=Frame(fm1,relief=GROOVE,borderwidth=1)
            fm3.grid(row=0,column=1,sticky=N+W+E+S)
            d_click=1
            tk.Button(fm3,text=text_done+",确定?",relief=RAISED,command=import_dd,font=ft2).grid(row=0,column=0)
    else:
        cur.execute(in_sql)
        cur.execute(up_sql)
        db.commit()
        cur.execute(q_sql)
        db.commit()
        db.close()
        d_click=0
        tk.Label(fm3,text=text_done+"登记完成！",relief=RAISED,font=ft2).grid(row=0,column=0)
        return 
        
  


#职位调整页面打开
def posi_open():
    global p_name,fm3,de_com,pos_com,en_date,result,butt,cur,db
    try:
        fm3.destroy()
    except NameError,e:
        pass
    fm3=Frame(fm1,relief=GROOVE,borderwidth=1)
    fm3.grid(row=0,column=1,sticky=N+W+E+S)
    
#在函数内部打开数据库，提高页面打开速度
    
    db=mydb.connect(host=db_host,user=db_user,passwd=db_passwd,port=db_port,db=db_base,charset='utf8')
    cur=db.cursor()
    pp_name=p_name.get()
#   update test set test.next=(select test.id from(select id from test where last=2053)text) where test.id=2053
    sql="select all_employees.姓名,department.部门,position.职位,datediff(curdate(),dd_mod.调整时间),dd_mod.id,all_employees.工号 from all_employees,position,department,dd_mod where department.id_d=all_employees.部门 and position.id_p=all_employees.职位 and all_employees.工号=dd_mod.工号 and dd_mod.next IS NULL and all_employees.姓名 like \"%"+pp_name+"%\""
    cur.execute(sql)
    result=cur.fetchall()
    if len(result)!=1:
        tk.Label(fm3,text="姓名输入有误!",font=ft6).grid(row=0,column=0,sticky=N+W)
        return
    else:
        tk.Label(fm3,font=ft5,height=0).grid(row=0,column=0)  
        n=0
        tk.Label(fm3,font=ft5).grid(row=0,column=0)                                  
        for x in ["姓名","当前部门","当前职位","当前职位时间(单位:天)"]:
            tk.Label(fm3,text=x,relief=RAISED,borderwidth=2,font=ft2).grid(row=1,column=n,sticky=N+S+W+E)
            tk.Label(fm3,text=result[0][n],relief=GROOVE,borderwidth=1,font=ft2,bg="#aaeeaa").grid(row=2,column=n,sticky=N+S+W+E)
            n=n+1
        fm4=Frame(fm3)
        fm4.grid(row=3,column=0,columnspan=4,sticky=N+S+W+E)
        dep_sql="select 部门 from department order by id_d"
        cur.execute(dep_sql)
        dep_result=cur.fetchall()
        pos_sql="select 职位 from position order by id_p"
        cur.execute(pos_sql)
        pos_result=cur.fetchall()
#       update test set test.next=(select test.id from(select id from test where last=2053)text) where test.id=2053
        tk.Label(fm4,font=ft5,height=4).grid(row=0,column=0,sticky=N+S+W+E)
        tk.Label(fm4,font=ft2,text="调动后部门",relief=RAISED).grid(row=1,column=0)
        de_value=tk.StringVar()
        de_com=ttk.Combobox(fm4,width=8,textvariable=de_value,state="readonly",style="TCombobox")
        de_com["values"]=dep_result
        de_com.grid(row=1,column=1)
        tk.Label(fm4,font=ft2,width=10,relief=RAISED).grid(row=1,column=2,sticky=E+W)
        tk.Label(fm4,font=ft2,text="调动后职位",relief=RAISED).grid(row=1,column=3)
        pos_value=tk.StringVar()
        pos_com=ttk.Combobox(fm4,width=8,textvariable=pos_value,state="readonly",style="TCombobox")
        pos_com["values"]=pos_result
        pos_com.grid(row=1,column=4)
        
        tk.Label(fm4,font=ft5,height=4).grid(row=2,column=0,sticky=N+S+W+E)
        tk.Label(fm4,font=ft3,height=1,text="调动时间",relief=RAISED).grid(row=3,column=1,sticky=N+S+W+E)
        en_date=tk.Entry(fm4,bg="#aaffaa")
        en_date.grid(row=3,column=2)
        tk.Label(fm4,font=ft5,height=4).grid(row=4,column=0,sticky=N+S+W+E)
        butt=tk.Button(fm4,text="确认调动",relief=RAISED,font=ft6,command=import_dd).grid(row=5,column=1,columnspan=2,sticky=N+S+W+E)
        tk.Label(fm4,font=ft4).grid(row=6,column=1)
        
        fm5=Frame(fm3)
        fm5.grid(row=4,column=0,columnspan=4,sticky=N+S+W+E)
        scan_sql="select dd_mod.调整时间,department.部门,position.职位,dd_mod.调整原因 from dd_mod inner join department on department.id_d=dd_mod.之后部门 inner join position on position.id_p=dd_mod.之后职位 where 工号="+str(result[0][5])+" order by 调整时间 desc"
        cur.execute(scan_sql)
        scan_result=cur.fetchall()
        tk.Label(fm5,text=result[0][0]+"岗位调整记录表",relief=RAISED,font=ft2).grid(row=0,column=0,columnspan=4,sticky=N+W+E+S)
        tk.Label(fm5,font=ft4).grid(row=1,column=0)
        tn=0
        for xx in ["调动时间","调往部门","调往职位","调动原因"]:
            tk.Label(fm5,text=xx,relief=RAISED,font=ft2).grid(row=2,column=tn,sticky=N+S+W+E)
            tn=tn+1
            
        
        ln=0
        while ln < len(scan_result):
            lln=0
            while lln< 4:

                tk.Label(fm5,text=scan_result[ln][lln],font=ft3,relief=GROOVE,bg="#aaffaa").grid(row=ln+3,column=lln,sticky=N+S+W+E)
                lln=lln+1
            ln=ln+1
        tk.Label(fm5,font=ft2).grid(row=ln+3,column=0,sticky=N+S+W+E)

#下面是宿舍调整内容
#宿舍调整确认
def do_confirm():
    global do_click,do_date,do_result,do_com,do_in_sql,do_up_sql,fm3,text_do
    if do_click%2==0:
        doo_date=do_date.get()
        if date_chk(doo_date)!=1 or date_diff(doo_date)<-3 or date_diff(doo_date)>4 or len(do_com.get())==0:
            do_date.delete(0,END)
            do_date.insert(END,"日期有误或宿舍没选")
            return
        else:
            if type(do_result[0][3]) is NoneType:
                do_in_sql="insert into do_mod(工号,调整时间,调整原因,之后宿舍) values(\""+str(do_result[0][1])+"\",\""+doo_date+"\",+\"首次入住\",\""+str(do_com.get())+"\")"
                text_do=do_result[0][4]+do_result[0][5]+do_result[0][0]+"\n将新入住"+str(do_com.get())+"宿舍"
            else:
                text_do=do_result[0][4]+do_result[0][5]+do_result[0][0]+"\n将从"+str(do_result[0][3])+"宿舍调换到"+str(do_com.get())+"宿舍"
                do_in_sql="insert into do_mod(工号,调整时间,调整原因,之后宿舍,之前宿舍) values(\""+str(do_result[0][1])+"\",\""+doo_date+"\",+\"首次入住\",\""+str(do_com.get())+"\",\""+str(do_result[0][3])+"\")"                                                                
            do_up_sql="update all_employees set 宿舍="+str(do_com.get())+" where 工号="+str(do_result[0][1])
            fm3.destroy()
            fm3=Frame(fm1,relief=GROOVE,borderwidth=1)
            fm3.grid(row=0,column=1,sticky=N+W+E+S)
            do_click=1
            tk.Button(fm3,text=text_do+",确定?",relief=RAISED,command=do_confirm,font=ft2).grid(row=0,column=0)
    else:
        cur.execute(do_in_sql)
        cur.execute(do_up_sql)
        db.commit()
       
        
        db.close()
        do_click=0
        tk.Label(fm3,text=text_do+"登记完成！",relief=RAISED,font=ft2).grid(row=0,column=0)
        return 
                                                                                
            
    
        

#宿舍调整页面展开
        
def do_open():
    global fm3,do_name,do_date,do_com,do_result,cur,db
    try:
        fm3.destroy()
    except NameError:
        pass
    fm3=Frame(fm1,relief=GROOVE,borderwidth=1)
    fm3.grid(row=0,column=1,sticky=N+W+E+S)
    db=mydb.connect(host=db_host,user=db_user,passwd=db_passwd,port=db_port,db=db_base,charset='utf8')
    cur=db.cursor()
    dd_name=do_name.get()
    sql="select 姓名,工号,性别,宿舍,department.部门,position.职位 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.姓名 like \"%"+dd_name+"%\""
    cur.execute(sql)
    do_result=cur.fetchall()
    if len(do_result)!=1:
        tk.Label(fm3,text="姓名输入有误!",font=ft6).grid(row=0,column=0,sticky=N+W)
        return
    else:
        do_list=["302","303","304","305","306","307","308","309","310","311","312","313","314","315","316","317","401","402","403","404","405","406","408","409","410","411","412","413","414","415","416","417","418"]
        dr=[]
        for x in do_list:
            sql="select 性别 from all_employees where 宿舍=\""+str(x)+"\""
            
            cur.execute(sql)
            do_chk=cur.fetchall()
            if len(do_chk)>1 and do_chk[0][0]!=do_result[0][2]:
                pass
            elif len(do_chk)==2 and do_chk[0][0]!=do_chk[1][0]:
                pass
            elif len(do_chk)>3:
                pass
            else:
                dr.append(x)
                pass

       
        if type(do_result[0][3]) is NoneType:
            ltext=do_result[0][4]+do_result[0][5]+do_result[0][0]+"之前没有宿舍。\n这次要调换到:"
        else:
            if do_result[0][3] in dr:
                dr.remove(do_result[0][3])
            ltext=do_result[0][4]+do_result[0][5]+do_result[0][0]+"之前住在"+str(do_result[0][3])+"宿舍。\n这次要调换到:"
        tk.Label(fm3,text=ltext,font=ft3).grid(row=0,column=0,columnspan=2)
        tk.Label(fm3,font=ft3).grid(row=1,column=0)
        tk.Label(fm3,font=ft3,text="目前可调换宿舍有:").grid(row=2,column=0,sticky=W+S+N)
        do_value=tk.StringVar()
        do_com=ttk.Combobox(fm3,width=8,textvariable=do_value,state="readonly",style="TCombobox")
        do_com["values"]=dr
        do_com.grid(row=2,column=1)
        tk.Label(fm3,font=ft3).grid(row=3,column=0)
        tk.Label(fm3,font=ft3,text="调换日期").grid(row=4,column=0)
        do_date=Entry(fm3,bg="#aaffaa",width=12)
        do_date.grid(row=4,column=1)
        tk.Label(fm3,font=ft3).grid(row=5,column=0)
        do_but=Button(fm3,text="确认",command=do_confirm,font=ft3)
        do_but.grid(row=6,column=0,columnspan=2,sticky=N+S+W+E)




#输入要调整薪资人员姓名

tk.Label(fm2,font=ft5,height=2,width=1).grid(row=0,column=0,sticky=N+W+S+E)
tk.Label(fm2,font=ft6,height=2,width=1,text="岗位调动").grid(row=1,column=1,columnspan=2,sticky=N+S+W+E)

tk.Label(fm2,font=ft5,height=2,width=1).grid(row=2,column=0,sticky=N+S+W+E)
tk.Label(fm2,text="要调整职位\n的员工姓名",font=ft3,height=2).grid(row=2,column=1,sticky=N+S+W+E)
p_name=Entry(fm2,font=ft3)
p_name.grid(row=2,column=2)
tk.Label(fm2,font=ft1,width=1).grid(row=4,column=3)
tk.Button(fm2,text="确认",command=posi_open,borderwidth=1,bg="#208040",font=ft3).grid(row=5,column=1,columnspan=2,sticky=N+S+W+E)
tk.Label(fm2,font=ft3).grid(row=6,column=0,columnspan=2)
tk.Label(fm2,font=ft6,text="宿舍调换").grid(row=7,column=1,columnspan=2)
tk.Label(fm2,font=ft5,width=2).grid(row=8,column=0,columnspan=2)
tk.Label(fm2,text="要调整宿舍\n的员工姓名",font=ft3,height=2).grid(row=9,column=1,sticky=N+S+W+E)

do_name=Entry(fm2,font=ft3)
do_name.grid(row=9,column=2)
tk.Label(fm2,font=ft6).grid(row=10,column=0,columnspan=2)
tk.Button(fm2,text="确认",command=do_open,borderwidth=1,bg="#208040",font=ft3).grid(row=11,column=1,columnspan=2,sticky=N+S+W+E)
tk.Label(fm2,font=ft6).grid(row=12,column=0,columnspan=2)



















#框架





#列出按钮函数
 

top.mainloop()
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
