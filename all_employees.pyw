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
import gc



#表格创建
top=tk.Tk()
top.title("人员列表")
#字体
ft1 = tkFont.Font(family='msyhbd', size=20, weight=tkFont.BOLD)
ft2 = tkFont.Font(family='msyh', size=15)
ft3 = tkFont.Font(family='msyh', size=12)
ft4 = tkFont.Font(family='msyh', size=10)
ft5 = tkFont.Font(family='msyhbd', size=6)


#结果导出类
class name_label:
    def __init__(self,name,gh,base,base1):
        self.name=name
        self.gh=gh
        self.base=base
        
        self.base1=base1

    def c_label(self):
        la=tk.Label(self.base,text=self.name,relief=GROOVE,borderwidth=1,font=ft3)
        
        return la
    
    def l_salary(self,x):
        sql="select all_employees.姓名,salary_mod.之后基薪,salary_mod.之后岗位薪资,salary_mod.之后绩效,salary_mod.之后全勤奖,salary_mod.之后话补,salary_mod.之后补贴,salary_mod.之后车补,all_employees.电话,all_employees.入职时间,all_employees.身份证号码,all_employees.dkp from salary_mod,all_employees WHERE all_employees.工号=salary_mod.工号 and all_employees.工号=\""+str(self.gh)+"\""
        cur.execute(sql)
        result=cur.fetchall()
        self.base1.destroy()
        self.base1=Frame(top)
        self.base1.grid(row=2,column=1,sticky=W)
        if len(result)==0:
            result=list(result)
            result=[["编外人员","编外人员","编外人员","编外人员","编外人员","编外人员","编外人员","编外人员","编外人员","编外人员","编外人员"]]
            
        n=0
        for xx in ["姓名","基本工资","岗位薪资","绩效考核","全勤奖","话补","补贴","车补","电话"]:
            Label(self.base1,text=xx,font=ft3,relief=GROOVE,borderwidth=1).grid(row=0,column=n,sticky=W+E+N+S)
            Label(self.base1,text=result[0][n],relief=GROOVE,font=ft3,borderwidth=1,width=12).grid(row=1,column=n,sticky=W+E+N+S)
            n=n+1
        Label(self.base1,text="入职时间:",relief=GROOVE,font=ft3,borderwidth=1,width=12).grid(row=2,column=0,sticky=W+E+N+S)
	Label(self.base1,text=str(result[0][n]),relief=GROOVE,font=ft3,borderwidth=1,width=12).grid(row=2,column=1,sticky=W+E+N+S)
        Label(self.base1,text="身份证号码:",relief=GROOVE,font=ft3,borderwidth=1,width=12).grid(row=2,column=2,sticky=W+E+N+S)
	Label(self.base1,text=str(result[0][n+1]),relief=GROOVE,font=ft3,borderwidth=1,width=12).grid(row=2,column=3,sticky=W+E+N+S,columnspan=2)
	Label(self.base1,text="工作表现:",relief=GROOVE,font=ft3,borderwidth=1,width=12).grid(row=2,column=5,sticky=W+E+N+S)
        Label(self.base1,text=str(result[0][n+2])+"分",relief=GROOVE,font=ft3,borderwidth=1,width=12).grid(row=2,column=6,sticky=W+E+N+S)
            
    	Label(self.base1,font=ft1).grid(row=3,column=1,sticky=W+E+N+S)
  

#数据库变量

db_user="eddrac_hr"
db_passwd=""
db_host=""
db_base="eddrac_employees"
db_port=10005

#数据库链接

db=mydb.connect(host=db_host,user=db_user,passwd=db_passwd,port=db_port,db=db_base,charset='utf8')
cur=db.cursor()

#查询数据



#框架



fm2=Frame(top,relief=GROOVE,borderwidth=1)
fm2.grid(row=1,column=0,sticky=N+W+E+S,rowspan=2)

#列出按钮函数
         
def all_employees():
    global fm3,fm4
    col=28
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.当前=\"在职\" ORDER BY 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=4*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=4*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*x+3,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=4*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+1,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)

        
def product():
    global fm3,fm4
    col=20
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.当前=\"在职\" and (department.部序 like \"06%\" or department.部序 like \"07%\") ORDER BY 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=4*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=4*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*x+3,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=4*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+1,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)

    
def office():
    global fm3,fm4
    col=10
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.当前=\"在职\" and department.id_d in (1,2,3,4,5,6,9,10,11) ORDER BY 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=4*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=4*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*x+3,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=4*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+1,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)


def month_new():
    global fm3,fm4
    col=5
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.当前=\"在职\" and DATE_FORMAT(入职时间,\'%Y%m\')=DATE_FORMAT(CURDATE( ),\'%Y%m\')ORDER BY 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=4*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=4*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*x+3,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=4*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+1,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)

        

def month_aug():
    global fm3,fm4
    col=5
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.当前=\"在职\" and DATE_FORMAT(转正日期,\'%Y%m\')=DATE_FORMAT(CURDATE( ),\'%Y%m\')ORDER BY 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=4*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=4*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*x+3,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=4*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+1,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)


def month_period():
    global fm3,fm4
    col=6
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.当前=\"在职\" and 转正日期 IS NULL ORDER BY 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=4*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=4*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*x+3,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=4*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+1,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=4*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)


def month_quit():
    global fm3,fm4
    col=6
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位,emp_quit.确认时间 from all_employees,department,position,emp_quit where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.工号=emp_quit.工号 and all_employees.当前=\"离职\" and DATE_FORMAT(emp_quit.确认时间, \"%Y%m\")=DATE_FORMAT(curdate(),\"%Y%m\") order by 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=5*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=5*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=5*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=5*x+3,sticky=N+W+E+S+E)
	Label(fm3,text="离职日期",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=5*x+4,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=5*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=5*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=5*int((n-1)/col)+1,sticky=N+S+W+E)
        Label(fm3,text=s_r[n-1][4],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=5*int((n-1)/col)+4,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=5*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=5*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=5*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)

#上月离职员工
def last_month_quit():
    global fm3,fm4
    col=6
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位,emp_quit.确认时间 from all_employees,department,position,emp_quit where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.工号=emp_quit.工号 and all_employees.当前=\"离职\" and DATE_FORMAT(emp_quit.确认时间, \"%Y%m\")=DATE_FORMAT(DATE_SUB(curdate(), INTERVAL 1 MONTH),\"%Y%m\") order by 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=5*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=5*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=5*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="职位",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=5*x+3,sticky=N+W+E+S+E)
	Label(fm3,text="离职日期",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=5*x+4,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=5*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=5*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=5*int((n-1)/col)+1,sticky=N+S+W+E)
        Label(fm3,text=s_r[n-1][4],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=5*int((n-1)/col)+4,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=5*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=5*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=5*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)




#准备离职员工
def pre_quit():
    global fm3,fm4
    col=6
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="SELECT department.部门,all_employees.姓名,all_employees.工号,position.职位,emp_quit.申请时间,emp_quit.到期时间 from all_employees,department,position,emp_quit where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.工号=emp_quit.工号 and emp_quit.确认时间 IS NULL and emp_quit.申请时间 IS NOT NULL order by 部序,职序,姓名"
    cur.execute(sql)
    s_r=cur.fetchall()

#计算需要几列，并列出题头

    n=int(len(s_r)/col) + 1
    for x in range(n):
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=6,borderwidth=1).grid(row=0,column=6*x,sticky=N+W+E+S+E)
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=6*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=6*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="入职日期",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=6*x+3,sticky=N+W+E+S+E)
	Label(fm3,text="提出申请日期",font=ft3,relief=RAISED,width=14,borderwidth=1).grid(row=0,column=6*x+4,sticky=N+W+E+S+E)
        Label(fm3,text="到期日期",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=6*x+5,sticky=N+W+E+S+E)
#列出人员
    n=1
    m=0
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
	
    while n<len(s_r):
        lua=name_label(s_r[n-1][1],s_r[n-1][2],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=6*int((n-1)/col)+2,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
        Label(fm3,text=s_r[n-1][3],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=6*int((n-1)/col)+3,sticky=N+S+W+E)
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=6*int((n-1)/col)+1,sticky=N+S+W+E)
        Label(fm3,text=s_r[n-1][4],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=6*int((n-1)/col)+4,sticky=N+S+W+E)
        Label(fm3,text=s_r[n-1][5],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=6*int((n-1)/col)+5,sticky=N+S+W+E)
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=6*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=col-m,column=6*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0

        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1).grid(row=n%col-m,column=6*int((n-1)/col),rowspan=m+1,sticky=N+S+W+E)
            m=0
        n=n+1
	Label(fm3,font=ft4).grid(row=col+1,column=0)


#住宿员工
def do_emp():
    global fm3,fm4
    col=26
    gc.collect()
    try:
	fm3.destroy()
    except NameError,e:
        pass
        
    try:
	fm4.destroy()
    except NameError,e:
        pass
    

    fm3=Frame(top,relief=GROOVE,borderwidth=1)
    fm3.grid(row=1,column=1,sticky=N+W+E+S)

    fm4=Frame(top)
    fm4.grid(row=2,column=1)
#首先查询数据库
    
    sql="select all_employees.宿舍,department.部门,all_employees.姓名,all_employees.工号,all_employees.性别 from all_employees,department where all_employees.部门=department.id_d and (all_employees.当前=\"在职\" or all_employees.当前=\"外包\") and all_employees.宿舍 is not Null order by 宿舍"
    cur.execute(sql)
    s_r=cur.fetchall()
#列出题头
    n=int(len(s_r)/col)
    x=0
    
    while x<n+1:
        Label(fm3,text="序号",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x,sticky=N+W+E+S+E)
        Label(fm3,text="宿舍",font=ft3,relief=RAISED,width=5,borderwidth=1).grid(row=0,column=4*x+1,sticky=N+W+E+S+E)
        Label(fm3,text="部门",font=ft3,relief=RAISED,width=7,borderwidth=1).grid(row=0,column=4*x+2,sticky=N+W+E+S+E)
        Label(fm3,text="姓名",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*x+3,sticky=N+W+E+S+E)
        x=x+1
    
    Label(fm3,text="空余宿舍",font=ft3,relief=RAISED,width=10,borderwidth=1).grid(row=0,column=4*(x-1)+4,sticky=N+W+E+S+E)
	
    
#列出人员
    
    
    s_r=list(s_r)
    s_r.append([1,2,3,4,5])
    for n in range(len(s_r)):
        s_r[n]=list(s_r[n])
    
    
    n=0          
    while n < len(s_r)-1:
        if s_r[n][0]==s_r[n+1][0] and s_r[n][4]!=s_r[n+1][4]:
#            s_r[n][0]=s_r[n][0]+"\n(夫妻房)"
#           s_r[n+1][0]=s_r[n][0]
            s_r[n][4]=0
            s_r[n+1][4]=0
#        else:
#            if s_r[n][4]!=0:
#                s_r[n][0]=s_r[n][0]+"\n("+s_r[n][4]+"宿舍)"
        
        if s_r[n][4]=="男":
            s_r[n][4]=1
        elif s_r[n][4]=="女":
            s_r[n][4]=2
            
        
        n=n+1
        
               
    n=1
    m=0
    d_list=["302","303","304","305","306","307","308","309","310","311","312","313","314","315","316","317","318","401","402","403","404","405","406","407","408","409","410","411","412","413","414","415","416","417","418","419"]
    bgc=["#ff700a","#5030ff","#ff2074"]
    while n<len(s_r):

        try:
            d_list.remove(s_r[n-1][0])
        except ValueError,e:
            pass
#序号拍第一        
        Label(fm3,text=str(n),relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col),sticky=N+S+W+E)
#宿舍拍第二
        if s_r[n-1][0]==s_r[n][0] and n%col!=0:
            m=m+1
          
        elif s_r[n-1][0]==s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1,bg=bgc[s_r[n-1][4]]).grid(row=col-m,column=4*int((n-1)/col)+1,rowspan=m+1,sticky=N+S+W+E)
            m=0
            
        elif s_r[n-1][0]!=s_r[n][0] and n%col==0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1,bg=bgc[s_r[n-1][4]]).grid(row=col-m,column=4*int((n-1)/col)+1,rowspan=m+1,sticky=N+S+W+E)
            m=0
            
        elif s_r[n-1][0]!=s_r[n][0] and n%col!=0:
            Label(fm3,text=s_r[n-1][0],relief=GROOVE,borderwidth=1,bg=bgc[s_r[n-1][4]]).grid(row=n%col-m,column=4*int((n-1)/col)+1,rowspan=m+1,sticky=N+S+W+E)
            m=0
            
        Label(fm3,text=s_r[n-1][1],relief=GROOVE,borderwidth=1,font=ft3).grid(row=(n-1)%col+1,column=4*int((n-1)/col)+2,sticky=N+S+W+E)
        
#姓名排第四
        lua=name_label(s_r[n-1][2],s_r[n-1][3],fm3,fm4)
        n_label=lua.c_label()
        n_label.grid(row=(n-1)%col+1,column=4*int((n-1)/col)+3,sticky=N+S+W+E)
        n_label.bind("<Enter>",lua.l_salary)
#性别排第五
        
        
       
        
        n=n+1
    c=4*int((n-2)/col)+2
    r=(n-2)%col+1
    
    for emp_do in d_list:
        Label(fm3,text=emp_do,relief=GROOVE,borderwidth=1,bg="#00dede").grid(row=r%col+1,column=int(r/col)+c+1,sticky=N+W+E+S+E)
        r=r+1

#岗位调动页面
def do_mod():
    cmd="c:\python27\python.exe D:\\Eddrac_hr\\dd_mod.pyw"
 
    subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE)





#列出标题和按钮

Label(top,text="伊德莱克木业在职人员一览表",font=ft1,borderwidth=2,relief=GROOVE,bg="#eeeeee").grid(row=0,column=0,columnspan=2,sticky=N+S+W+E)


#打开首先列出全部人员

all_employees()




#空出第一行
Label(fm2,font=ft3,width=3).grid(row=0,column=0,)
#按钮1
a_e=Button(fm2,text="全部在职人员",command=all_employees,font=ft3,borderwidth=1,relief=RAISED,width=12)
a_e.grid(row=1,column=1)
#空出第三行，同时留出空间
Label(fm2,font=ft3,width=3).grid(row=2,column=2)

#按钮2
Label(fm2,font=ft5,width=3).grid(row=3,column=0)
prod=Button(fm2,text="车间人员",command=product,font=ft3,borderwidth=1,relief=RAISED,width=12)
prod.grid(row=4,column=1)
Label(fm2,font=ft5,width=3).grid(row=5,column=2)

#按钮3
Label(fm2,font=ft5,width=3).grid(row=6,column=0)
off=Button(fm2,text="业务及后勤",command=office,font=ft3,borderwidth=1,relief=RAISED,width=12)
off.grid(row=7,column=1)
Label(fm2,font=ft5,width=3).grid(row=8,column=2)

#按钮4
Label(fm2,font=ft5,width=3).grid(row=9,column=0)
m_n=Button(fm2,text="本月新入职",command=month_new,font=ft3,borderwidth=1,relief=RAISED,width=12)
m_n.grid(row=10,column=1)
Label(fm2,font=ft5,width=3).grid(row=11,column=2)

#按钮5
Label(fm2,font=ft5,width=3).grid(row=12,column=0)
m_a=Button(fm2,text="本月转正",command=month_aug,font=ft3,borderwidth=1,relief=RAISED,width=12)
m_a.grid(row=13,column=1)
Label(fm2,font=ft5,width=3).grid(row=14,column=2)

#按钮6
Label(fm2,font=ft5,width=3).grid(row=15,column=0)
m_q=Button(fm2,text="试用期员工",command=month_period,font=ft3,borderwidth=1,relief=RAISED,width=12)
m_q.grid(row=16,column=1)
Label(fm2,font=ft5,width=3).grid(row=17,column=2)

#按钮7

Label(fm2,font=ft5,width=3).grid(row=18,column=0)
m_q=Button(fm2,text="本月离职员工",command=month_quit,font=ft3,borderwidth=1,relief=RAISED,width=12)
m_q.grid(row=19,column=1)
Button(fm2,text="上月离职员工",command=last_month_quit,font=ft3,borderwidth=1,relief=RAISED,width=12,bg="#45ff23").grid(row=20,column=1)


#按钮8

Label(fm2,font=ft5,width=3).grid(row=21,column=0)
m_q=Button(fm2,text="准备离职员工",command=pre_quit,font=ft3,borderwidth=1,relief=RAISED,width=12)
m_q.grid(row=22,column=1)
Label(fm2,font=ft5,width=3).grid(row=23,column=2)

#按钮9

Label(fm2,font=ft5,width=3).grid(row=24,column=0)
m_q=Button(fm2,text="住宿员工",command=do_emp,font=ft3,borderwidth=1,relief=RAISED,width=12)
m_q.grid(row=25,column=1)
Label(fm2,font=ft5,width=3).grid(row=26,column=2)

#按钮10
Label(fm2,font=ft5,width=3).grid(row=27,column=0)
m_q=Button(fm2,text="岗位宿舍调动",command=do_mod,font=ft3,borderwidth=1,relief=RAISED,width=12)
m_q.grid(row=28,column=1)
Label(fm2,font=ft5,width=3).grid(row=29,column=2)

top.mainloop()
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
