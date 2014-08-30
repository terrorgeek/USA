import sys
import MySQLdb
import re
import os
import xlrd
import xlwt

def connect_db():
    try:
        conn=MySQLdb.connect(host='localhost',user='root',passwd='',db='samplejob',port=3306)
    except MySQLdb.Error,e:
        print "Mysql Error %d: %s" % (e.args[0], e.args[1])
    return conn


def Procedure():
    conn=connect_db()
    cur_delete=conn.cursor()
    cur_delete.execute("delete from contacts")
    conn.commit()
    wb=xlrd.open_workbook("Contacts.xls")
    sheet=wb.sheets()[0]
    index=0
    for rownum in range(sheet.nrows):
        if rownum==0:
            continue
        sql='''insert into contacts (full_name,first_name,last_name,nick_name,phone1,other1,phone2,
               other2,email,extra_info,mailing_address,nature,industry,sub,loc,other,association) values ( '''
        for cell in sheet.row_values(rownum):
            if cell =="":
                sql+="null,"
            else:
                sql+="\""+str(cell)+"\","
        cur=conn.cursor()
        cur.execute(sql[0:len(sql)-1]+")")
        conn.commit()
        cur.close()
        print sql[0:len(sql)-1]+")"



def Subroutine1_Industry():
    legend={}#legend hashtable
    #assign the hashtable of legend
    f=open("Industy Legend.txt","r")
    string=""
    pattern=re.compile(r"(^[A-Z?]+)\s+([a-zA-Z ]+)")
    for line in f.readlines():
      if line!="" or pattern.match(line) or line!="/n":
          m=re.match(r"(^[A-Z?]+)\s+([a-zA-Z ]+)", str(line))
          legend[m.group(1)]=m.group(2)
      else:
          continue
    string=""
    conn=connect_db()
    answer_file=open("industry_sort.txt","a+");
    answer_file.write("Full Name"+" "*31)
    answer_file.write("Association"+" "*37)
   # answer_file.write("Industry"+" "*37)
   # answer_file.write("Sub-Industry"+" "*37)
    answer_file.write("Phone1"+" "*35+"\n\n\n")
    sql="select full_name,association,industry,sub,phone1,ic from contacts group by full_name order by industry,sub, last_name "
    cur=conn.cursor()
    cur.execute(sql)
    page_break="None"
    page_break2="None"
    for row in cur.fetchall():
        if str(row[2])!=page_break and str(row[3])==page_break2:
            answer_file.write("\n\n***********************************************          "+legend[str(row[2])]+":         **********************************************\n\n")
        if str(row[3])!=page_break2:
            if str(row[3])=="None":
                answer_file.write("\n\n***********************************************          "+legend[str(row[2])]+":         **********************************************\n\n")
            else:
                answer_file.write("\n\n***********************************************          "+legend[str(row[2])]+":"+legend[str(row[3])]+"          **********************************************\n\n")
            page_break2=str(row[3])

        page_break=str(row[2])

        for column in range(6):
            if column==2 or column==3:
                continue
            if column==5:
                continue
            if column!=5:
              if len(str(row[column]))>13:
                  answer_file.write(str(row[column])+(30-(len(str(row[column]))-13))*" ")
              elif len(str(row[column]))<13:
                  answer_file.write(str(row[column])+(30+(13-len(str(row[column]))))*" ")
              elif len(str(row[column]))==13:
                  answer_file.write(str(row[column])+" "*30)
        answer_file.write("\n");
    answer_file.close()


def Subroutine2_Location():
    page_break="None"
    conn=connect_db()
    file=open("location_sort.txt","a+")
    file.write("Full Name"+" "*30)
    file.write("Association"+" "*37)
    file.write("Phone Number"+" "*30+"\n\n\n")
    sql="select full_name,association,loc,phone1,phone2,ic from contacts group by full_name order by loc,last_name"
    cur=conn.cursor()
    cur.execute(sql)
    for row in cur.fetchall():
        if str(row[2])!=page_break:
            page_break=str(row[2])
            file.write("\n\n***********************************************          "+str(row[2])+":          **********************************************\n\n")
        for column in range(6):
            if column==2 or column==4:
                continue
            if column==3 and str(row[3])=="":
                if len(str(row[4]))>13:
                  file.write(str(row[4])+(30-(len(str(row[4]))-13))*" ")
                elif len(str(row[4]))<13:
                  file.write(str(row[4])+(30+(13-len(str(row[4]))))*" ")
                elif len(str(row[4]))==13:
                  file.write(str(row[4])+" "*30)
            if column==5:
                continue
            if column!=5:
             if len(str(row[column]))>13:
                 file.write(str(row[column])+(30-(len(str(row[column]))-13))*" ")
             elif len(str(row[column]))<13:
                 file.write(str(row[column])+(30+(13-len(str(row[column]))))*" ")
             elif len(str(row[column]))==13:
                 file.write(str(row[column])+" "*30)

        file.write("\n");
    file.close()


def Subroutine1_Industry_For_IC():
    legend={}#legend hashtable
    #assign the hashtable of legend
    f=open("Industy Legend.txt","r")
    string=""
    pattern=re.compile(r"(^[A-Z?]+)\s+([a-zA-Z ]+)")
    for line in f.readlines():
      if line!="" or pattern.match(line) or line!="/n":
          m=re.match(r"(^[A-Z?]+)\s+([a-zA-Z ]+)", str(line))
          legend[m.group(1)]=m.group(2)
      else:
          continue
    string=""
    conn=connect_db()
    answer_file=open("industry_sort_ic.txt","a+");
    answer_file.write("Full Name"+" "*31)
    answer_file.write("Association"+" "*37)
    answer_file.write("Phone1"+" "*35+"\n\n\n")
    sql="select full_name,association,industry,sub,phone1,ic from contacts group by full_name order by industry,sub, last_name "
    cur=conn.cursor()
    cur.execute(sql)
    page_break="None"
    page_break2="None"
    for row in cur.fetchall():
        if str(row[2])!=page_break and str(row[3])==page_break2:
            answer_file.write("\n\n***********************************************          "+legend[str(row[2])]+":         **********************************************\n\n")
        if str(row[3])!=page_break2:
            if str(row[3])=="None":
                answer_file.write("\n\n***********************************************          "+legend[str(row[2])]+":         **********************************************\n\n")
            else:
                answer_file.write("\n\n***********************************************          "+legend[str(row[2])]+":"+legend[str(row[3])]+"          **********************************************\n\n")
            page_break2=str(row[3])

        page_break=str(row[2])

        if str(row[5])=="1.0":
         for column in range(6):
            if column==2 or column==3:
                continue
            if column==5:
                continue
#                if str(row[5])=="1.0":
#                    if len(str(row[column]))>13:
#                      answer_file.write("1"+(30-(len(str(row[column]))-13))*" ")
#                    elif len(str(row[column]))<13:
#                      answer_file.write("1"+(30+(13-len(str(row[column]))))*" ")
#                    elif len(str(row[column]))==13:
#                      answer_file.write("1"+" "*30)
            if column!=5:
              if len(str(row[column]))>13:
                  answer_file.write(str(row[column])+(30-(len(str(row[column]))-13))*" ")
              elif len(str(row[column]))<13:
                  answer_file.write(str(row[column])+(30+(13-len(str(row[column]))))*" ")
              elif len(str(row[column]))==13:
                  answer_file.write(str(row[column])+" "*30)
         answer_file.write("\n");
    answer_file.close()


def Subroutine2_Location_For_IC():
    page_break="None"
    conn=connect_db()
    file=open("location_sort_ic.txt","a+")
    file.write("Full Name"+" "*30)
    file.write("Association"+" "*37)
    file.write("Phone Number"+" "*30+"\n\n\n")
    sql="select full_name,association,loc,phone1,phone2,ic from contacts group by full_name order by loc,last_name"
    cur=conn.cursor()
    cur.execute(sql)
    for row in cur.fetchall():
        if str(row[2])!=page_break:
            page_break=str(row[2])
            file.write("\n\n***********************************************          "+str(row[2])+":          **********************************************\n\n")
        if str(row[5])=="1.0":
         for column in range(6):
            if column==2 or column==4:
                continue
            if column==3 and str(row[3])=="":
                if len(str(row[4]))>13:
                  file.write(str(row[4])+(30-(len(str(row[4]))-13))*" ")
                elif len(str(row[4]))<13:
                  file.write(str(row[4])+(30+(13-len(str(row[4]))))*" ")
                elif len(str(row[4]))==13:
                  file.write(str(row[4])+" "*30)
            if column==5:
                continue

            if column!=5:
             if len(str(row[column]))>13:
                 file.write(str(row[column])+(30-(len(str(row[column]))-13))*" ")
             elif len(str(row[column]))<13:
                 file.write(str(row[column])+(30+(13-len(str(row[column]))))*" ")
             elif len(str(row[column]))==13:
                 file.write(str(row[column])+" "*30)

         file.write("\n");
    file.close()

Subroutine1_Industry()
Subroutine2_Location()
Subroutine2_Location_For_IC()
Subroutine1_Industry_For_IC()