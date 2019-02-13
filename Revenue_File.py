#Title:File splitting using Python
#version:4.0
#Changes: Dummy line introduced in each split file
#Progress bar introduced
#Check for duplicate records introduced on 31-03-2012
#Additional Report Item and KeyError Exception introduced on 25-04-2012
#Additional Progress Bar and stderr redirection introduced on 26-04-2012
#IO Exception handled for permission denied error on win2k3 on 27-04-2012
#Customer Count Added In The Report of Subdivision on 14-05-2012
#Report No. 5 and some formatting to existing reports done on 04-09-2012
#Design of Reports in Excel format introduced on 02-10-2012
#Change in file creation logic 15-08-2017
import sys
import os
import datetime
import string
import codes    #This includes the subdivision code names
import shutil
import pickle
import time
from progressbar import *
from xlwt import Workbook,easyxf,Formula
from tempfile import TemporaryFile
from operator import itemgetter, attrgetter
d={}
filedir={}
print "PYTHON File Splitter Ver 5.0\n"
#Fetching line count for progress bar
def getlinecount(filename):
        fs=open(filename,'r')
        line=len(list(fs))
        fs.close()
        return line
#Fetching date value from the file and storing it in variable
try:
        fi=open("kashmir.txt",'r')
except IOError:
        print "The system could not find the file kashmir.txt.Copy the file and try again"
        time.sleep(2)
        sys.exit()
for l,i in enumerate(fi):                     
        if (len(i.split(',')[2])==13 and len(i.split(','))==9 and l==2):
                filedate=i.split(',')[1][3:10]
                break
fi.close()
filedate=filedate.replace('-','')
dirname="BANK-"+filedate

#Removing current diretories created to avoid duplicate error message
for f in os.listdir(os.getcwd()):
                if f==filedate:
                        flag=raw_input('    "'+filedate+'"'+" DIRECTORY EXISTS PRESS n TO EXIT OR ANY KEY TO CONTINUE: \n")
                        if flag=='n':
                                sys.exit()
                        shutil.rmtree(f)
                        print 'Old Directory',f,'removed\n'
                elif f in (['newkashmir.txt','ERROR-'+filedate+'.txt',"COLL-"+filedate+".txt",'JMU-'+filedate+'.txt','duplicates.txt']):
                        os.remove(f)
                        print 'Old file',f,'is removed\n'
                elif f==filedate:
                        shutil.rmtree(filedate)
                elif f==dirname:
                        shutil.rmtree(dirname)
os.mkdir(dirname)
#Reformatting the Input file for valid subdiv code and excluding the codes of jammu region
fi=open("kashmir.txt",'r')
fo=open("newkashmir.txt",'a')
duplicate={}
linecount=getlinecount("kashmir.txt")
print "File Processing Started watch out for Errors if any\n\n"
widgets = ['Total Revenue Lines '+str(linecount)+':', Percentage(), ' ', Bar(marker=RotatingMarker()),
               ' ', ETA(), ' ', FileTransferSpeed()]
pbar = ProgressBar(widgets=widgets, maxval=linecount).start()
for l,i in enumerate(fi):
        if l>0:
                if i=='\n':continue
                if (len(i.split(',')[2])==13 and codes.names.has_key(i.split(',')[2][0:6]) and len(i.split(','))==9 and i.split(',')[2][0:2]<>'01' and i.split(',')[2][0:2]=='02' and i.split(',')[8]<>'\n' ):
                    dupindex=i.split(',')[1]+i.split(',')[2]+i.split(',')[5]
                    if duplicate.has_key(dupindex):
                            fd=open("duplicates.txt",'a')    
                            duplicate[dupindex]=duplicate[dupindex]+1    
                            fd.write(i)
                            fd.close()
                            
                    else:
                        pbar.update(l)
                        duplicate[dupindex]=1    
                        sud=i.split(',')[2][0:6]
                        value=string.atoi(i.split(',')[5])
                        bank=i.split(',')[8].rstrip("\n").upper()
                        dated=i.split(',')[1]
                        cat=bank+"      "+dated
                        if d.has_key(sud):
                                d[sud][2]=d[sud][2] + 1
                                filedir[sud].append(i.rstrip("\n").split(','))
                                if d[sud][0].has_key(bank):
                                        d[sud][0][bank]=d[sud][0][bank] + value
                                        d[sud][4][bank]=d[sud][4][bank] +1
                                else:
                                        d[sud][0][bank]=value
                                        d[sud][4][bank]=1
                                if d[sud][1].has_key(dated):
                                        d[sud][1][dated]=d[sud][1][dated] + value
                                else:
                                        d[sud][1][dated]=value
                                if d[sud][3].has_key(cat):
                                        d[sud][3][cat]=d[sud][3][cat] + value
                                else:
                                        d[sud][3][cat]=value        
                        else:
                                d[sud]=[{bank:value},{dated:value},1,{cat:value},{bank:1}]
                                filedir[sud]=[i.rstrip("\n").split(',')]
                        fo.write(i)
                        
                elif i.split(',')[2][0:2]=='01':
                    fj=open("JMU-"+filedate+".txt",'a')    
                    fj.write(i)
                    fj.close()
                elif i.split(',')[8]=='\n':
                    fb=open("BANKERROR-"+filedate+".txt",'a')    
                    fb.write(i)
                    fb.close()
                else:
                    fe=open("ERROR-"+filedate+".txt",'a')    
                    fe.write(i)
                    fe.close()
        else:
                continue
        

fo.close()
fi.close()
pbar.finish()

for f in os.listdir(os.getcwd()):
        if f=='ERROR-'+filedate+'.txt':
                print "\nRecords With Errors are In  "+f+" File\n"
                x='True'
        elif f=='BANKERROR-'+filedate+'.txt':
                print "\nRecords Without Bank Info In "+f+" File\n"
                x='True'
try:x
except NameError:
        pass
else:
   feed=raw_input('Input File Has Errors Press n to Exit  OR any Key to Continue:-\n')
   if feed=='n':
         sys.exit()
widgets = ['Total Subdivisions '+str(len(filedir))+':', Percentage(), ' ', Bar(marker=RotatingMarker()),
               ' ', ETA(), ' ', FileTransferSpeed()]
pbar = ProgressBar(widgets=widgets, maxval=len(filedir)).start()
for l,i in enumerate(sorted(filedir)):
    newdir=dirname+'/'+codes.names[i]+'-'+i+'-'+filedate
    os.mkdir(newdir)
    for j in sorted(filedir[i],key=itemgetter(8),reverse=True):
        if os.path.isfile(newdir+'\\'+i+j[8]+'.txt'):
                fs=open(newdir+'\\'+i+j[8]+'.txt','a')
                fs.write(str(j).replace('[','').replace(']','').replace("'",'')+"\n")
                fs.close()
                
                
        else:
                fs=open(newdir+'\\'+i+j[8]+'.txt','a')
                fs.write("Txn No.,Txn Date,Description,ChequeNo.,Cr/Dr,Transaction Amount(INR),Balance(INR),sol_id,br_code\n101  ,31-01-2012,0909099999999,        ,CR,1      ,          ,1256,DUMMYS\n")
                fs.write(str(j).replace('[','').replace(']','').replace("'",'')+"\n")
                fs.close()
    pbar.update(l)
    
pbar.finish()


colfilename="COLL-"+filedate+".xls"
#Data Dumping to External File
f=open(filedate+".pk",'wb')
pickle.dump(d,f)
f.close()
#Excel Cell Formatting Variable Definitions
st=easyxf('font: name Arial;'
'borders: left thick, right thick, top thick, bottom thick;'
'pattern: pattern solid, fore_colour red;'
)
st_cell=easyxf('font: name Arial;'
'borders: left thin, right thin, top thin, bottom thin;'
'pattern: pattern solid, fore_colour white;'
)
st_header=easyxf('font: name Arial;'
'borders: left thick, right thick, top thick, bottom thick;'
'pattern: pattern solid, fore_colour aqua;'
)
#Header Creation
book = Workbook()
dash=book.add_sheet('Dashboard',cell_overwrite_ok=True)
cell = easyxf('pattern: pattern solid, fore_colour light_blue')
for i in range(0,32,1):
        for j in range(0,21,1):
                dash.write(i,j,None,cell)
dash.write_merge(0,4,0,20,'JKPDD BANK COLLECTION REPORT DASHBOARD  FOR THE MONTH OF '+filedate[0:2]+'-'+filedate[-4:],easyxf('font: name Showcard Gothic,bold true,colour white,height 400;'
'pattern: pattern solid, fore_colour green;'
'borders: left thick, right thick, top thick, bottom thick;'
'align: vertical center, horizontal center;'
))
dash.write_merge(5,8,0,10,'MONTHLY REPORT DETAILS',easyxf('font: name Showcard Gothic,bold true,colour white,height 300;'
'pattern: pattern solid, fore_colour light_blue;'
'borders: left thick, right thick, top thick, bottom thick;'
'align: vertical center, horizontal center;'
))
dash.write_merge(5,8,11,20,'MONTHLY REVENUE HIGHLIGHTS',easyxf('font: name Showcard Gothic,bold true,colour white,height 300;'
'pattern: pattern solid, fore_colour light_blue;'
'borders: left thick, right thick, top thick, bottom thick;'
'align: vertical center, horizontal center;'
))
pallete=easyxf('font: name Tw Cen MT,bold true,colour white,height 220;'
'pattern: pattern solid, fore_colour ocean_blue;'
'borders: left thick, right thick, top thick, bottom thick;'
'align: vertical center, horizontal center,wrap True,shrink_to_fit True;'
)
link1='HYPERLINK("['+colfilename+']\'SD-COUNT-AMOUNT\'!a1";"1. SUB DIVISION COUNT & AMOUNT DETAILS")'
link2='HYPERLINK("['+colfilename+']\'SD-BANK-AMOUNT\'!a1";"2. SUB DIVISION BANK & AMOUNT DETAILS")'
link3='HYPERLINK("['+colfilename+']\'SD-DATE-AMOUNT\'!a1";"3. SUB DIVISION DATE & AMOUNT DETAILS")'
link4='HYPERLINK("['+colfilename+']\'BANK-AMOUNT\'!a1";"4. BANK & AMOUNT DETAILS")'
link5='HYPERLINK("['+colfilename+']\'SD-BANK-DATE\'!a1";"5. SUB DIVISION BANK & DATE & AMOUNT DETAILS")'
dash.write_merge(9,10,12,17,Formula(link1),pallete)
dash.write_merge(11,12,12,17,Formula(link2),pallete)
dash.write_merge(13,14,12,17,Formula(link3),pallete)
dash.write_merge(15,16,12,17,Formula(link4),pallete)
dash.write_merge(17,18,12,17,Formula(link5),pallete)
#Dashboard Figures
list1=[]
for i in d:
        try:       
            s1=0    
            for j in d[i][0]:
                    s1=s1+d[i][0][j]
            list1.append((codes.names[i],s1,d[i][2]))
        except KeyError:
                 continue
                           
def get_sort(name,value):
    if value=='max-name':
            list2=[]
            for i in sorted(list1,key=itemgetter(1),reverse=True):
                    list2.append(i[1])
            if list2.count(list2[0]) > 1:
                    st=''                    
                    for i in range(0,list2.count(list2[0]),1):    
                            st=sorted(list1,key=itemgetter(1),reverse=True)[i][0]+' '+st
                    return st
            else:
                    return sorted(name,key=itemgetter(1),reverse=True)[0][0]
    elif value=='min-name':
            list2=[]
            for i in sorted(list1,key=itemgetter(1)):
                    list2.append(i[1])
            if list2.count(list2[0]) > 1:
                    st=''                    
                    for i in range(0,list2.count(list2[0]),1):    
                            st=sorted(list1,key=itemgetter(1))[i][0]+' '+st
                            return st
            else:
                    return sorted(name,key=itemgetter(1))[0][0]
    elif value=='max-amt':
            return sorted(name,key=itemgetter(1),reverse=True)[0][1]
    elif value=='min-amt':
            return sorted(name,key=itemgetter(1))[0][1]
    elif value=='max-cntname':
            list2=[]
            for i in sorted(list1,key=itemgetter(2),reverse=True):
                    list2.append(i[2])
            if list2.count(list2[0]) > 1:
                    st=''                    
                    for i in range(0,list2.count(list2[0]),1):    
                            st=sorted(list1,key=itemgetter(2),reverse=True)[i][0]+','+st
                    return st
            else:
                    return sorted(name,key=itemgetter(2),reverse=True)[0][0]
    elif value=='min-cntname':
            list2=[]
            for i in sorted(list1,key=itemgetter(2)):
                    list2.append(i[2])
            if list2.count(list2[0]) > 1:
                    st=''                    
                    for i in range(0,list2.count(list2[0]),1):    
                            st=sorted(list1,key=itemgetter(2))[i][0]+' '+st
                    return st
            else:
                    return sorted(name,key=itemgetter(2))[0][0]
    elif value=='max-cnt':
            return sorted(name,key=itemgetter(2),reverse=True)[0][2]

    elif value=='min-cntname':
            list2=[]
            for i in sorted(list,key=itemgetter(2)):
                    list2.append(i[2])
            if list2.count(list2[0]) > 1:
                    st=''                    
                    for i in range(0,list2.count(list2[0]),1):    
                            st=sorted(list1,key=itemgetter(2))[i][0]+' '+st
                    return st
    elif value=='min-cnt':
            return sorted(name,key=itemgetter(2))[0][2]
banklist=[]
for i in d:
    for j in d[i][0]:
            banklist.append((j,d[i][0][j]))
bankcust={}
for i in d:
    for j in d[i][4]:
            if bankcust.has_key(j):
                    bankcust[j]=bankcust[j]+d[i][4][j]
            else:
                    bankcust[j]=d[i][4][j]

dash.write_merge(9,10,1,6,'MAXIMUM-REVENUE-SUBDIVISION',pallete)
dash.write_merge(9,10,7,8,get_sort(list1,'max-name'),pallete)
dash.write_merge(9,10,9,10,'Rs '+str(get_sort(list1,'max-amt'))+'/=',pallete)
dash.write_merge(11,12,1,6,'MINIMUM-REVENUE-SUBDIVISION',pallete)
dash.write_merge(11,12,7,8,get_sort(list1,'min-name'),pallete)
dash.write_merge(11,12,9,10,'Rs '+str(get_sort(list1,'min-amt'))+'/=',pallete)
dash.write_merge(13,14,1,6,'MAXIMUM-CONSUMERS-SUBDIVISION',pallete)
dash.write_merge(13,14,7,8,get_sort(list1,'max-cntname'),pallete)
dash.write_merge(13,14,9,10,get_sort(list1,'max-cnt'),pallete)
dash.write_merge(15,16,1,6,'MINIMUM-CONSUMERS-SUBDIVISION',pallete)
dash.write_merge(15,16,7,8,get_sort(list1,'min-cntname'),pallete)
dash.write_merge(15,16,9,10,get_sort(list1,'min-cnt'),pallete)
dash.write_merge(17,18,1,6,'MAXIMUM-REVENUE-BANK-BRANCH',pallete)
dash.write_merge(17,18,7,8,sorted(banklist,key=itemgetter(1),reverse=True)[0][0],pallete)
dash.write_merge(17,18,9,10,'Rs '+ str(sorted(banklist,key=itemgetter(1),reverse=True)[0][1])+'/=',pallete)
#Row 19,20 written inside Report 4
dash.write_merge(21,22,1,6,'MAX-CUSTOMER-VISITS-BANK-BRANCH',pallete)
dash.write_merge(21,22,7,8,sorted(bankcust.iteritems(),key=itemgetter(1),reverse=True)[0][0],pallete)
dash.write_merge(21,22,9,10,str(sorted(bankcust.iteritems(),key=itemgetter(1),reverse=True)[0][1]),pallete)
dash.write_merge(23,24,1,6,'MIN-CUST-VISITS-BANK-BRANCH',pallete)
dash.write_merge(23,24,7,8,sorted(bankcust.iteritems(),key=itemgetter(1))[0][0],pallete)
dash.write_merge(23,24,9,10,str(sorted(bankcust.iteritems(),key=itemgetter(1))[0][1]),pallete)
#Row 25,26 written inside Report 1
dash.write_merge(27,28,1,8,'NUMBER OF COLLECTION BANK BRANCHES',pallete)
dash.write_merge(27,28,9,10,str(len(list(bankcust))),pallete)
dash.write_merge(29,30,1,8,'NUMBER OF SUBDIVISIONS INVOLVED',pallete)
dash.write_merge(29,30,9,10,str(len(list(d))),pallete)

#Output for report 1 starts from here
sheet1=book.add_sheet("SD-COUNT-AMOUNT")
sheet1.write(0,0,'SNo',st_header)
sheet1.write(0,1,'SUBDIV',st_header)
sheet1.write(0,2,'NAME',st_header)
sheet1.write(0,3,'AMOUNT',st_header)
sheet1.write(0,4,'CUST_COUNT',st_header)
total=0
sno=0
tc=0
xl=0
for i in sorted(d):
        try:
                x=sum(d[i][0].values())
                sno=sno+1
                xl=xl+1
                sheet1.write(xl,0,sno,st_cell)
                sheet1.write(xl,1,i,st_cell)
                sheet1.write(xl,2,codes.names[i],st_cell)
                sheet1.write(xl,3,x,st_cell)
                sheet1.write(xl,4,d[i][2],st_cell)
                total=total+x
                tc=tc+d[i][2]
        except KeyError:
                continue
sheet1.write_merge(xl+1,xl+1,0,2,'Total Collection',st)
dash.write_merge(25,26,1,8,'TOTAL BANK COLLECTION AMOUNT',pallete)
dash.write_merge(25,26,9,10,'Rs '+str(total)+'/=',pallete)
sheet1.write(xl+1,3,total,st)
sheet1.write(xl+1,4,tc,st)
#Output for Report 2 Starts Here
sheet2=book.add_sheet("SD-BANK-AMOUNT")
sheet2.write(0,0,'SUBDIV',st_header)
sheet2.write(0,1,'NAME',st_header)
sheet2.write(0,2,'BANK',st_header)
sheet2.write(0,3,'CUST_COUNT',st_header)
sheet2.write(0,4,'AMOUNT',st_header)
total=0
xl=0
for i in sorted(d):
        try:
                m=0
                xl=xl+1
                for j in sorted(d[i][0]):
                        sheet2.write(xl,0,i,st_cell)
                        sheet2.write(xl,1,codes.names[i],st_cell)
                        sheet2.write(xl,2,j.rstrip("\n"),st_cell)
                        sheet2.write(xl,3,d[i][4][j],st_cell)
                        sheet2.write(xl,4,d[i][0][j],st_cell)
                        xl=xl+1
                        m=m+d[i][0][j]
                sheet2.write(xl,0,'Total Collection for SubDivision',st)
                sheet2.write_merge(xl,xl,1,3,codes.names[i],st)
                sheet2.write(xl,4,m,st)             
                total=total+m
        except KeyError:
                continue
sheet2.write_merge(xl+1,xl+1,0,3,'Total Collection for All SubDivision',st)
sheet2.write(xl+1,4,total,st)
#Output for Report 3 Starts Here
sheet3=book.add_sheet("SD-DATE-AMOUNT")
sheet3.write(0,0,'Sno',st_header)
sheet3.write(0,1,'SUBDIVCODE',st_header)
sheet3.write(0,2,'NAME',st_header)
sheet3.write(0,3,'DATE',st_header)
sheet3.write(0,4,'AMOUNT',st_header)
xl=0
total=0
for i in sorted(d):
        try:
                m=0
                sno=1
                xl=xl+1
                for j in sorted(d[i][1]):
                        sheet3.write(xl,0,sno,st_cell)
                        sheet3.write(xl,1,i,st_cell)
                        sheet3.write(xl,2,codes.names[i],st_cell)
                        sheet3.write(xl,3,j,st_cell)
                        sheet3.write(xl,4,d[i][1][j],st_cell)
                        xl=xl+1
                        m=m+d[i][1][j]
                        sno=sno+1
                sheet3.write_merge(xl,xl,0,1,"Total Collection for SubDiv:",st)
                sheet3.write_merge(xl,xl,2,3,codes.names[i],st)
                sheet3.write(xl,4,m,st)
                total=total+m
        except KeyError:
                continue
sheet3.write_merge(xl+1,xl+1,0,3,"Total Collection ",st)
sheet3.write(xl+1,4,total,st)
#Output for report 4 starts from here
sheet4=book.add_sheet("BANK-AMOUNT")
sheet4.write(0,0,'Sno',st_header)
sheet4.write(0,1,'BANKNAME',st_header)
sheet4.write(0,2,'AMOUNT',st_header)
bank={}
sno=1
total=0
xl=0
for i in sorted(d):
        for j in sorted(d[i][0]):
                if bank.has_key(j):
                        bank[j]=bank[j]+d[i][0][j]
                else:
                        bank[j]=d[i][0][j]
for i in sorted(bank):
        xl=xl+1
        sheet4.write(xl,0,sno,st_cell)
        sheet4.write(xl,1,i,st_cell)
        sheet4.write(xl,2,bank[i],st_cell)
        total=total+bank[i]
        sno=sno+1
sheet4.write_merge(xl+1,xl+1,0,1,"Total Bank Collection",st)
sheet4.write(xl+1,2,total,st)
dash.write_merge(19,20,1,6,'MINIMUM-REVENUE-BRANCH',pallete)
dash.write_merge(19,20,7,8,sorted(bank.iteritems(),key=itemgetter(1))[0][0],pallete)
dash.write_merge(19,20,9,10,'Rs '+str(sorted(bank.iteritems(),key=itemgetter(1))[0][1])+'/=',pallete)
#Output from Report 5 starts from here
sheet5=book.add_sheet("SD-BANK-DATE")
sheet5.write(0,0,'SUBDIV',st_header)
sheet5.write(0,1,'NAME',st_header)
sheet5.write(0,2,'BANK',st_header)
sheet5.write(0,3,'DATED',st_header)
sheet5.write(0,4,'AMOUNT',st_header)
total=0
xl=0
for i in sorted(d):
        try:
                m=0
                xl=xl+1
                for j in sorted(d[i][3]):
                        sheet5.write(xl,0,i,st_cell)
                        sheet5.write(xl,1,codes.names[i],st_cell)
                        sheet5.write(xl,2,j[0:6].rstrip("\n"),st_cell)
                        sheet5.write(xl,3,j[-10:].rstrip("\n"),st_cell)
                        sheet5.write(xl,4,d[i][3][j],st_cell)
                        m=m+d[i][3][j]
                        xl=xl+1
                sheet5.write_merge(xl,xl,0,1,"Total Collection for SUB Div",st)
                sheet5.write_merge(xl,xl,2,3,codes.names[i],st)
                sheet5.write(xl,4,m,st)
                total=total+m
        except KeyError:
                continue
sheet5.write_merge(xl+1,xl+1,0,3,"Total Collection of All Subdiv's ",st)
sheet5.write(xl+1,4,total,st)
book.save(colfilename)
book.save(TemporaryFile())
#Wrapping the directories and files created above into a single folder
os.mkdir(filedate)
src='BANK-'+filedate
des=filedate
shutil.move(src,des)
for i in ['JMU-'+filedate+'.txt',' ERROR-'+filedate+'.txt','newkashmir.txt','KASHMIR.txt','COLL-'+filedate+'.xls','BANKERROR-'+filedate+'.txt','duplicates.txt']:
        copystring='move '+i+' '+filedate+' 1>nul 2>&1'
        os.system(copystring)
