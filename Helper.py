import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from tkinter import *
from tkinter import ttk

import datetime
import scrolledframe as SF
from copy import deepcopy
import math

print("Libraries imported")

def getRecipients(email):
    header=email['payload']['headers']
    o=[]
    for comp in header:
        if comp['name']=='Delivered-To':
            o.append(comp['value'])
    return(o)
def getTFS(email):
    header=email['payload']['headers']
    to=[]
    sub=""
    sent=""
    for comp in header:
        if comp['name']in['To','Cc']:
            to.append(comp['value'])
        elif comp['name']=='Subject':
            sub=comp['value']
        elif comp['name']=='From':
            sent=comp['value']
    return((to,sent,sub))
def getEmail(txt):
    txt=txt.lower()
    if "<" in txt:
        return(txt.split('<')[1].split('>')[0])
    else:
        return(txt)
def getDate(email):
    header=email['payload']['headers']
    date=""
    for comp in header:
        if comp['name']=='Date':
            date=comp['value']
            break
    return(date)
def fetchEmails(dateStart=(23,5,1),dateEnd="Today"):    
    creds = None
    if os.path.exists('token.json'):
        print("Token file found")
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing token")
            try:
                creds.refresh(Request())
            except:
                print("Unknown error, attempting reauthentication")
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            print("Saving token")
            token.write(creds.to_json())

    try:
        print("Grabbing most recent emails")
        service = build('gmail', 'v1', credentials=creds)
        
        sday=f"{dateStart[0]}/{dateStart[1]}/{dateStart[2]}"
        if dateEnd=="Today":
            query=f"after:{sday}"
        else:
            sday=f"{dateEnd[0]}/{dateEnd[1]}/{dateEnd[2]}"
            query=f"before:{eday} after:{sday}"
        pgtk=None
        fmsgs=[]
        while 1:
            if pgtk==None:
                recents = service.users().messages().list(userId='me',q=query,maxResults=500).execute()
            else:
                recents = service.users().messages().list(userId='me',q=query,maxResults=500,pageToken=pgtk).execute()
            pgtk=recents.get('nextPageToken')
            msgs=recents['messages']
            if pgtk!=None:
                fmsgs+=msgs
            else:
                break
        msgs=fmsgs
        print(f"{len(msgs)} emails recieved")
        A=datetime.datetime.now()
        msg=msgs[1]
        ems=[]
        it=0
        b=0
        for msg in msgs:
            email=service.users().messages().get(userId='me',id=msg['id'], format='metadata').execute()
            try:
                date=getDate(email)
                date=date[:(date+' (').index(" (")]
                form='%a, %d %b %Y %H:%M:%S %z'
                form=form[4*("," not in date):]
                try:
                    ems.append(getTFS(email)+(deepcopy(datetime.datetime.strptime(date,form)),))
                except:
                    form='%a, %d %b %Y %H:%M:%S %Z'
                    ems.append(getTFS(email)+(deepcopy(datetime.datetime.strptime(date,form)),))
##                print(date,form)
            except:
                print(date,form)
            it+=1
            per=100*it/len(msgs)
            if int(per)>b:
                b=int(per)
                delta=(datetime.datetime.now()-A)
                rem=len(msgs)-it
                print(f"{b}% complete, estimated time remaining: {format((rem/it)*delta)}")
        return(ems)
    except HttpError as error:
        print(f'An error occurred: {error}')
def uniEm(emails):
    for email in emails:
##        T,F,S=getTFS(email)
        try:
            T,F,S,D=email
        except:
            print(email)
        for Ts in T:
            if getEmail(Ts) not in To:
                To.append(getEmail(Ts))
        if getEmail(F) not in Sent:
            Sent.append(getEmail(F))
def saveEms():
    monid=int(repr(datetime.date.today()).split(',')[1][1:])-1
    mon=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][monid]
    year=repr(datetime.date.today()).split(',')[0].split('(')[1]
    filename=f'Emails {year}-{mon}.txt'
    with open(filename,'w',encoding='utf-8') as f:
            f.write(repr(Ems))
def loadEms():
    global Ems
    monid=int(repr(datetime.date.today()).split(',')[1][1:])-1
    mon=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][monid]
    year=repr(datetime.date.today()).split(',')[0].split('(')[1]
    filename=f'Emails {year}-{mon}.txt'
    print(filename)
    if os.path.exists(filename) or True:
        print("Email file found")
        with open(filename,'r',encoding='utf-8') as f:
            Ems=eval(f.read())
        return(True)
    print("Not found")
    return(False)
##def loadProj():
##    filename="Projects.txt"
##    if os.path.exists(filename):
##        with open(filename,'r',encoding='utf-8') as f:
##            txt=f.read()
##        for line in txt.split('\n'):
##            if line!="":
##                a,b=line.split(':',1)
##                projs[a]=eval(b)
##def saveProj():
##    filename="Projects.txt"
##    with open(filename,'w') as f:
##        for key in projs.keys():
##            f.write(key+":"+repr(projs[key])+"\n")
def loadProj():
    filename="Projects.txt"
    if os.path.exists(filename):
        with open(filename,'r',encoding='utf-8') as f:
            txtP,txtTP,txtKW,txtGP,txtOV,txtKN=f.read().split("\n::\n")
            for line in txtP.split('\n'):
                if line!="":
                    a,b=line.split(':',1)
                    projs[a]=eval(b)
            for line in txtTP.split('\n'):
                if line!="":
                    a,b=qsplit(line,":","'")
                    tentprojs[a]=eval(b)
            for line in txtKW.split('\n'):
                if line!="":
                    a,b=qsplit(line,":","'")
                    kword[a]=eval(b)
            for line in txtGP.split('\n'):
                if line!="":
                    a,b=qsplit(line,":","'")
                    groups[a]=eval(b)
            for line in txtOV.split('\n'):
                if line!="":
                    try:
                        a,b=qsplit(line,":","'")
                    except:
                        a,b=qsplit(line,":",'"')
                    knowns[a]=(b)
def saveProj():
    filename="Projects.txt"
    with open(filename,'w',encoding='utf-8') as f:
        for key in projs.keys():
            f.write(key+":"+repr(projs[key])+"\n")
        f.write("::\n")
        for key in tentprojs.keys():
            f.write(key+":"+repr(tentprojs[key])+"\n")
        f.write("::\n\n")
        for key in kword.keys():
            f.write(key+":"+repr(kword[key])+"\n")
        f.write("::\n\n")
        for key in groups.keys():
            f.write(key+":"+repr(groups[key])+"\n")
        f.write("::\n\n")
        for key in knowns.keys():
            f.write(key+":"+knowns[key]+"\n")
        f.write("::\n\n")
def newComb(row):
    global ssproj
    col=len(people[row][1])
    Co=ttk.Combobox(f3)
    people[row][1].append(Co)
    Co.grid(column=col+1, row=(row+2), sticky="w")
    Co['values'] = ssproj
    Co.state(['readonly'])
    def func(x,y):
        def ff(a):
            global people,projs
            updateProjMems()
            val=people[x][1][y].get()
            if val!="" and (y+1)==len(people[x][1]):
                newComb(x)
            name=people[x][0]['text']
            if val!="":
                if name not in projs[val]:
                    projs[val].append(name)
            if 1:
                for key in projs.keys():
                    if name in projs[key]:
                        rem=1
                        for c in people[x][1]:
                            if c.get()==key:
                                rem=0
                        if rem:
                            projs[key].remove(name)
            regenProj()
        return ff
    people[row][1][col].bind('<<ComboboxSelected>>', func(row,col))
def updateProjDropdown():
    global ssproj,sssproj
    ssproj=fsorted(['']+list(projs.keys()))
    sssproj=[x for x in ssproj if x!=""]
    for person in people:
        for c in person[1]:
            c['values'] = ssproj
def updateProjMems():
    for proj in projects:
        proj[1]['text']= "\n"+",\n".join(projs[proj[0]['text']])+"\n"
def addProj():
    if Ent.get()!="" and not (list(Ent.get()).count(' ')==len(Ent.get()) and len(Ent.get())>0):
        projs[Ent.get()]=projs.get(Ent.get(),[])
        Ent['text']=""
        updateProjDropdown()
        regenProj()
        applyKword()
def regenProj():
    global projects
    for proj in projects:
        proj[0].destroy()
        proj[1].destroy()
        proj[2].destroy()
    projects=[]
    for i,key in enumerate(sssproj):
        projects.append([ttk.Label(f2,text=key),ttk.Label(f2,text="\n"+",\n".join(sorted(projs[key]))+"\n"),Entry(f2)])
        projects[-1][0].grid(column=0, row=(i+4), sticky="w")
        projects[-1][1].grid(column=2, row=(i+4), sticky="w")
        projects[-1][2].grid(column=5, row=(i+4), sticky="nsw")
        projects[-1][2].delete(0, END)
        projects[-1][2].insert(0, ", ".join(kword.get(key,[key])))
        
    regenEmails()
def regenPeople():
    global people
    for person in people:
        person[0].destroy()
        for c in person[1]:
            c.destroy()
    people=[]
    for i,sender in enumerate(esorted(Sent+To)[pemlow:pemlow+pemwidth]):
        people.append([ttk.Label(f3,text=sender),[]])
        people[-1][0].grid(column=0, row=(i+2), sticky="w")
        newComb(i)
        for key in sorted(projs.keys()):
            if sender in projs[key]:
                newComb(i)
                people[-1][1][-2].set(key)
def esorted(x):
    np=list(set(x[:]))
    global ESparams,pemlen
    if ESparams.get('domainfirst',True):
        np=[c[0] for c in sorted([(c,"@".join(c.split("@")[::-1])) for c in np],key=lambda e:e[1])]
    else:
        np=sorted(np)
    sc=ESparams.get('search',"")
    np=[v for v in np if sc in v]
    pemlen=len(np)
    pnem(0,False)
    return(np)
def fsorted(x):
    return([c[0] for c in sorted([(c,c.lower()) for c in x],key=lambda e:e[1])])
def ffsorted(x):
    if CB1.get()=="Sender":
        return([c[0] for c in sorted([(c,"@".join(getEmail(c[1]).split("@")[::-1]).lower()) for c in x],key=lambda e:e[1])])
    else:
        return(x)
def emsort():
    global emlen,prem
    
    oEms=[]
    f1Ems=[]
    for i,lbl in enumerate(prem):
        if lbl=="" and C1.instate(['selected']):
            f1Ems.append(Ems[i])
        elif lbl=="???" and C2.instate(['selected']):
            f1Ems.append(Ems[i])
        elif lbl=="---" and C3.instate(['selected']):
            f1Ems.append(Ems[i])
        elif C4.instate(['selected']):
            f1Ems.append(Ems[i])
        
    for i,Email in enumerate(ffsorted(f1Ems)):
        if E41.get() in Email[1]:
            if E42.get() in repr(Email[0]):
                if E43.get() in repr(Email[0])+Email[1]:
                    if E44.get() in Email[2]:
                        oEms.append(Email)
    emlen=len(oEms)
    nem(0,False)
    return(oEms)
def regenEmails():
    global emRow, L3
    global emComb,Emrow,suEms
    for row in Emrow:
        for it in row:
            it.destroy()
    emComb=[]
    Emrow=[]
    suEms=emsort()
    for i,Email in enumerate(suEms[emlow:emlow+emwidth]):
        Emc=[]
        col=["#efefef","#cfcfcf"][i%2]

        Ibb=ttk.Label(f4,text="")
        Ibb.grid(column=0, row=2*i+3, sticky="nswe",columnspan=8)
        Ibb['background']=col
        
        I=ttk.Label(f4,text=getEmail(Email[1]))
        I.grid(column=0, row=2*i+3, sticky="w")
        I['background']=col
        Emc.append(I)
        I=ttk.Label(f4,text=commajoin([getEmail(x)for x in Email[0]]))
        I.grid(column=2, row=2*i+3, sticky="w")
        I['background']=col
        Emc.append(I)
        subjj=Email[2]
        I=ttk.Label(f4,text=commajoin(subjj.split(" "),100," "))
        I.grid(column=4, row=2*i+3, sticky="w")
        I['background']=col
        Emc.append(I)
        #labeling is here
        Ib=None
        Ii=ttk.Label(f4,text="")
        opts=list(projs.keys())
        ans=knowns.get(repr(Email),"")
        qr=0
        if ans!="":
            Ii['text']="Override"
        if ans=="":
            if Me not in [getEmail(x)for x in Email[0]] and Me != getEmail(Email[1]):
                ans="Junk"
                Ii['text']="Doesnt include "+Me
            for opt in opts:
                for key in kword.get(opt,[]):
                    if key!="":
                        if key.lower() in Email[2].lower():
                            ans=opt
                            qr=1
                            qkey=key
        if ans!="" and Ii['text']=="":
            Ii['text']="Keyword"
        if ans=="":
            ans=groups.get(repr(set(Email[0]+[Email[1]])),"")
        if ans!="" and Ii['text']=="":
            Ii['text']="Group-"+str(set(Email[0]+[Email[1]]))
        opt1,att1=logic(Email,opts,projs)
        opt=opt1[:]
        if len(opt1)!=1:
            opt2,att2=logic(Email,opts,tentprojs)
            opt=opt2[:]
            if len(opt2)!=1:
                opt=opt1
        if ans=="":
            if len(opt)==1:
                if opt==opt1:
                    Ii['text']="Filter"
                else:
                    Ii['text']="Smart Fill"
                Ib=ttk.Button(f4,text="Mark as correct",command=func2(i))
                Ib.grid(column=7, row=2*i+3, sticky="w")
        if ans!="":
            if len(opt)>1:
                groups[repr(set(Email[0]+[Email[1]]))]=ans
            opt=[ans]
        #Answer here
        Co=ttk.Combobox(f4)
        Co.unbind_class('TCombobox',"<MouseWheel>")
        emComb.append(Co)
        Emc.append(Co)
        Co.grid(column=6, row=(2*i+3), sticky="w")
        I=ttk.Separator(f4, orient=HORIZONTAL)
        I.grid(column=0, row=2*i+4, columnspan=8, sticky="we")
        Emc.append(I)
        Co['values'] = ssproj
        Co.state(['readonly'])
        Co.bind('<<ComboboxSelected>>', func2(i))
        if len(opt)==1:
            Co.set(opt[0])
        elif len(opt)>1:
            Co.set("???")
        elif (att1 and len(opt1)==0) and (att2 and len(opt2)==0):
            Co.set("---")
        if Ib!=None:
            Emc.append(Ib)
        Ii.grid(column=8, row=2*i+3, sticky="w")
        Emc.append(Ii)
        Emc.append(Ibb)
        Emrow.append(Emc)
        if qr:
            print("Keyword found", qkey, "on line",i)
##            print(f"Line {i}")
##            func2(i)()
##            print("Uncaught Keyword")
##            error
##            applyKwords()
##            return(0)
##    A=datetime.datetime.now()
    blank,partial = labelData()
##    print(f"Label time: {datetime.datetime.now()-A}")
    L3['text']=f"{blank} Blank\n{partial} Partial"
def labelData():
    global prem,nset
    part=0
    blank=0
    conf=0
    amb=0
    comp=0
    tot=0
    prem=[]
    nset={}
    for i,Email in enumerate(Ems):
        opts=list(projs.keys())
        ans=knowns.get(repr(Email),"")
        qr=0
        if ans=="":
            if Me not in [getEmail(x)for x in Email[0]] and Me != getEmail(Email[1]):
                ans="Junk"
            for opt in opts:
                for key in kword.get(opt,[]):
                    if key!="":
                        if key.lower() in Email[2].lower():
                            ans=opt
        if ans=="":
            ans=groups.get(repr(set(Email[0]+[Email[1]])),"")
        opt1,att1=logic(Email,opts,projs)
        opt=opt1[:]
        if len(opt1)!=1:
            opt2,att2=logic(Email,opts,tentprojs)
            opt=opt2[:]
            if len(opt2)!=1:
                opt=opt1
        if ans!="":
            opt=[ans]
        if len(opt)==1:
            comp+=1
            prem.append(opt[0])
            DATE=format(Email[3],"%m-%d")
            se,rc=nset.get(DATE,{}).get(opt[0],[0,0])
            if Me == getEmail(Email[1]):
                se+=1
            else:
                rc+=1
            nset[DATE]=nset.get(DATE,{})
            nset[DATE][opt[0]]=[se,rc]
        elif len(opt)>1:
            part+=1
            amb+=1
            prem.append("???")
        elif (att1 and len(opt1)==0) and (att2 and len(opt2)==0):
            part+=1
            conf+=1
            prem.append("---")
        else:
            blank+=1
            prem.append("")
        tot+=1
    return(blank,part)
def func2(row):
    def ff(a=None):
        global Emrow,emlow,knowns
        Email=suEms[row+emlow]
        Rec,Sent,Subj=Email[0:3]
        Row=Emrow[row]
        for person in [Sent]+Rec:
            Person=getEmail(person)
            if Person!=Me:
                Val=Row[3].get()
                knowns[repr(Email)]=Val
                if Val not in["","???"]:
                    ist=projs[Val]
                    if Person not in ist:
                        ist.append(Person)
                    projs[Val]=ist
        regenProj()
    return ff
def logic(Email,uopts,projs):
    att=0
    oopts=uopts[:]
    opts=uopts[:]
    for rec in Email[0]+[Email[1]]:
        rec=getEmail(rec)
        nopt=[]
        for opt in opts:
            if rec in projs.get(opt,[]):
                nopt.append(opt)
        bas=0
        for opt in oopts:
            if rec in projs.get(opt,[]):
                bas=1
                att=1
        if bas:
            opts=nopt
    if att:
        return(opts,att)
    else:
        return([],att)
def commajoin(ist,leng=40,val=", "):
    l=""
    c=0
    for i in ist:
        l=l+i+val
        c+=len(i)+2
        if c>leng:
##            print("too long")
            l=l+"\n"
            c=0
    if l[-1:]=="\n":
        l=l[:-1]
    if l[-2:]==", ":
        l=l[:-2]
    return(l)
def fixEms():
    global Ems
    for i,Email in enumerate(Ems):
        try:
            Ems[i]=(qsplit((Email[0])[0],", "),Email[1],Email[2],Email[3])
        except:
            print(Email,i,sep="\n")
def qsplit(txt,sep,quote='"'):
    o=[]
    buf=""
    q=0
    for c in txt:
        if c==quote:
            q=not q
        buf+=c
        if len(buf)>=len(sep) and not q:
            if buf[-len(sep):]==sep:
                o.append(buf[:-len(sep)])
                buf=""
    o.append(buf)
    return(o)
def nem(x,regen=True):
##    print("Nem entered")
    global emwidth,emlow,emlen,tentprojs
    emlow+=x
    if emlen<100:
        emwidth=emlen
    else:
        emwidth=100
    if emlow<0:
        emlow=0
    if emlow+emwidth>emlen:
        emlow=emlen-emwidth
    if x==0:
        tentprojs=deepcopy(projs)
    
    L1['text']=f"Showing: {1+emlow}-{emlow+emwidth} of {emlen}"
    if regen:
        regenEmails()
##    print("Nem exited")
def pnem(x,regen=True):
    global pemwidth,pemlow,pemlen
    pemlow+=x
    if pemlen<200:
        pemwidth=pemlen
    else:
        pemwidth=200
    if pemlow<0:
        pemlow=0
    if pemlow+pemwidth>pemlen:
        pemlow=pemlen-pemwidth
    
    PL1['text']=f"Showing: {1+pemlow}-{pemlow+pemwidth} of {pemlen}"
    if regen:
        regenPeople()
def smrtComp():
    global emComb,Emrow
    changed=False
    for i,Email in enumerate(suEms[emlow:emlow+emwidth]):
        i+=emlow
        Rec,Sent,Subj=Email[0:3]
        Row=Emrow[i-emlow]
        for person in [Sent]+Rec:
            Person=getEmail(person)
            if Person!=Me:
                Val=emComb[i-emlow].get()
                if Val not in["","???"]:
                    ist=tentprojs.get(Val,[])
                    if Person not in ist:
                        ist.append(Person)
                        changed=True
                    tentprojs[Val]=ist
    regenPeople()
    regenProj()
    return(changed)
def smrtComps():
    tentprojs=deepcopy(projs)
    i=0
    while 1:
        i+=1
        L2['text']=f"Running... Iter: {i}"
        root.update()
        otp=deepcopy(tentprojs)
        x=smrtComp()
        if otp==tentprojs:
            break
        if not x:
            break
    L2['text']=f"Finished in {i} passes"
        
def panic(x,depth=""):
    num=0
    cds=x.winfo_children()
    if len(cds)>0:
        for cd in cds:
            num+=panic(cd,depth+'-'+x.winfo_class())
        print(depth+'-'+x.winfo_class(),num)
        return(num)
    else:
        return(1)
def updateSearch():
    ESparams['search']=E1.get()
    regenPeople()
def applyKword():
    for row in projects:
        vals=qsplit(row[2].get(),",")
        if vals!=[""]:
            vals=[val[val[0]==" ":]for val in vals]
            kword[row[0]['text']]=vals
        else:
            kword[row[0]['text']]=[" "*100]
    applyKwords()
##    regenEmails()
def applyKwords():
    global Emrow,emlow,knowns
##    print("Good label started")
    nc=[]
    for i,Email in enumerate(Ems):
        opts=list(projs.keys())
        ans=knowns.get(repr(Email),"")
        qr=0
        if ans=="":
            for opt in opts:
                for key in kword.get(opt,[]):
                    if key!="":
                        if key.lower() in Email[2].lower():
                            nc.append([i,opt])
##                            print(f"Line {i}")
##    print(nc)
    for i,pro in nc:
        Email=Ems[i]
        Rec,Sent,Subj=Email[0:3]
        for person in [Sent]+Rec:
            Person=getEmail(person)
            if Person!=Me:
                Val=pro
                knowns[repr(Email)]=Val
                if Val not in["","???"]:
                    ist=projs[Val]
                    if Person not in ist:
                        ist.append(Person)
                    projs[Val]=ist
##                    print(f"Line {i} modded project-{Val}")
    regenProj()
##    print("finished")
def regenHours():
    global hproj,tnset
    for proj in hproj:
        proj.destroy()
##        proj[0].destroy()
##        proj[1].destroy()
##        proj[2].destroy()
    for i,key in enumerate([x for x in sssproj if x!="Junk"]):
        hproj.append(ttk.Label(f1,text=key))
        hproj[-1].grid(column=2+2*i, row=0)
        hproj.append(ttk.Separator(f1, orient=VERTICAL))
        hproj[-1].grid(column=1+2*i, row=0, rowspan=1000, sticky="ns")
    projects=[]
    fday=max([int(x[0:2])*100+int(x[3:5]) for x in nset.keys()])
    day=datetime.datetime(2023,2,1)
    DATE=format(day,"%m-%d")
    vday=100*day.month+day.day
    row=1
    tnset={}
    while vday<=fday:
        row+=1
        nnset=nset.get(DATE,{})
        hproj.append(ttk.Label(f1,text=" "+DATE+" "))
        hproj[-1].grid(column=0, row=2*row, sticky="w")
        hproj.append(ttk.Separator(f1, orient=HORIZONTAL))
        hproj[-1].grid(column=0, row=2*row-1, columnspan=1000, sticky="we")
        for i,key in enumerate([x for x in sssproj if x!="Junk"]):
            nsk=nnset.get(key,None)
            if nsk!=None:
                hours=(nsk[0]*15+nsk[1]*5)/60
                tnset[key]=tnset.get(key,0)+hours
                hproj.append(ttk.Label(f1,text=round(hours,2)))
                hproj[-1].grid(column=2+2*i, row=2*row)
        
        day+=datetime.timedelta(1)
        DATE=format(day,"%m-%d")
        vday=100*day.month+day.day
    hproj.append(ttk.Label(f1SS,text="Total"))
    hproj[-1].grid(column=0, row=2, sticky="w")
    hproj.append(ttk.Label(f1,text="Total"))
    hproj[-1].grid(column=0, row=2, sticky="w")
    hproj.append(ttk.Separator(f1SS, orient=HORIZONTAL))
    hproj[-1].grid(column=0, row=1, columnspan=1000, sticky="we")
##    print(tnset)
    for i,key in enumerate([x for x in sssproj if x!="Junk"]):
        nsk=tnset.get(key,None)
        if nsk!=None:
            hproj.append(ttk.Label(f1SS,text=round(nsk,2)))
            hproj[-1].grid(column=2+2*i, row=2)
            hproj.append(ttk.Label(f1,text=round(nsk,2)))
            hproj[-1].grid(column=2+2*i, row=2)
    for i,key in enumerate([x for x in sssproj if x!="Junk"]):
        hproj.append(ttk.Label(f1SS,text=key))
        hproj[-1].grid(column=2+2*i, row=0)
        hproj.append(ttk.Separator(f1SS, orient=VERTICAL))
        hproj[-1].grid(column=1+2*i, row=0, rowspan=3, sticky="ns")
    hproj.append(ttk.Label(f1SS,text=" "*9))
    hproj[-1].grid(column=2+2*(i+1), row=0)
    hproj.append(ttk.Label(f1SS,text=" 04-04 ",foreground="#efefef"))
    hproj[-1].grid(column=0, row=0)
##    hproj[-1]['foreground']=hproj[-1].cget('background')
def export():
    txt=","
    for i,key in enumerate([x for x in sssproj if x!="Junk"]+["Subjects"]):
        txt+=key+","
    txt=txt[:-1]+"\n"
    txt+="Total,"
    for i,key in enumerate([x for x in sssproj if x!="Junk"]):
        nsk=tnset.get(key,0)
        txt+=str(round(nsk,2))+","
    txt=txt[:-1]+"\n"
    for i,key in enumerate([x for x in sssproj if x!="Junk"]+["Subjects"]):
        txt+=key+","
    txt=txt[:-1]+"\n"
    fday=max([int(x[0:2])*100+int(x[3:5]) for x in nset.keys()])
    day=datetime.datetime(2023,2,1)
    DATE=format(day,"%m-%d")
    vday=100*day.month+day.day
    row=1
    while vday<=fday:
        row+=1
        nnset=nset.get(DATE,{})
        txt+=DATE+","
        for i,key in enumerate([x for x in sssproj if x!="Junk"]):
            nsk=nnset.get(key,0)
            if nsk!=0:
                nsk=(nsk[0]*15+nsk[1]*5)/60
            txt+=str(round(nsk,2))+","
        txt=txt[:-1]+"\n"
        day+=datetime.timedelta(1)
        DATE=format(day,"%m-%d")
        vday=100*day.month+day.day
    with open("export_hours.csv","w") as f:
        f.write(txt)
    export2()
def export2():
    txt="Date,Sender,Recipients,Subject,Project\n"
    for i,Email in enumerate(Ems):
        opts=list(projs.keys())
        ans=knowns.get(repr(Email),"")
        qr=0
        if ans=="":
            if Me not in [getEmail(x)for x in Email[0]] and Me != getEmail(Email[1]):
                ans="Junk"
            for opt in opts:
                for key in kword.get(opt,[]):
                    if key!="":
                        if key.lower() in Email[2].lower():
                            ans=opt
        if ans=="":
            ans=groups.get(repr(set(Email[0]+[Email[1]])),"")
        opt1,att1=logic(Email,opts,projs)
        opt=opt1[:]
        if len(opt1)!=1:
            opt2,att2=logic(Email,opts,tentprojs)
            opt=opt2[:]
            if len(opt2)!=1:
                opt=opt1
        if ans!="":
            opt=[ans]
        if len(opt)==1:
            prem=(opt[0])
        elif len(opt)>1:
            prem=("???")
        elif (att1 and len(opt1)==0) and (att2 and len(opt2)==0):
            prem=("---")
        else:
            prem=("")
        txt+=cvsafe(format(Email[3],"%m-%d"))+","
        txt+=cvsafe(Email[1])+","
        txt+=cvsafe(", ".join(Email[0]))+","
        txt+=cvsafe(Email[2])+","
        txt+=prem+","
        txt=txt[:-1]+"\n"
    with open("export_emails.csv","w",encoding='utf-8') as f:
        f.write(txt)
def cvsafe(txt):
    txt=txt.replace('"','""')
    return('"'+txt+'"')
            
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

To=[]
Sent=[]
Me='' #REDACTED

root = Tk()
root.title("Hours Helper")
root.option_add('*tearOff', False)
##root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))

core = ttk.Frame(root, padding="8 8 8 8")
n = ttk.Notebook(core)#,height=900,width=1900)

f1SS = ttk.Frame(n)
f2S = ttk.Frame(n)
f3S = ttk.Frame(n)
f4SS = ttk.Frame(n)

f4S = ttk.Frame(f4SS)
AAA=ttk.Label(f4SS,text="\n"*56)
AAA.grid(row=1,column=1)
f4S.grid(row=1,column=0,sticky="nsew")
f4s=SF.VerticalScrolledFrame(f4S)
f4=f4s.interior
f4s.pack(side=LEFT, fill=BOTH, expand=TRUE)

f3s=SF.VerticalScrolledFrame(f3S)
f3=f3s.interior
f3s.pack(side=LEFT, fill=BOTH, expand=TRUE)

f2s=SF.VerticalScrolledFrame(f2S)
f2=f2s.interior
f2s.pack(side=LEFT, fill=BOTH, expand=TRUE)

f1S = ttk.Frame(f1SS)
AAA=ttk.Label(f1SS,text="\n"*56)
AAA.grid(row=3,column=51)
AAA=ttk.Label(f1SS,text=" "*300)
AAA.grid(row=4,column=0,columnspan=50)
f1S.grid(row=3,column=0,sticky="nswe",columnspan=50)
f1s=SF.VerticalScrolledFrame(f1S)
f1=f1s.interior
f1s.pack(side=LEFT, fill=BOTH, expand=TRUE)

##f1s=SF.VerticalScrolledFrame(f1S)
##f1=f1s.interior
##f1s.pack(side=LEFT, fill=BOTH, expand=TRUE)
##ttk.Label(f1,text="\n"*60).grid(column=0, row=0)

n.add(f1SS, text='Hours')
n.add(f2S, text='Projects')
n.add(f3S, text='People')
n.add(f4SS, text='Emails')

tabs = ttk.Frame(f2)

core.grid(column=0, row=0)
tabs.grid(column=0, row=1)
n.grid(column=0, row=0)

tframes=[]
projs={'Junk':[]}
tentprojs=deepcopy(projs)
Ems=[]
people=[]
projects=[]
emComb=[]
Emrow=[]
kword={}
emlow=0
emwidth=100
pemlow=0
pemwidth=200
ESparams={'search':""}
knowns={}
groups={}
hproj=[]
nset={}

if not loadEms():
    Ems=fetchEmails()
    fixEms()
    saveEms()
Ems=Ems[:]
print("Emails saved")
loadProj()
##tentprojs=deepcopy(projs)

emlen=len(Ems)

ssproj=ssproj=fsorted(['']+list(projs.keys()))
sssproj=[x for x in ssproj if x!=""]

uniEm(Ems)
labelData()

print("Projects imported")

#People page
ttk.Label(f3,text="Email").grid(column=0, row=0)
ttk.Label(f3,text="Projects").grid(column=1, row=0)
ttk.Label(f3,text="Search:").grid(column=2, row=0)
E1=ttk.Entry(f3)
E1.grid(column=3,row=0)
E1.bind('<Return>',lambda e: updateSearch())
ttk.Separator(f3, orient=HORIZONTAL).grid(column=0, row=1, columnspan=7, sticky="we")
ttk.Button(f3,text="<<",command=lambda:pnem(-1000)).grid(column=4, row=0,sticky="we")
ttk.Button(f3,text="<",command=lambda:pnem(-200)).grid(column=5, row=0,sticky="we")
PL1=ttk.Label(f3,text="Showing: 1-500\nOut of: 3000")
PL1.grid(column=6, row=0,sticky="")
ttk.Button(f3,text=">",command=lambda:pnem(200)).grid(column=7, row=0,sticky="we")
ttk.Button(f3,text=">>",command=lambda:pnem(1000)).grid(column=8, row=0,sticky="we")

pemlen=len(esorted(Sent+To))

##ttk.Button(f3,text="Debug",command=lambda: print(root.winfo_geometry())).grid(column=3,row=0)


regenPeople()

#Hours page

#Projects page
ttk.Label(f2,text="Project").grid(column=0, row=0)
ttk.Label(f2,text="Members").grid(column=2, row=0)
ttk.Label(f2,text="Comma seperated\nKeywords").grid(column=5, row=0)
ttk.Button(f2,text="Apply keywords",command=applyKword).grid(column=5, row=2)
Ent=ttk.Entry(f2)
Ent.grid(column=0, row=2)
ttk.Button(f2,text="Add",command=addProj).grid(column=2, row=2)
ttk.Button(f2,text="Export to CSV",command=export).grid(column=4, row=2)
ttk.Button(f2,text="Save",command=saveProj).grid(column=4, row=0)
ttk.Separator(f2, orient=HORIZONTAL).grid(column=0, row=1, columnspan=6, sticky="we")
ttk.Separator(f2, orient=HORIZONTAL).grid(column=0, row=3, columnspan=6, sticky="we")
ttk.Separator(f2, orient=VERTICAL).grid(column=1, row=0, rowspan=1000, sticky="ns")
ttk.Separator(f2, orient=VERTICAL).grid(column=3, row=0, rowspan=3, sticky="ns")

#Emails Page
AA=ttk.Label(f4SS,text=" "*620)
AA.grid(row=0,column=0)
ttk.Label(f4,text="Sender").grid(column=0, row=1)
ttk.Separator(f4, orient=VERTICAL).grid(column=1, row=1, rowspan=1000, sticky="ns")
ttk.Label(f4,text="Recipients").grid(column=2, row=1)
ttk.Separator(f4, orient=VERTICAL).grid(column=3, row=1, rowspan=1000, sticky="ns")
ttk.Label(f4,text="Subject").grid(column=4, row=1)
ttk.Separator(f4, orient=VERTICAL).grid(column=5, row=1, rowspan=1000, sticky="ns")
ttk.Label(f4,text="Project").grid(column=6, row=1)
ttk.Label(f4,text="Inferred").grid(column=7, row=1)
ttk.Separator(f4, orient=HORIZONTAL).grid(column=0, row=2, columnspan=8, sticky="we")
##ttk.Button(f4,text=">>",command=lambda:nem(500)).grid(column=0, row=0,sticky="nw")

ssf4=ttk.Frame(f4SS)
ssf4.grid(column=0, row=0,columnspan=100,sticky="nw")
ttk.Button(ssf4,text="<<",command=lambda:nem(-500)).grid(column=0, row=0,sticky="we")
ttk.Button(ssf4,text="<",command=lambda:nem(-100)).grid(column=1, row=0,sticky="we")
L1=ttk.Label(ssf4,text="Showing: 1-100\nOut of: 500")
L1.grid(column=2, row=0,sticky="")
ttk.Button(ssf4,text=">",command=lambda:nem(100)).grid(column=3, row=0,sticky="we")
ttk.Button(ssf4,text=">>",command=lambda:nem(500)).grid(column=4, row=0,sticky="we")
ttk.Button(ssf4,text="Search",command=lambda:nem(0)).grid(column=9, row=0,sticky="we")
ttk.Button(ssf4,text="Smart Fill",command=lambda:smrtComps()).grid(column=5, row=0,sticky="we")
L2=ttk.Label(ssf4,text="Finished in 0 passes")
L2.grid(column=6, row=0,sticky="we")
ttk.Button(ssf4,text="Debug",command=lambda:panic(root)).grid(column=7, row=0,sticky="we")
L3=ttk.Label(ssf4,text="100% Blank\n0% Partial")
L3.grid(column=18, row=0,sticky="we")
ttk.Label(ssf4,text="From:").grid(column=10, row=0,sticky="e")
E41=ttk.Entry(ssf4)
E41.grid(column=11, row=0,sticky="we")
ttk.Label(ssf4,text="To:").grid(column=12, row=0,sticky="e")
E42=ttk.Entry(ssf4)
E42.grid(column=13, row=0,sticky="we")
ttk.Label(ssf4,text="Mentions:").grid(column=14, row=0,sticky="e")
E43=ttk.Entry(ssf4)
E43.grid(column=15, row=0,sticky="we")
ttk.Label(ssf4,text="Includes:").grid(column=16, row=0,sticky="e")
E44=ttk.Entry(ssf4)
E44.grid(column=17, row=0,sticky="we")
C1=ttk.Checkbutton(ssf4,text="Show Blanks") 
C1.grid(column=19, row=0,sticky="w")
C1.invoke()
C2=ttk.Checkbutton(ssf4,text="Show Ambiguous") 
C2.grid(column=20, row=0,sticky="w")
C2.invoke()
C3=ttk.Checkbutton(ssf4,text="Show Conflicts") 
C3.grid(column=21, row=0,sticky="w")
C3.invoke()
C4=ttk.Checkbutton(ssf4,text="Show Filled") 
C4.grid(column=22, row=0,sticky="w")
C4.invoke()
ttk.Label(ssf4,text="Sort by:").grid(row=1,column=0)
CB1=ttk.Combobox(ssf4,width=8)
CB1['values'] = ["Date","Sender"]
CB1.state(['readonly'])
CB1.set("Date")
CB1.grid(row=1,column=1)

ttk.Separator(ssf4, orient=HORIZONTAL).grid(column=0, row=2, columnspan=100, sticky="we")


##nem(0)
regenProj()
regenEmails()
applyKwords()
regenHours()
pnem(0)

input("Press Enter to Exit")
root.mainloop()
