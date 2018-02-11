from math import cos,sin,atan2,radians,asin,sqrt,pi
from tkinter import *
import tkinter.filedialog,glob,os
import xlrd
import dateutil.parser
import pygmaps1,webbrowser
def haversine(lon2, lat2, lon1, lat1):
    # convert decimal degrees to radians 
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])

    # haversine formula 
    dlon = lon2 - lon1 
    dlat = lat2 - lat1 
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a)) 
    r = 6371 # Radius of earth in kilometers. Use 3956 for miles
    return c * r

def bearing(lon2,lat2,lon1,lat1):
    lat1,lon1,lat2,lon2=map(radians,(lat1,lon1,lat2,lon2))
    delta=lon2-lon1
    x=cos(lat2)*sin(delta)
    y=cos(lat1)*sin(lat2)-sin(lat1)*cos(lat2)*cos(delta)
    bearing=(atan2(x,y))*180/pi
    if bearing<0:
        bearing=360+bearing
    return bearing

def reverse(angle):
    if angle<180:
        angle=angle+180
    else:
        angle=angle-180
    return angle

def fmt_angle(angle):
    if angle<0:
        angle*=-1
    if angle>180:
        angle=360-angle
    return angle

def filter_(latlon,dates,an,noises):
    i=0
    while latlon[i]==latlon[i+1]:
        i+=1
    ref=bearing(latlon[i+1][0],latlon[i+1][1],latlon[i][0],latlon[i][1])
    doubt=False
    num=0
    paths=[[],[]]
    noise_angle=None
    speed={}
    dist={}
    div_b=[]
    div_s=[]
    bear_check=int(app.ent1.get())
    if app.is_speed_check.get():
        speed_check=0.01*int(app.ent2.get())
    noises.append('S No.\t\tData Points\t\t\tDate\t\tAN1\n')
    while(i<len(latlon)-10):
        local=bearing(latlon[i+1][0],latlon[i+1][1],latlon[i][0],latlon[i][1])
        if local==0:
            i+=1
            continue
        angle=local-ref
        angle=fmt_angle(angle)
        if doubt:
            m = i+5
            while m>=i-15 and m>=0:
                try:
                    spd=speed[m]
                except KeyError:
                    try:
                        dis=haversine(latlon[m+1][0],latlon[m+1][1],latlon[m][0],latlon[m][1])
                        dist[m]=dis
                        time1=dateutil.parser.parse(dates[m]).timestamp()
                        time2=dateutil.parser.parse(dates[m+1]).timestamp()
                        time=(time2-time1)/(60*60)
                        spd=dis/time
                        speed[m]=spd
                    except IndexError:
                        m-=1
                        continue
                m-=1
            if app.is_speed_check.get():
                if i>15:
                    av_speed=[speed[i] for i in range(m+1,i-1)]
                    av_speed.sort()
                    av_speed=sum(av_speed[2:])/(len(av_speed)-2)
                    if av_speed<30:
                        jump=av_speed*(speed_check+0.01)
                    else:
                        jump=av_speed*speed_check
                    step=av_speed*0.2
                else:
                    jump=10
                    step=1
            j1=i-2
            while speed[j1]==0:
                j1=j1-1
            temp1=bearing(latlon[i-1][0],latlon[i-1][1],latlon[j1][0],latlon[j1][1])
            j=i+1
            while speed[j]==0 and j<i+4:
                j=j+1
            temp2=bearing(latlon[j][0],latlon[j][1],latlon[i-1][0],latlon[i-1][1])
            angle2=temp2-temp1
            angle2=fmt_angle(angle2)
            if app.is_bearing_check.get():
                if noise_angle>110:
                    condition1=angle>noise_angle-5 and angle2<40 and dist[i]>=dist[i-1]>=dist[j1]
                elif noise_angle<60:
                    condition1=angle>180-noise_angle and angle2<20 and dist[i-1]>dist[j1] and dist[i]>dist[j1]
                condition=condition1
                if condition:
                    div_b.append((round(noise_angle),round(angle)))
            if app.is_speed_check.get():
                condition2=speed[i]>speed[i-1]>speed[j1]+jump and speed[i+1]<speed[j1]+step
                condition=condition2
                if condition:
                    div_s.append((round(speed[j1]),round(speed[i-1])))
                    
            if app.is_speed_check.get() and app.is_bearing_check.get():
                condition=condition1 and condition2
                if condition:
                    div_b.append((round(noise_angle),round(angle)))
                    div_s.append((round(speed[i-1]),round(speed[i])))
            if not(app.is_speed_check.get() or app.is_bearing_check.get()):
                return
            if condition:
                path=[]
                dat=[]
                num+=1
                global n_vcr
                noise='noise-'+str(n_vcr)+'.'+str(num)+' : '+str(latlon[i][0])+" "+str(latlon[i][1])
                noises.append(noise+'\t'+dates[i]+'\t'+an[i]+'\n')
                print(noise,'\t',dates[i],'\t',an[i],end=" ")
                print([round(speed[m],2) for m in range(i-3,i+3)],end="\t")
                print([round(dist[m]*1000) for m in range(i-3,i+3)])
                k=i-4
                while k<i and k>0:
                    path.append((latlon[k][1],latlon[k][0]))
                    dat.append((dates[k],an[k]))
                    k+=1
                path.append((latlon[i][1],latlon[i][0]))
                dat.append((dates[i],an[i]))
                k=i+1
                while k<i+4 and k<len(latlon):
                    path.append((latlon[k][1],latlon[k][0]))
                    dat.append((dates[k],an[k]))
                    k+=1
                paths[0].append(path)
                paths[1].append(dat)
                ref=temp2
                i+=1 
            else:
                ref=local
                i+=1
            noise_angle=None
            doubt=False
            continue
        if angle>bear_check:
            if angle>110 or angle<60:
                doubt=True
                noise_angle=angle
            else:
                doubt=False
        else:
            doubt=False
        ref=local
        i+=1
    print('Total noises is',len(paths[0]))
    noises.append('\n\n')
    return paths,div_b,div_s

def plot(data,paths,vehicle,div_b,div_s):
    map_plot = pygmaps1.maps(data[0],data[1],16,vehicle,div_b,div_s)
    k1=str(data[0])
    for i in range(len(paths[0])):
        map_plot.addpath(paths[0][i],paths[1][i],"#000000")
    pwd=os.getcwd()
    name=pwd+'\\map_plot.draw'+str(k1)+'.html'
    map_plot.draw(name)
    webbrowser.open_new_tab(name)
    return

def load_data(wb,start_date,stop_date,sheets):
    for j in sheets:
        global n_vcr
        n_vcr+=1
        noise=[]
        noise.append(j+'\n')
        sheet=wb.sheet_by_name(j)
        i=1
        latlon=[]
        dates=[]
        an=[]
        is_read=False
        while True:
            try:
                row=sheet.row_values(i)
                date=str(row[3])
                if date[:10]>=start_date:
                    is_read=True
                if date[:10]>=stop_date:
                    is_read=False
                if row[16]!='13' and row[16]!='94' and is_read:
                    latlon.append((float(row[0]),float(row[1])))
                    dates.append(date)
                    an.append(row[13])
                i=i+1
            except IndexError:
                break
        if not latlon:
            Label(app,text='Interval not found',foreground='red').grid(row=8,column=0,columnspan=2,sticky=W)
            return
        try:
            paths,b_noise,s_noise=filter_(latlon,dates,an,noise)
        except ValueError:
            Label(app,text='Missing values\t',foreground='red').grid(row=8,column=0,columnspan=2,sticky=W)
            return
        results[n_vcr]=(paths,j,b_noise,s_noise)
        noises.append(noise)
    return

class App2(Frame):
    def __init__(self,master):
        super(App2,self).__init__(master)
        self.grid(row=0,column=0,rowspan=4,columnspan=10,sticky=W)
        self.set_widget()
        self.master=master
        self.is_read=True
        
    def set_widget(self):
        Label(self,text="Directory").grid(row=0,column=0,sticky=W)
        self.txt1=Text(self,width=50,height=1,wrap=CHAR)
        self.txt1.grid(row=0,column=1,columnspan=5,sticky=W)
        self.var=StringVar(self)
        self.var.set(None)
        Radiobutton(self,text='Browse Folder',
                    variable=self.var,
                    value='Folder',
                    command=self.get_filename
                    ).grid(row=0,column=7,sticky=W)
        Radiobutton(self,text='Browse File',
                    variable=self.var,
                    value='File',
                    command=self.get_filename
                    ).grid(row=0,column=8,sticky=W)
    def get_filename(self):
        if self.var.get()=='File':
            fname=tkinter.filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"),
                                           ("HTML files", "*.html;*.htm") ))
            bt_name='Get_IDs'
            self.text(fname,bt_name)
        else:
            pname = tkinter.filedialog.askdirectory()
            path=pname+'/*.xlsx'
            self.files=glob.glob(path)
            bt_name='Parse Files'
            self.text(pname,bt_name)
    def text(self,name1,name2):
        if hasattr(self,'btn'):
            self.btn.destroy()
        self.btn=Button(self,text=name2,command=self.run_id)
        self.btn.grid(row=2,column=0,sticky=W)
        if name1:
            if not self.is_read:
                app.close_all()
                app.destroy()
            self.txt1.delete(0.0,END)
            self.txt1.insert(0.0,name1)
            self.is_read=True
    def run_id(self):
        if self.is_read:
            global app
            app=App(self.master)
class App(Frame):
    def __init__(self,master):
        super(App,self).__init__(master)
        self.grid(row=6,column=0,rowspan=6,columnspan=10,sticky=W)
        self.master=master
        self.months=[(i+1) for i in range(12)]
        self.months=['Months']+self.months
        self.days=[(i+1) for i in range(31)]
        self.days=['Days']+self.days
        self.year=[i for i in range(2016,2030)]
        self.year=['Year']+self.year
        self.type=app2.var.get()
        self.set_widget()
    def open_file(self,name):
        wb=xlrd.open_workbook(name)
        sheets=wb.sheet_names()
        return wb,sheets
    def close_all(self):
        for i in ('app1','root','app4'):
            try:
                atr=self.__dict__[i]
                atr.destroy()
            except KeyError:
                continue
            except TclError:
                continue
        return
    def set_widget(self):
        self.sheets={}
        if self.type=='File':
            fname=app2.txt1.get(0.0,END)
            fname=fname[:-1]
            self.wb,sheets=self.open_file(fname)
            sheets=['Vehicle ID']+sheets
            self.var1=StringVar(self)
            self.var1.set('Vehicle ID')
            OptionMenu(self,self.var1,*sheets).grid(row=2,column=0,columnspan=2,sticky=W)
        else:
            for i in app2.files:
                wb,sheets=self.open_file(i)
                self.sheets[wb]=sheets
        app2.is_read=False
        Label(self,text='Start Date').grid(row=3,column=0,sticky=W)
        self.var2=StringVar(self)
        self.var2.set('Days')
        self.var3=StringVar(self)
        self.var3.set('Months')
        self.var4=StringVar(self)
        self.var4.set('Year')
        OptionMenu(self,self.var2,*self.days).grid(row=3,column=1,sticky=W)
        OptionMenu(self,self.var3,*self.months).grid(row=3,column=2,sticky=W)
        OptionMenu(self,self.var4,*self.year).grid(row=3,column=3,sticky=W)
        Label(self,text='Stop Date').grid(row=4,column=0,sticky=W)
        self.var5=StringVar(self)
        self.var5.set('Days')
        self.var6=StringVar(self)
        self.var6.set('Months')
        self.var7=StringVar(self)
        self.var7.set('Year')
        OptionMenu(self,self.var5,*self.days).grid(row=4,column=1,sticky=W)
        OptionMenu(self,self.var6,*self.months).grid(row=4,column=2,sticky=W)
        OptionMenu(self,self.var7,*self.year).grid(row=4,column=3,sticky=W)
        self.is_bearing_check=BooleanVar()
        self.is_bearing_check.set(True)
        Label(self,text='Bearing').grid(row=5,column=0,sticky=W)
        self.ent1=Entry(self,width=10)
        self.ent1.grid(row=5,column=1,sticky=W)
        self.ent1.insert(0,'110')
        Label(self,text='Speed %').grid(row=5,column=2,sticky=W)
        self.ent2=Entry(self,width=10)
        self.ent2.grid(row=5,column=3,sticky=W)
        self.ent2.insert(0,'30')
        Checkbutton(self,text="Include bearing check",variable=self.is_bearing_check
                    ).grid(row=6,column=0,columnspan=2,sticky=W)
        self.is_speed_check=BooleanVar()
        self.is_speed_check.set(True)
        Checkbutton(self,text="Include speed check",variable=self.is_speed_check
                    ).grid(row=6,column=2,columnspan=2,sticky=W)
        Button(self,text='Analyse',command=self.analyse).grid(row=7,column=0,sticky=W)
        self.haver=BooleanVar()
        Checkbutton(self,text='Show Haversine',
                    variable=self.haver,
                    command=self.run,
                    ).grid(row=7,column=3,columnspan=2,sticky=W)
    def run(self):
        if self.haver.get():
            self.app1=App1(self.master)
        else:
            self.app1.destroy()
            
    def get_interval(self):
        start_day=str(self.var2.get())
        if len(start_day)==1:
            start_day='0'+start_day
        start_month=str(self.var3.get())
        if len(start_month)==1:
            start_month='0'+start_month
        stop_day=str(self.var5.get())
        if len(stop_day)==1:
            stop_day='0'+stop_day
        stop_month=str(self.var6.get())
        if len(stop_month)==1:
            stop_month='0'+stop_month
        start_date=str(self.var4.get())+'-'+start_month+'-'+start_day
        stop_date=str(self.var7.get())+'-'+stop_month+'-'+stop_day
        return start_date,stop_date
    
    def analyse(self):
        global results
        global noises
        global n_vcr
        n_vcr=0
        results={}
        noises=[]
        self.close_all()
        Label(self,text='\t\t\t\t').grid(row=8,column=0,columnspan=3,sticky=W)
        self.start_date,self.stop_date=self.get_interval()
        if self.sheets:
            for i in self.sheets.keys():
                self.vehicle=self.sheets[i]
                print(self.vehicle)
                load_data(i,self.start_date,self.stop_date,self.vehicle)
        else:
            self.vehicle=str(self.var1.get())
            if self.vehicle=='Vehicle ID':
                Label(self,text='Give vehicle ID\t',foreground='red').grid(row=8,column=0,columnspan=2,sticky=W)
                return
            load_data(self.wb,self.start_date,self.stop_date,[self.vehicle])
        if noises:
            self.show_result()
    def show_result(self):
        self.app4=Frame(self.master)
        self.app4.grid(row=7,column=8,sticky=W)
        Label(self.app4,text='Analysis Complete!').grid(row=1,column=1,rowspan=2,columnspan=8,sticky=S)
        Button(self.app4,text='Show Results',command=self.open_win).grid(row=3,column=1,rowspan=2,columnspan=8,sticky=N)
        for i in range(10):
            Label(self.app4,text='_').grid(row=0,column=i,sticky=S)
            Label(self.app4,text='_').grid(row=5,column=i,sticky=S)
        for i in range(6):
            Label(self.app4,text='_').grid(row=i,column=0,sticky=S)
            Label(self.app4,text='_').grid(row=i,column=9,sticky=S)
    def open_win(self):
        try:
            self.root.state()
        except AttributeError:
            self.show_noises()
        except TclError:
            self.show_noises()
    def show_noises(self):
        self.root=Tk()
        self.root.title('Result')
        self.root.geometry('510x360')
        self.root.resizable(0,0)
        app3=Frame(self.root)
        app3.grid()
        app3.txt=Text(app3,width=60,height=20,borderwidth=3,wrap=WORD)
        app3.txt.grid(row=0,column=0,rowspan=5,columnspan=5,sticky=W)
        scrollb = Scrollbar(app3, command=app3.txt.yview)
        scrollb.grid(row=0, column=5,rowspan=5, sticky='nsew')
        app3.txt['yscrollcommand'] = scrollb.set
        app3.txt.tag_config("tag",foreground="blue",underline=True)
        app3.txt.tag_bind('tag','<Button-1>',self.call_event)
        for i in noises:
            for j in i:
                if 'noise' in j:
                    app3.txt.insert(END,j[0:9],'tag')
                    app3.txt.insert(END,j[9:])
                else:
                    if len(j)<=16:
                        self.cur_vcr=j[:-1]
                    app3.txt.insert(END,j)
        Button(app3,text='Show all on Map',command=self.show_map_full).grid(row=5,column=0,sticky=W)
        self.root.mainloop()
    def call_event(self,event):
        index=event.widget.index('@%s,%s' %(event.x,event.y))
        tag_index=event.widget.tag_ranges('tag')
        for start,end in zip(tag_index[0::2],tag_index[1::2]):
            if event.widget.compare(start,'<=',index) and event.widget.compare(index,'<',end):
                noise=event.widget.get(start,end)
                noise=noise.split('-')
                noise=noise[1].split('.')
                self.show_map(noise)
    def show_map(self,num):
        result=results[int(num[0])]
        vcl=result[1]
        path1=result[0][0][int(num[1])-1]
        dat1=result[0][1][int(num[1])-1]
        try:
            div_b=result[2][int(num[1])-1]
        except IndexError:
            div_b=None
        try:
            div_s=result[3][int(num[1])-1]
        except IndexError:
            div_s=None
        l=len(path1)
        data=path1[l//2]
        plot(data,[[path1],[dat1]],vcl,div_b,div_s)
    def show_map_full(self):
        for i in results.values():
            l1=len(i[0][0])
            l2=len(i[0][0][l1//2])
            data=i[0][0][l1//2][l2//2]
            plot(data,i[0],i[1])
class App1(Frame):
    def __init__(self,master):
        super(App1,self).__init__(master)
        self.grid(row=12,column=0,rowspan=5,columnspan=10,sticky=W)
        self.master=master
        self.set_widget()
    def set_widget(self):
        Label(self,text='Latitude').grid(row=0,column=0,sticky=W)
        Label(self,text='Longitude').grid(row=0,column=1,sticky=W)
        self.ent1=Entry(self)
        self.ent1.grid(row=1,column=0,sticky=W)
        self.ent2=Entry(self)
        self.ent2.grid(row=1,column=1,sticky=W)
        self.ent3=Entry(self)
        self.ent3.grid(row=2,column=0,sticky=W)
        self.ent4=Entry(self)
        self.ent4.grid(row=2,column=1,sticky=W)
        self.btn=Button(self,text='Distance',command=self.calculate)
        self.btn.grid(row=3,column=0,sticky=W)
    def calculate(self):
        try:
            if hasattr(self,'lbl'):
                self.lbl.destroy()
            lat1=float(self.ent1.get())
            lon1=float(self.ent2.get())
            lat2=float(self.ent3.get())
            lon2=float(self.ent4.get())
            dis=haversine(lon2,lat2,lon1,lat1)
            self.btn.destroy()
            self.ent5=Entry(self)
            self.ent5.grid(row=3,column=0,sticky=N)
            self.ent5.delete(0,END)
            self.ent5.insert(0,str(dis))
        except ValueError:
            self.lbl=Label(self,text='Empty Lat/Lon',foreground='red')
            self.lbl.grid(row=4,column=0,sticky=W)
root=Tk()
root.title('ANALYSIS')
root.geometry('700x350')
root.resizable(0,0)
app2=App2(root)
root.mainloop()

















