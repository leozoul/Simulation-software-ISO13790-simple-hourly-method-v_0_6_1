# -*- coding: utf-8 -*-
"""
Created on Mon Jun 20 12:52:10 2022

@author: Leonidas Zouloumis
\nDescription: This file contains the functions necessary for the indoor air
calculation, according to EN ISO 13790, simple hourly method
"""

import pandas as pd
import numpy as np
from math import log as ln, cos, sin, pi, acos
from matplotlib import pyplot as plt
import matplotlib.dates as mdates
from TimeLib import TimeList, GetMonthtxt
from DEYAK_primary_secondary4_final_optimize_sched import create_sched

class Building():
    
    def __init__(
            self, 
            clim_file='Athens.xlsx', 
            phi = 37.98381,
            bld_type='Light', 
            bld_tinit=18,
            time_interval=60,
        ):
        self.dt = time_interval
        self.OpaqueList=[]
        self.FrameList=[]
        self.FloorArea = 0
        self.Hop = 0
        self.OpArea = 0
        self.Hw = 0
        self.Hve = 0
        self.VdotList_handvalues = [0]*24
        self.WinArea = 0
        self.Um = 0
        self.Type = bld_type
        self.SetBldType()
        self.Qhvac = 0
        self.Qhvac_temp = 0
        self.Temp_iv = bld_tinit
        self.tm1 = self.Temp_iv
        self.Tairprev = bld_tinit
        self.Tsprev = bld_tinit
        self.Topprev = 0.3*self.Tairprev+ 0.7*self.Tsprev
        self.ExternalPower = [None]
        self.Latitude = 37.98381
        # ---- Climate Data
        self.fl = clim_file
        # ---- Simulation clock
        self.step = 0 #it is  like a list accessor ([0, 24*3600] with timestep dt)
        self.HrClock = 0
        self.Nom_consumer = None #In case of linking it with a consumer this 
        # is the value that is taken into account during calculations, 
        # used with RunSim and Update
        # ---- Data Lists
        self.TairList = []
        self.TsList = []
        self.TmList = []
        self.TopList = []
        self.TopListRAW = []
        self.TheList = []
        self.ToutList = []
        self.IsolmeasList = None
        self.ToutmeasList = None
        self.ThsupmeasList = None
        self.FsolList = []
        self.FintList = []
        self.FhlossList = []
        self.FmechvelossList = [] # not currently used by simulations
        self.FnatvelossList = [] # uses either air supply hand values, or calculates by window opening
        self.QhvacList = []
        self.TimeConstant = 1
        self.TClist = []
        
        
    def AddElement(
            self, 
            attr=' ', 
            l=1, 
            h=1, 
            k=80000, 
            o='None', 
            f=1, 
            a=0.6, 
            ul=0.5, 
            r=0.04, 
            g=0.3, 
            ffr=0.2,
            ven_zones=None
            ):
        if attr == 'Window' or attr == 'Door':
            new = Frame(length=l, height=h, kj=k, ori=o, fsh=f, ggl=g, u=ul, ff=ffr, nvz=ven_zones)
            self.FrameList.append(new)
        else:
            new = Opaque(length=l, height=h, kj=k, ori=o, fsh=f, abso=a, u=ul, res=r)
            if attr=='Floor':
                self.FloorArea += new.Area
                new.Fsh = 0
            self.OpaqueList.append(new)
        new.Attr = attr
    
    def AssignFrametoOpaque(self, win_indx, opq_indx):
        self.OpaqueList[opq_indx].AttachedFrames.append(self.FrameList[win_indx])
        self.OpaqueList[opq_indx].Area -= self.FrameList[win_indx].Area
        self.OpaqueList[opq_indx].AttachedFrames_indx.append(win_indx)
        
    def RemoveFramefromOpaque(self, win_indx, opq_indx):
        self.OpaqueList[opq_indx].Area += self.FrameList[win_indx].Area
        self.OpaqueList[opq_indx].AttachedFrames.remove(self.FrameList[win_indx])
        self.OpaqueList[opq_indx].AttachedFrames_indx.remove(win_indx)
    
    def ResetOpaqueAttachedList(self, opq_indx):
        self.OpaqueList[opq_indx].Area = self.OpaqueList[opq_indx].Length * self.OpaqueList[opq_indx].Height
        self.OpaqueList[opq_indx].AttachedFrames = []
        self.OpaqueList[opq_indx].AttachedFrames_indx = []
    
    def Convert_timestep(self, dt2, whateverthatis):
        tinterp = []
        thourly = []
        for i in range(0, 86400, self.dt):
            tinterp.append(i)
        for i in range(0, 90000, dt2):
            thourly.append(i)
        final_array = np.interp(tinterp, thourly, list(whateverthatis))
        return list(final_array)
    
    def Convert_timestep_24h(self, dt2, whateverthatis):
        tinterp = []
        thourly = []
        for i in range(0, 86400, self.dt):
            tinterp.append(i)
        for i in range(0, 86400, dt2):
            thourly.append(i)
        final_array = np.interp(tinterp, thourly, list(whateverthatis))
        return list(final_array)
    
    def Construct_scheds(self, sched_array=[0, 0], value_array=[0], off_value=0):
        DailySchedule = list()
        for i in range(0, 86400, self.dt):
            DailySchedule.append(off_value)
        for i in range(len(sched_array)):
            for j in range(len(DailySchedule)):
                if j>=(sched_array[i][0]/self.dt) and j<(sched_array[i][1]/self.dt):
                    DailySchedule[j] = value_array[i]
        return DailySchedule
            
    def DoYCalc(self, month, day):
        month_days = [
            31,
            28,
            31,
            30,
            31,
            30,
            31,
            31,
            30,
            31,
            30,
            31
            ]
        n = day
        cnt = 0
        while cnt < len(month_days) and cnt < month - 1:
            n += month_days[cnt]
            cnt += 1
        return n
    
    def FsolCalc(self):
        Fsol = 0
        for x in self.OpaqueList:
            Isol_value = self.Isol[x.Orientation_num][self.step]
            Fsol += x.Fsh*x.Asol*Isol_value
        for x in self.FrameList:
            Isol_value = self.Isol[x.Orientation_num][self.step]
            Fsol += x.Fsh*x.Asol*Isol_value
        return Fsol
    
    def FintCalc(self): #Needs to be fixed
        Fint = self.People[self.step]*60
        return Fint
        
    def SetBldType(self):
        if self.Type == 'Very light':
            self.Ccoef = 80000
        elif self.Type == 'Light':
            self.Ccoef = 110000
        elif self.Type == 'Medium':
            self.Ccoef = 165000
        elif self.Type == 'Heavy':
            self.Ccoef = 260000
        elif self.Type == 'Very heavy':
            self.Ccoef = 370000
        else:
            self.Ccoef = 0
            self.Type = 'Invalid'
        
    def GetIsolList(self, fl):
        df = pd.read_excel(fl)
        c = df.values
        a = np.transpose(c)
        Gn = a[11][:]
        Gs = a[8][:]
        Ge = a[9][:]
        Gw = a[10][:]
        Gh = a[5][:]
        month = a[1][:]
        day = a[2][:]
        return Gh, Gn, Ge, Gs, Gw, month, day
    
    def GetTempOutList(self, fl):
        df = pd.read_excel(fl)
        c = df.values
        a = np.transpose(c)
        Temp = a[4][:]
        month = a[1][:]
        day = a[2][:]
        return Temp, month, day
    
    def GetOutdoorTemp(self, m=1, d=1):
        if self.ToutmeasList == None:
            Temp2, Month, Day = self.GetTempOutList(self.fl)
            cnt = 0
            first = 0
            last = 0
            found_first = 0
            found_last = 0
            while found_last == 0 and cnt<len(Month):
                if Month[cnt] == m and Day[cnt] == d:
                     if found_first == 0:
                         first = cnt
                         found_first = 1
                     if cnt == len(Temp2)-1:
                         last = cnt
                         found_last = 1
                elif  cnt > first and found_first == 1:
                     last = cnt - 1
                     found_last = 1
                cnt += 1
            Temp = Temp2[first : last+2]
            if  len(Temp) == 24:
                Temp = np.append(Temp, Temp2[0])
            self.ToutmeasList = self.Convert_timestep(3600, Temp)
            self.ThsupmeasList = self.Convert_timestep(3600, Temp)
        elif self.ThsupmeasList == None:
            self.ThsupmeasList = self.Convert_timestep(
                90000//len(self.ToutmeasList), 
                self.ToutmeasList
                )
            self.ToutmeasList = self.Convert_timestep(
                90000//len(self.ToutmeasList), 
                self.ToutmeasList
                )
        else:
            self.ThsupmeasList = self.Convert_timestep_24h(
                86400//len(self.ToutmeasList), 
                self.ThsupmeasList
                )
            self.ToutmeasList = self.Convert_timestep_24h(
                86400//len(self.ToutmeasList), 
                self.ToutmeasList
                )
            
        self.TheList = self.ToutmeasList[:]
        self.ThsupList = self.ThsupmeasList[:]
        return self.TheList
       
    def GetIsol(self, m=1, d=1, b=90):
        if self.IsolmeasList == None:
            Gh2, Gn2, Ge2, Gs2, Gw2, Month, Day = self.GetIsolList(self.fl)
            cnt = 0
            first = 0
            last = 0
            found_first = 0
            found_last = 0
            while found_last == 0 and cnt<len(Month):
                if Month[cnt] == m and Day[cnt] == d:
                     if found_first == 0:
                         first = cnt
                         found_first = 1
                     if cnt == len(Gn2)-1:
                         last = cnt
                         found_last = 1
                elif  cnt > first and found_first == 1:
                     last = cnt - 1
                     found_last = 1
                cnt += 1
            Gh = Gh2[first : last+2]
            Gn = Gn2[first : last+2]
            Ge = Ge2[first : last+2]
            Gs = Gs2[first : last+2]
            Gw = Gw2[first : last+2]
            if  len(Gh) == 24:
                Gh = np.append(Gh, Gh2[0])
                Gn = np.append(Gn, Gn2[0])
                Ge = np.append(Ge, Ge2[0])
                Gs = np.append(Gs, Gs2[0])
                Gw = np.append(Gw, Gw2[0])
            Gh = self.Convert_timestep(3600, Gh)
            Gn = self.Convert_timestep(3600, Gn)
            Ge = self.Convert_timestep(3600, Ge)
            Gs = self.Convert_timestep(3600, Gs)
            Gw = self.Convert_timestep(3600, Gw)
        else:
            Gh, Gs, Ge, Gw, Gn = self.GetIsolmeasTiltedList(
                month=m, 
                day=d,   
                b_tilt=b,
                dt=self.dt
                )
        return Gh, Gs, Ge, Gw, Gn
    
    def GetIsolmeasTiltedList(
            self,
            month,
            day,
            b_tilt, #tilt angle, in degrees
            dt, # timestep, in seconds (for hourly measurements)
            ):
        f = self.Latitude * pi / 180
        Gsc = 1367
        G_list = self.IsolmeasList #W/m2
        gamma_list = [0, -pi/2, pi/2, pi] # south east west north
        n = self.DoYCalc(month, day)
        B = (n-1) * 360 / 365
        d = (
            0.006918 
            - 0.399912 * cos(B)
            + 0.070257 * sin(B)
            - 0.006758 * cos(2*B) 
            + 0.000907 * sin(2*B) 
            - 0.002697 * cos(3*B)
            + 0.001480 * sin(3*B)
            )
        w1_list = []
        w2_list = []
        w1_list.append(23*15-180)
        w2_list.append(0*15+180)
        for i in range(23):
            w1_list.append(i*15-180)
            w2_list.append((i+1)*15-180)
        Gt_list = [[], [], [], [], []]
        for i in range(len(G_list)):
            w1 = w1_list[i//int(3600/dt)] * pi / 180
            w2 = w2_list[i//int(3600/dt)] * pi / 180
            G = G_list[i]
            I = G * 3600 
            Io = (
                12 * 3600 / pi * Gsc * (1+0.033*cos(360/365*n)) 
                * (
                    cos(f) * cos(d) * (sin(w2)-sin(w1)) 
                    + (w2-w1) * sin(f) * sin(d)
                    )
                )
            Ghor = 0
            kt = I/Io
            if kt >= 0: # Io > 0 signifies the sun has risen
                if kt > 1:
                    I = Io
                if kt <= 0.22:
                    Div_Id_I = 1 - 0.09*kt
                elif 0.22 < kt <= 0.8:
                    Div_Id_I = (
                        0.9511 - 0.1604 * kt
                        + 4.388 * kt**2 - 16.638*kt**3 + 12.336*kt**4
                        )
                else:
                    Div_Id_I = 0.165
                Id = Div_Id_I * I
                Ib = (1-Div_Id_I) * I
                pg = 0.5
                # horizontal calculation
                Ghor = I / 3600
            Gt_list[0].append(Ghor)
            # tilted calculation
            b = b_tilt * pi / 180
            for j in range(len(gamma_list)):
                Gt = 0
                if kt > 0:
                    g = gamma_list[j]
                    w = (w1+w2)/2
                    w_sunrise = - acos((cos(-pi/2) - sin(f)*sin(d)) / cos(f)*cos(d))
                    w_sunset = acos((cos(pi/2) - sin(f)*sin(d)) / cos(f)*cos(d))
                    if not (w_sunrise <= w <= w_sunset):
                        print("Captain. He's not on the list. What should we do?")
                    costhz = cos(f) * cos(d) * cos(w) + sin(f) * sin(d)
                    thz = acos(costhz)
                    sign_w = w / abs(w)
                    gamma_s = sign_w * abs(
                        acos(
                            (cos(thz)*sin(f)-sin(d))/(sin(thz)*cos(f))
                            )
                        )
                    costh = (
                        cos(thz) * cos(b) 
                        + sin(thz) * sin(b) * cos(gamma_s - g)
                        )
                    th = acos(costh) * 180 / pi
                    Rb = costh / costhz
                    if not (0 <= th <= 90):
                        Rb = 0
                    # print('g:',round(g, 2), ', costh:', round(costh, 2), ', costhz:' , round(costhz, 2))
                    It = (
                        + Ib * Rb 
                        + Id * ((1+cos(b))/2) 
                        + I * pg * ((1-cos(b))/2)
                        )
                    if It >= I:
                        It = 0.65 * I
                    Gt = It / 3600
                Gt_list[j+1].append(Gt)
        return Gt_list
    
    # def GetIsolmeasList2(self, 
    #                 month,
    #                 day,
    #                 f,
    #                 b_tilt,
    #                 dt
    #                 ):
    #     f = f * pi / 180
    #     b_tilt = b_tilt * pi / 180
    #     Gsc = 1367
    #     G_list = self.IsolmeasList #W/m2

    #     n = self.DoYCalc(month, day)
    #     gamma_list = [0, -pi/2, pi/2, pi]
    #     B = (n-1) * 360 / 365
    #     d = (
    #         0.006918 
    #         - 0.399912 * cos(B)
    #         + 0.070257 * sin(B)
    #         - 0.006758 * cos(2*B) 
    #         + 0.000907 * sin(2*B) 
    #         - 0.002697 * cos(3*B)
    #         + 0.001480 * sin(3*B)
    #         )
    #     w1_list = []
    #     w2_list = []
    #     w1_list.append(23*15)
    #     w2_list.append(0*15)
    #     for i in range(23):
    #         w1_list.append(i*15)
    #         w2_list.append((i+1)*15)
    #     Gt_list = [[], [], [], [], []]
    #     for i in range(len(G_list)):
    #         w1 = w1_list[i//int(3600/dt)] * pi / 180
    #         w2 = w2_list[i//int(3600/dt)] * pi / 180
    #         G = G_list[i]
    #         I = G_list[i] * 3600 
    #         Io = ( 
    #             12 * 3600 / pi * Gsc * (1+0.033*cos(360/365*n)) 
    #             * (
    #                 cos(f) * cos(d) * (sin(w2)-sin(w1)) 
    #                 + (w2-w1) * sin(f) * sin(d)
    #                 )
    #             )
    #         kt = I/Io
    #         if kt < 0:
    #             kt = 0
    #         print(kt)
    #         if kt <= 0.22:
    #             Div_Id_I = 1 - 0.09*kt
    #         elif 0.22 < kt <= 0.8:
    #             Div_Id_I = (
    #                 0.9511 - 0.1604 * kt
    #                 + 4.388 * kt**2 - 16.638*kt**3 + 12.336*kt**4
    #                 )
    #         else:
    #             Div_Id_I = 0.165
    #         Gd = Div_Id_I * G
    #         Gb = (1-Div_Id_I) * G
    #         pg = 0.5
    #         # horizontal calculation
    #         g = 0
    #         b = 0
    #         w = (w1+w2)/2
    #         w_sunrise = - acos((cos(-pi/2) - sin(f)*sin(d)) / cos(f)*cos(d))
    #         w_sunset = acos((cos(pi/2) - sin(f)*sin(d)) / cos(f)*cos(d))
    #         Gt = 0
    #         if w >= w_sunrise and w <= w_sunset:
    #             costh = (
    #                     sin(d) * sin(f) * cos(b) 
    #                     - sin(d) * cos(f) * sin(b) * cos(g)
    #                     + cos(d) * cos(f) * cos(b) * cos(w) 
    #                     + cos(d) * sin(f) * sin(b) * cos(g) * cos(w)
    #                     + cos(d) * sin(b) * sin(g) * sin(w)
    #                     )
    #             costhz = cos(f) * cos(d) * cos(w) + sin(f) * sin(d)
    #             Rb = costh / costhz
    #             Gt = (
    #                 + Gb * Rb 
    #                 + Gd * ((1+cos(b))/2) 
    #                 + G * pg * ((1-cos(b))/2)
    #                 )
    #         Gt_list[0].append(Gt)
    #         # tilted calculation
    #         b = b_tilt
    #         for j in range(len(gamma_list)):
    #             Gt = 0
    #             g = gamma_list[j]
    #             costh = (
    #                 sin(d) * sin(f) * cos(b) 
    #                 - sin(d) * cos(f) * sin(b) * cos(g)
    #                 + cos(d) * cos(f) * cos(b) * cos(w) 
    #                 + cos(d) * sin(f) * sin(b) * cos(g) * cos(w)
    #                 + cos(d) * sin(b) * sin(g) * sin(w)
    #                 )
    #             if (w >= w_sunrise and w <= w_sunset) and (costh>=0 and costh<=1):
    #                 Rb = costh / costhz
    #                 Gt = (
    #                     + Gb * Rb 
    #                     + Gd * ((1+cos(b))/2) 
    #                     + G * pg * ((1-cos(b))/2)
    #                     )
    #             Gt_list[j+1].append(Gt)
    #     return Gt_list
        
    def GetHeatFlows(self):
        Fsol = self.FsolCalc()
        Fint = self.FintCalc()
        self.Fia = 0.5 * Fint
        self.Fm = self.Am / self.At * (0.5*Fint + Fsol)
        self.Fst = (1 - self.Am/self.At - self.Hw/(9.1*self.At))*(0.5*Fint + Fsol)
        self.FsolList.append(round(Fsol,0))
        self.FintList.append(round(Fint,0))
    
    def GetTair(self, thsup=10, the=10):
        FHCnd = self.Qhvac
        Fmtot =  self.Fm + self.Hem*the + self.H3*(self.Fst + self.Hw*the + self.H1*(((self.Fia + FHCnd)/self.Hve) + thsup)) / self.H2
        tm2 = (self.tm1*(self.Ctotal/3600 - 0.5*(self.H3 + self.Hem)) + Fmtot) / (self.Ctotal/3600 + 0.5*(self.H3 + self.Hem))
        Tm = (self.tm1 + tm2)/2
        Ts = (self.Hms*Tm + self.Fst + self.Hw*the + self.H1*(thsup +(self.Fia + FHCnd)/self.Hve))/(self.Hms + self.Hw + self.H1)
        Tair = (self.His*Ts + self.Hve*thsup + self.Fia + FHCnd)/(self.His + self.Hve)
        tm2 = my_interp(tm2, self.tm1, self.dt)
        Ts = my_interp(Ts, self.Tsprev, self.dt)
        Tair = my_interp(Tair, self.Tairprev, self.dt)
        Top = 0.3*Tair + 0.7*Ts
        return Tair, Top, tm2, Ts
    
    def HveCalc(self):
        dair = 1.2 # kg/m3
        Cpair = 1006 # J/kg*K
        if not self.VdotList_handvalues[self.step*self.dt//3600]:
            Vdot = self.VdotCalc() # m3/h
        else:
            Vdot = self.VdotList_handvalues[self.step*self.dt//3600] # m3/h
        Hve = Vdot * dair * Cpair / 3600
        return Hve
    
    def InitializeTemp(self, temp_init=18, initial_opaque = -100):
        self.Temp_iv = temp_init
        if initial_opaque == -100:
            self.tm1 = BldTempInit(self.Type)
        else:
            self.tm1 = initial_opaque
        self.Tairprev = self.Temp_iv
        self.Tsprev = self.Temp_iv
        self.Topprev = 0.3*self.Tairprev+ 0.7*self.Tsprev
        
    def InitParamsISO(
        self, 
        month=1, 
        day=1,
        HC_schedule=[[0, 0]],
        people_schedule = [[0, 0]],
        people_num = [0],
        HC_maxload = [0],
        Tsetpoints = [[0],[0]],
        Ven_Schedule = [[0, 0]]
        ): #Use before running simulation
        self.SetActionZones(
            HC_sched = HC_schedule,
            people_sched = people_schedule,
            people_n = people_num,
            HC_mload = HC_maxload,
            Tsp = Tsetpoints,
            Ven_sched = Ven_Schedule
            )
        self.GetOutdoorTemp(month, day)
        self.Isol = list(self.GetIsol(month, day))
        # ---
        self.Ctotal = self.Ccoef * self.FloorArea
        self.Hop = 0
        self.OpArea = 0
        for x in self.OpaqueList:
            self.Hop += x.Area*x.U
            self.OpArea += x.Area
        self.Hw = 0
        self.WinArea = 0
        for x in self.FrameList:
            self.Hw += x.Area*x.U
            self.WinArea += x.Area
        if self.Ccoef <= 165000:
            Lam = 2.5
        elif self.Ccoef <= 260000:
            Lam = 2.5 + 0.5 * ((self.Ccoef-165000)/((260000-165000)))
        elif self.Ccoef <= 370000:
            Lam = 3.0 + 0.5 * ((self.Ccoef-260000)/((370000-260000)))
        else:
            Lam = 3.5
        self.Am = Lam * self.FloorArea
        self.At = self.OpArea +self.WinArea #4.5 * self.FloorArea
        self.Um = (self.Hop + self.Hw) / (self.OpArea + self.WinArea)
        self.Hms = 9.1 * self.Am
        self.Hem = 1/(1/self.Hop - 1/self.Hms)
        self.His = 3.45 * self.At
        self.Htot = self.Hop + self.Hw + self.Hve
        self.TimeConstant = self.Ctotal/(3600*self.Htot)
    
    def NextDay(self, month=1, day=1):
        day += 1
        if (
                (month == 1 and day > 31) or
                (month == 2 and day > 28) or 
                (month == 3 and day > 31) or 
                (month == 4 and day > 30) or 
                (month == 5 and day > 31) or 
                (month == 6 and day > 30) or 
                (month == 7 and day > 31) or 
                (month == 8 and day > 31) or 
                (month == 9 and day > 30) or 
                (month == 10 and day > 31) or
                (month == 11 and day > 30)
                ):
                month += 1
                day = 1
        elif month == 12 and day > 31:
            month = 1
            day = 1
        return day, month
        
    def RunSim(self):
        # ---- Iteration ----
        for i in range(0, 86400, self.dt):
            # ---- Get climatic data ----
            self.GetHeatFlows()
            Thsup = self.ThsupList[self.step]
            The = self.TheList[self.step]
            self.SetHve(self.HveCalc()) #W/K
            # ---- GetTair,Tm (turned off hvac)
            self.Qhvac = 0
            Tair0, _, _, _ = self.GetTair(Thsup, The)
            # ---- check Tair set point and decide what actual FHCnd is
            if Tair0 < self.Tset_heat[self.step] and self.HVAC[self.step]:
                # ---- GetTair,Tm (HVAC power = 10 * FloorArea)
                self.Qhvac = 10 * self.FloorArea
                Tair10, _, _, _ = self.GetTair(Thsup, The)
                Qhvac = 10*self.FloorArea*(self.Tset_heat[self.step] - Tair0)/(Tair10 - Tair0)
            elif Tair0 > self.Tset_cool[self.step] and self.HVAC[self.step]:
                self.Qhvac = -10 * self.FloorArea
                # ---- GetTair,Tm (HVAC power = -10 * FloorArea)
                Tair10, _, _, _ = self.GetTair(Thsup, The)
                Qhvac = -10*self.FloorArea*(self.Tset_cool[self.step] - Tair0)/(Tair10 - Tair0)
            else:
                Qhvac = 0
            # Check if there is sufficient heating/cooling
            self.Qhvac = -self.HVAC[self.step]*(Qhvac<=-self.HVAC[self.step]) + self.HVAC[self.step]*(Qhvac>=self.HVAC[self.step]) + Qhvac*(Qhvac>-self.HVAC[self.step] and Qhvac<self.HVAC[self.step])
            # ---- GetTair,Tm final
            Tair, Top, tm2, Ts = self.GetTair(Thsup, The)
            # ---- Save temperature data for plots
            self.TairList.append(round(Tair,2))
            self.ToutList.append(round(The,2))
            self.TmList.append(round(tm2,2))
            self.TsList.append(round(Ts,2))
            self.TopList.append(round(Top,2))
            self.TopListRAW.append(Top)
            self.QhvacList.append(round(self.Qhvac,2))
            # self.FmechvelossList.append(self.Hve*(self.TairList[-1]-Thsup))
            self.FnatvelossList.append(self.Hve*(self.TairList[-1]-The))
            self.FhlossList.append((self.Hop+self.Hw)*(self.TairList[-1]-The))
            # ----
            # ---- Time Constant is calculated here ----
            if i > 0:
                HVAC_off = self.QhvacList[-1] == 0 and self.QhvacList[-2] == 0
                No_solar = self.FsolList[-1] == 0 and self.FsolList[-2] == 0
                No_int = self.FintList[-1] == 0 and self.FintList [-2] == 0 
                Tlastav = (self.ToutList[-2]+self.ToutList[-1])/2
                a = (self.TopListRAW[-2]-Tlastav)/(self.TopListRAW[-1]-Tlastav)
                free_fall = HVAC_off and No_solar and No_int
                if  a > 1 and free_fall:
                    self.TClist.append(round(1/ln(a)*self.dt/3600,2))
            # ---- Prepare for the next timestep
            self.tm1 = tm2
            self.Tairprev = Tair
            self.Tsprev = Ts
            self.step += 1 #it is not a clock, more like a list accessor
            self.HrClock = self.step * self.dt // 3600
            
    def RunSimThermostat(self, operation=None, dT=2):
        # ---- Iteration ----
        for i in range(0, 86400, self.dt):
            # ---- Get climatic data ----
            self.GetHeatFlows()
            Thsup = self.ThsupList[self.step]
            The = self.TheList[self.step]
            self.SetHve(self.HveCalc()) #W/K
            # # ---- GetTair,Tm (turned off hvac)
            self.Qhvac = 0
            Tair0, _, _, _ = self.GetTair(Thsup, The)
            if operation == 'heating':
                if Tair0 < self.Tset_heat[self.step]:
                    self.Qhvac_temp = self.HVAC[self.step]
                elif Tair0 >= self.Tset_heat[self.step] + dT:
                    self.Qhvac_temp = 0
            elif operation == 'cooling':
                if Tair0 > self.Tset_cool[self.step]:
                    self.Qhvac_temp = - self.HVAC[self.step]
                elif Tair0 <= self.Tset_cool[self.step] - dT:
                    self.Qhvac_temp = 0
            else:
                self.Qhvac_temp = 0
            self.Qhvac = self.Qhvac_temp
            # ---- GetTair,Tm final
            Tair, Top, tm2, Ts = self.GetTair(Thsup, The)
            # ---- Save temperature data for plots
            self.TairList.append(round(Tair,2))
            self.ToutList.append(round(The,2))
            self.TmList.append(round(tm2,2))
            self.TsList.append(round(Ts,2))
            self.TopList.append(round(Top,2))
            self.TopListRAW.append(Top)
            self.QhvacList.append(round(self.Qhvac,2))
            # self.FmechvelossList.append(self.Hve*(self.TairList[-1]-Thsup))
            self.FnatvelossList.append(self.Hve*(self.TairList[-1]-The))
            self.FhlossList.append((self.Hop+self.Hw)*(self.TairList[-1]-The))
            # ----
            # ---- Time Constant is calculated here ----
            if i > 0:
                HVAC_off = self.QhvacList[-1] == 0 and self.QhvacList[-2] == 0
                No_solar = self.FsolList[-1] == 0 and self.FsolList[-2] == 0
                No_int = self.FintList[-1] == 0 and self.FintList [-2] == 0 
                Tlastav = (self.ToutList[-2]+self.ToutList[-1])/2
                a = (self.TopListRAW[-2]-Tlastav)/(self.TopListRAW[-1]-Tlastav)
                free_fall = HVAC_off and No_solar and No_int
                if  a > 1 and free_fall:
                    self.TClist.append(round(1/ln(a)*self.dt/3600,2))
            # ---- Prepare for the next timestep
            self.tm1 = tm2
            self.Tairprev = Tair
            self.Tsprev = Ts
            self.step += 1 #it is not a clock, more like a list accessor
            self.HrClock = self.step * self.dt // 3600
            
    def ResetClock(self, Month=1, Day=1):
        self.HrClock = 0
        self.step = 0
        # self.GetOutdoorTemp(Month, Day)
        self.ToutmeasList = None
        self.ThsupmeasList = None
        self.Isol = list(self.GetIsol(Month, Day))
    
    def ResetLists(self):
        self.TairList = []
        self.TsList = []
        self.TmList = []
        self.TopList = []
        self.TopListRAW = []
        self.TheList = []
        self.ToutList = []
        self.IsolmeasList = None
        self.ToutmeasList = None
        self.ThsupmeasList = None
        self.FsolList = []
        self.FintList = []
        self.FhlossList = []
        self.FmechvelossList = []
        self.FnatvelossList = []
        self.QhvacList = []
    
    def SetActionZones(
            self,
            HC_sched = [[0, 0]],
            people_sched = [[0, 0]],
            people_n = [0],
            HC_mload = [0],
            Tsp = [[0], [0]],
            Ven_sched = [[0, 0]],
            ):
        # ---- Occupant schedule
        self.People = self.Construct_scheds(
            people_sched, 
            people_n
            )
        # ---- HVAC schedule
        self.HVAC = self.Construct_scheds(
            HC_sched, 
            HC_mload
            )
        # ---- temp setpoints
        self.Tset_heat = self.Construct_scheds(
            HC_sched, 
            Tsp[0]
            )
        self.Tset_cool = self.Construct_scheds(
            HC_sched, 
            Tsp[1],
            off_value=100
            )
        self.Ventilation = self.Construct_scheds(
            Ven_sched, 
            [0.3]*len(Ven_sched)
            )
        for i in self.FrameList:
            i.VenZones = self.Ventilation
    
    def SetHve(self, hve):
        self.Hve = hve #W/K
        self.H1 = 1/(1/self.Hve + 1/self.His)
        self.H2 = self.H1 + self.Hw
        self.H3 = 1/(1/self.H2 + 1/self.Hms)
  
    def Update(self):
        # ---- Get climatic data ----
        self.GetHeatFlows()
        Thsup = self.ThsupList[self.step]
        The = self.TheList[self.step]
        self.SetHve(self.HveCalc()) #W/K
        # ---- GetTair,Tm (turned off hvac)
        self.Qhvac = 0
        Tair0, _, _, _ = self.GetTair(Thsup, The)
        # ---- check Tair set point and decide what actual FHCnd is
        if Tair0 < self.Tset_heat[self.step] and self.HVAC[self.step]:
            # ---- GetTair,Tm (HVAC power = 10 * FloorArea)
            self.Qhvac = 10 * self.FloorArea
            Tair10, _, _, _ = self.GetTair(Thsup, The)
            Qhvac = 10*self.FloorArea*(self.Tset_heat[self.step] - Tair0)/(Tair10 - Tair0)
        elif Tair0 > self.Tset_cool[self.step] and self.HVAC[self.step]:
            self.Qhvac = -10 * self.FloorArea
            # ---- GetTair,Tm (HVAC power = -10 * FloorArea)
            Tair10, _, _, _ = self.GetTair(Thsup, The)
            Qhvac = -10*self.FloorArea*(self.Tset_cool[self.step] - Tair0)/(Tair10 - Tair0)
        else:
            Qhvac = 0
        # Check if there is sufficient heating/cooling
        self.Qhvac = -self.HVAC[self.step]*(Qhvac<=-self.HVAC[self.step]) + self.HVAC[self.step]*(Qhvac>=self.HVAC[self.step]) + Qhvac*(Qhvac>-self.HVAC[self.step] and Qhvac<self.HVAC[self.step])
        # ---- GetTair,Tm final
        Tair, Top, tm2, Ts = self.GetTair(Thsup, The)
        # ---- Save temperature data for plots
        self.TairList.append(round(Tair,2))
        self.ToutList.append(round(The,2))
        self.TmList.append(round(tm2,2))
        self.TsList.append(round(Ts,2))
        self.TopList.append(round(Top,2))
        self.TopListRAW.append(Top)
        self.QhvacList.append(round(self.Qhvac,2))
        # self.FmechvelossList.append(self.Hve*(self.TairList[-1]-Thsup))
        self.FnatvelossList.append(self.Hve*(self.TairList[-1]-The))
        self.FhlossList.append((self.Hop+self.Hw)*(self.TairList[-1]-The))
        # ---- Prepare for the next timestep
        self.tm1 = tm2
        self.Tairprev = Tair
        self.Tsprev = Ts
        self.step += 1 #it is not a clock, more like a list accessor
        self.HrClock = self.step * self.dt // 3600
        
    def UpdateReal(self):
        # ---- Get climatic data ----
        self.GetHeatFlows()
        Thsup = self.ThsupList[self.step]
        The = self.TheList[self.step]
        self.SetHve(self.HveCalc()) #W/K
        # ---- GetTair,Tm final
        Tair, Top, tm2, Ts = self.GetTair(Thsup, The)
        # ---- Save temperature data for plots
        self.TairList.append(round(Tair,5))
        self.ToutList.append(round(The,5))
        self.TmList.append(round(tm2,5))
        self.TsList.append(round(Ts,5))
        self.TopList.append(round(Top,5))
        self.TopListRAW.append(Top)
        self.QhvacList.append(round(self.Qhvac,2))
        # self.FmechvelossList.append(self.Hve*(self.TairList[-1]-Thsup))
        self.FnatvelossList.append(self.Hve*(self.TairList[-1]-The))
        self.FhlossList.append((self.Hop+self.Hw)*(self.TairList[-1]-The))
        # ---- Prepare for the next timestep
        self.tm1 = tm2
        self.Tairprev = Tair
        self.Tsprev = Ts
        self.step += 1 #it is not a clock, but a list accessor
        self.HrClock = self.step * self.dt // 3600
        
    def UpdateThermostat(self, operation=None, dT=2):
        # ---- Get climatic data ----
        self.GetHeatFlows()
        Thsup = self.ThsupList[self.step]
        The = self.TheList[self.step]
        self.SetHve(self.HveCalc()) #W/K
        # # ---- GetTair,Tm (turned off hvac)
        self.Qhvac = 0
        # -----------------------------------------------------------
        if operation == 'heating':
            if self.Tairprev < self.Tset_heat[self.step]:
                self.Qhvac_temp = self.HVAC[self.step]
            elif self.Tairprev >= (self.Tset_heat[self.step] + dT):
                self.Qhvac_temp = 0
        elif operation == 'cooling':
            if self.Tairprev > self.Tset_cool[self.step]:
                self.Qhvac_temp = - self.HVAC[self.step]
            elif self.Tairprev <= self.Tset_cool[self.step] - dT:
                self.Qhvac_temp = 0
        else:
            self.Qhvac_temp = 0
        # # CAUTION: TO BE USED ONLY IN DEMOKRITOS FILE, WHEN CONTROLLING LOADS
        # shait =  int(10*3600)
        # if self.step == (shait//self.dt) and self.Tset_heat[self.step]!=0:
        #     self.Qhvac_temp = self.HVAC[self.step]
        # ------------------------------------
        self.Qhvac = self.Qhvac_temp
        # ---- GetTair,Tm final
        Tair, Top, tm2, Ts = self.GetTair(Thsup, The)
        # ---- Save temperature data for plots
        self.TairList.append(round(Tair,2))
        self.ToutList.append(round(The,2))
        self.TmList.append(round(tm2,2))
        self.TsList.append(round(Ts,2))
        self.TopList.append(round(Top,2))
        self.TopListRAW.append(Top)
        self.QhvacList.append(round(self.Qhvac,0))
        # self.FmechvelossList.append(self.Hve*(self.TairList[-1]-Thsup))
        self.FnatvelossList.append(self.Hve*(self.TairList[-1]-The))
        self.FhlossList.append((self.Hop+self.Hw)*(self.TairList[-1]-The))
        # ---- Prepare for the next timestep
        self.tm1 = tm2
        self.Tairprev = Tair
        self.Tsprev = Ts
        self.step += 1 #it is not a clock, more like a list accessor
        self.HrClock = self.step * self.dt // 3600
    
    def VdotCalc(self):
        s = 0
        Ct = 0.01
        Cw = 0.001
        Cst = 0.0035
        vmet = 0.5 #m/s
        if len(self.TairList) == 0:
            Ti = self.Temp_iv
        else:
            Ti = self.TairList[-1]
        Te = self.TheList[self.step]
        for i in self.FrameList:
            Hwindow = i.Height
            Aow = i.Area * i.VenZones[self.step]
            V = Ct + Cw * vmet ** 2 + Cst * Hwindow * abs(Ti-Te)
            Vdot = 3.6 * 500 * Aow * V ** 0.5
            s += Vdot
        if s == 0:
            s += 0.000001
        return s #m3/h
    
    
class Opaque():
    
    def __init__(self, length=1, height=1, kj=0, ori='None', fsh=1, abso=0.6, u=0.5, res=0.04):
        self.Length = length
        self.Height = height
        self.Area = self.Length*self.Height
        self.C = kj
        self.Orientation = ori
        self.SetOrientation_num()
        self.Fsh = fsh
        self.abso = abso
        self.U = u
        self.ExternalRes = res
        self.Asol = self.abso*self.ExternalRes*self.U*self.Area
        self.AttachedFrames = []
        self.AttachedFrames_indx = []
        self.Attr = None
        
    def SetOrientation_num(self):
        if self.Orientation == 'North':
            self.Orientation_num = 4
        elif self.Orientation == 'East':
            self.Orientation_num = 2
        elif self.Orientation == 'South':
            self.Orientation_num = 1
        elif self.Orientation == 'West':
            self.Orientation_num = 3
        else:
            self.Orientation_num = 0
        
      
class Frame():
    
    def __init__(self, length=1, height=1, kj=0, ori='None', fsh=1, ggl=0.3, u=2.5, ff=0.2, nvz=[[0, 0]]):
        self.Length = length
        self.Height = height 
        self.Area = self.Length*self.Height
        self.C = kj
        self.Orientation = ori
        self.SetOrientation_num()
        self.U = u
        self.Fsh = fsh
        self.ggl = ggl
        self.Ff = ff
        self.Asol = (1 - self.Ff)*self.ggl*self.Area
        self.VenZones = nvz
        self.Attr = None
        
    def SetOrientation_num(self):
        if self.Orientation == 'North':
            self.Orientation_num = 4
        elif self.Orientation == 'East':
            self.Orientation_num = 2
        elif self.Orientation == 'South':
            self.Orientation_num = 1
        elif self.Orientation == 'West':
            self.Orientation_num = 3
        else:
            self.Orientation_num = 0


def SampleBld(
        tinit=18, 
        weight='Heavy', 
        deltat=3600, 
        climate_file='Athens.xlsx'
        ):
    # ---- Customizable parameters
    ChosenType = weight #Very light / Light / Medium / Heavy / Very heavy
    Tinit = tinit
    # ----
    bld = Building(
        clim_file=climate_file, 
        bld_tinit=Tinit,
        time_interval=deltat,
        bld_type=ChosenType
        )
    # ---- Add building elements
    # ---- Opaques
    bld.AddElement(attr='Wall', l=10, h=3.6, o='North', ul=0.365)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='East', ul=0.365)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='South', ul=0.365)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='West', ul=0.365)
    bld.AddElement(attr='Roof', l=10, h=10, ul=0.305)
    bld.AddElement(attr='Floor', l=10, h=10, ul=0)
    # ---- Windows
    bld.AddElement(attr='Window', l=5, h=2.12, ul=1.62)
    bld.AssignFrametoOpaque(0, 2)
    bld.AddElement(attr='Door', l=1, h=2, ul=1.62)
    bld.AssignFrametoOpaque(1, 1)
    return bld


def SampleBld2(
        tinit=18, 
        weight='Heavy', 
        deltat=3600, 
        climate_file='Athens.xlsx'
        ):
    # ---- Customizable parameters
    ChosenType = weight #Very light / Light / Medium / Heavy / Very heavy
    Tinit = tinit
    # ----
    bld = Building(
        clim_file=climate_file, 
        bld_tinit=Tinit,
        time_interval=deltat,
        bld_type=ChosenType
        )
    # ---- Add building elements
    # ---- Opaques
    bld.AddElement(attr='Wall', l=5, h=3.6, o='North', ul=0.365)
    bld.AddElement(attr='Wall', l=5, h=3.6, o='East', ul=0.365)
    bld.AddElement(attr='Wall', l=5, h=3.6, o='South', ul=0.365)
    bld.AddElement(attr='Wall', l=5, h=3.6, o='West', ul=0.365)
    bld.AddElement(attr='Roof', l=5, h=5, ul=0.305)
    bld.AddElement(attr='Floor', l=5, h=5, ul=0)
    # ---- Windows
    bld.AddElement(attr='Window', l=3, h=2.12, ul=1.62)
    bld.AssignFrametoOpaque(0, 2)
    bld.AddElement(attr='Door', l=1, h=2, ul=1.62)
    bld.AssignFrametoOpaque(1, 1)
    return bld


def SampleBld3(
        tinit=18, 
        weight='Heavy', 
        deltat=3600, 
        climate_file='Athens.xlsx'
        ):
    # ---- Customizable parameters
    ChosenType = weight #Very light / Light / Medium / Heavy / Very heavy
    Tinit = tinit
    # ----
    bld = Building(
        clim_file=climate_file, 
        bld_tinit=Tinit,
        time_interval=deltat,
        bld_type=ChosenType
        )
    # ---- Add building elements
    # ---- Opaques
    bld.AddElement(attr='Wall', l=10, h=3.6, o='North', ul=0.365)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='East', ul=0.365)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='South', ul=0.365)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='West', ul=0.365)
    bld.AddElement(attr='Roof', l=10, h=10, ul=0.305)
    bld.AddElement(attr='Floor', l=10, h=10, ul=0)
    # ---- Windows
    bld.AddElement(attr='Window', l=5, h=2.12, ul=1.62)
    bld.AssignFrametoOpaque(0, 2)
    bld.AddElement(attr='Door', l=1, h=2, ul=1.62)
    bld.AssignFrametoOpaque(1, 1)
    return bld


def SampleBldLowIns(m=1, 
              d=1, 
              tinit=18, 
              weight='Heavy', 
              deltat=3600, 
              heat_zones=[[0*3600, 0*3600]],
              people_zones=[[0*3600, 0*3600]],
              ven_zones = [[0*3600, 0*3600]]
              ):
    # ---- Customizable parameters
    Month = m
    Day = d #30/10 example
    dt = deltat
    Occupants = create_sched(people_zones, dt)
    HC = create_sched(heat_zones, dt)
    VenZones = create_sched(ven_zones, dt)
    ChosenType = weight #Very light / Light / Medium / Heavy / Very heavy
    Tinit = tinit
    # ----
    bld = Building(clim_file='Kozani.xlsx', people=Occupants, bld_type=ChosenType, bld_tinit= Tinit, hvac=HC)
    # ---- Add building elements
    # ---- Opaques
    bld.AddElement(attr='Wall', l=10, h=3.6, o='North', ul=1)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='East', ul=1)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='South', ul=1)
    bld.AddElement(attr='Wall', l=10, h=3.6, o='West', ul=1)
    bld.AddElement(attr='Roof', l=10, h=10, ul=0.8)
    bld.AddElement(attr='Floor', l=10, h=10, ul=0)
    # ---- Windows
    bld.AddElement(attr='Window', l=5, h=2.12, ul=2.5, ven_zones = VenZones)
    bld.AssignFrametoOpaque(0, 2)
    bld.AddElement(attr='Door', l=1, h=2, ul=3, ven_zones = VenZones)
    bld.AssignFrametoOpaque(1, 1)
    # ---- Perform ISO13790 - Simple hourly method initialization
    bld.InitParamsISO(Month, Day)
    return bld


# def RandomBld(m=1, 
#               d=1,  
#               deltat=3600, 
#               heat_zones=[[0*3600, 0*3600]],
#               people_zones=[[0*3600, 0*3600]]
#               ):
#     # ---- Customizable parameters
#     Month = m
#     Day = d #30/10 example
#     dt = deltat
#     Occupants = create_sched(people_zones, dt)
#     HC = create_sched(heat_zones, dt)
#     weight = ['Very light', 'Light', 'Medium', 'Heavy', 'Very heavy']
#     # weight = ['Very light']
#     ChosenType = weight[rd.randrange(0,len(weight))] #Very light / Light / Medium / Heavy / Very heavy
#     Tinit = BldTempInit(ChosenType)
#     # ----
#     bld = Building(clim_file='Kozani.xlsx', people=Occupants, bld_type=ChosenType, bld_tinit=Tinit, hvac=HC)
#     # ---- Add building elements
#     # ---- Opaques
#     rl = rd.randrange(30,201)/10
#     rw = rd.randrange(30,201)/10
#     rh = rd.randrange(10,36)/10
#     rulwall = rd.randrange(100,1001)/1000
#     rulroof = rd.randrange(100,1001)/1000
#     rulfloor = rd.randrange(0,501)/1000
#     bld.AddElement(attr='Wall', l=rw, w=rh, o='North', ul=rulwall)
#     bld.AddElement(attr='Wall', l=rl, w=rh, o='East', ul=rulwall)
#     bld.AddElement(attr='Wall', l=rw, w=rh, o='South', ul=rulwall)
#     bld.AddElement(attr='Wall', l=rl, w=rh, o='West', ul=rulwall)
#     bld.AddElement(attr='Roof', l=rl, w=rw, ul=rulroof)
#     bld.AddElement(attr='Floor', l=rl, w=rw, ul=rulfloor)
#     # ---- Windows
#     rlwin = rd.randrange(10,101)/10
#     rwwin = rd.randrange(5,16)/10
#     rulwin = rd.randrange(1000,1501)/1000
#     bld.AddElement(attr='Window', l=rlwin, w=rwwin, ul=rulwin)
#     bld.AssignFrametoOpaque(0, 2)
#     bld.AddElement(attr='Door', l=1, w=2, ul=1.5)
#     bld.AssignFrametoOpaque(1, 1)
#     # ---- Perform ISO13790 - Simple hourly method initialization
#     bld.InitParamsISO(Month, Day)
#     return bld

def my_interp(x, xprev, dt):
    return (dt/3600)*x + (1-dt/3600)*xprev

def clim_interp(HrClimList, dt, current_hour, current_second):
    HrClimValue = HrClimList[current_hour]
    if current_hour < 23:
        HrClimValue = current_second%3600/3600*HrClimList[current_hour+1] + (1 - current_second%3600/3600)*HrClimList[current_hour] 
    return HrClimValue

def BldTempInit(wtype):
    x = 15
    if wtype=='Very light':
        x = 16.5
    elif wtype=='Light':
        x = 17.2
    elif wtype=='Medium':
        x = 17.5
    elif wtype=='Heavy' or wtype=='Very heavy':
        x = 18
    return x

def ClimateTMYzing(
        import_path='C:\\Users\\Admin\\OneDrive - uowm.gr\\PhD\\Python\\Scripts\\SimulationProject',
        Climate_file='Athens.xlsx',
        export_path='C:\\Users\\Admin\\OneDrive - uowm.gr\\PhD\\Python\\Scripts\\SimulationProject'
        ):
    # ---- data from file
    pdlist = pd.read_excel(io=import_path+'\\'+Climate_file)
    neg = pdlist[pdlist.iloc[:,5:]<0]
    neg = neg[neg.any(axis='columns')]
    neg[neg<0] = 0
    pdlist.update(neg)
    # neg = 0 # make negative radiations zero
    # ---- exported dataframe
    pdTMYlist = pdlist.copy()
    # ----
    for i in range(12):
        for j in range(24):
            a = pdTMYlist[pdTMYlist['m']==(i+1)][pdTMYlist['h']==(j+1)]
            b = a.copy()
            for k in range(len(b)):
                a.iloc[k, 4:]=b.mean()[4:]
                pdTMYlist.update(a)
    pdTMYlist.to_excel(
        excel_writer=export_path+'\\'+Climate_file[:-5]+'_TMY.xlsx'
        )
    return pdlist, pdTMYlist
    

# ---- Test a building ----
if __name__ == "__main__":
    mode = 1
    # =========================================================================
    # ---- Customizable parameters
    Month = 10
    Day = 30 #30/10 example
    dt = 60
    ChosenType = 'Heavy' #Very light / Light / Medium / Heavy / Very heavy
    if mode == 1:
        Tinp = [25.45, 25.3, 25.15, 25, 24.8, 24.5, 24.3, 24.25, 24.1, 24.05, 24, 24.1, 24.5, 24.7, 24.9, 25.1, 25.3, 25.45, 25.5, 25.3, 25.2, 25.1, 25, 24.8, 25.45]
        TinpEP = [25.45, 25.3, 25.15, 25.1, 25, 24.9, 24.8, 24.75, 24.7, 24.75, 25, 25.1, 25.2, 25.45, 25.55, 25.7, 25.75, 25.8, 25.75, 25.7, 25.55, 25.5, 25.4, 25.2, 25.45]
        TinEP = [25.46, 24.66, 24.33, 24.13, 24.07, 24.01, 23.95, 23.89, 23.85, 23.86, 23.96, 24.19, 24.43, 24.65, 24.79, 24.86, 24.83, 24.7, 24.61, 24.57, 24.51, 24.46, 24.4, 24.34, 25.45]
        TinEP = [26.03, 25.48, 24.93, 24.79, 24.75, 24.66, 24.66, 24.66, 24.76, 25.08, 25.32, 25.52, 25.64, 25.69, 25.6, 25.46, 25.31, 25.23, 25.18, 25.12, 25.06, 25, 24.94, 24.88] # Uwall=0.087
        Tinit = Tinp[0]+0.2/3600*dt
        # ----
        bld = Building(
            clim_file='Kozani.xlsx', 
            bld_type=ChosenType, 
            bld_tinit= Tinit,
            time_interval=dt,
        )
        # ---- Add building elements
        # ---- Opaques
        bld.AddElement(attr='Wall', l=10, h=3.6, o='North', ul=0.365)
        bld.AddElement(attr='Wall', l=10, h=3.6, o='East', ul=0.365)
        bld.AddElement(attr='Wall', l=10, h=3.6, o='South', ul=0.365)
        bld.AddElement(attr='Wall', l=10, h=3.6, o='West', ul=0.365)
        bld.AddElement(attr='Roof', l=10, h=10, ul=0.305)
        bld.AddElement(attr='Floor', l=10, h=10, ul=0)
        # ---- Windows
        bld.AddElement(attr='Window', l=5, h=2.12, ul=1.62)
        bld.AssignFrametoOpaque(0, 2)
        bld.AddElement(attr='Door', l=1, h=2, ul=1.62)
        bld.AssignFrametoOpaque(1, 1)
        # ---- Perform ISO13790 - Simple hourly method initialization
        bld.ToutmeasList = [13.5, 12, 10.5, 10.45, 10.4, 10.3, 10.2, 10.2, 10.7, 12.1, 15, 18, 20, 21, 21.8, 21.95, 21.6, 20.5, 19, 17, 15.4, 14, 13.4, 12.6, 13.5] #Ballarini et al. (2019)
        bld.InitParamsISO(Month, Day)
        # ---- Run (Choose between Update and RunSim, they do the same)
        bld.RunSim()
        # for i in range(0, 86400, dt):
            # bld.Update(dt)
        timList = TimeList(dt, bld.step)
        # ---- Single Equation Model
        Um = 0.5
        Aside = 344
        Ctotal = bld.Ctotal/dt
        Ctotal = Ctotal + 1/1000
        Tsem_m_list = []
        Tsin_bfr = Tinit
        for x in range(86400//dt):
            Tsin = Tsin_bfr - 1/Ctotal*(-bld.QhvacList[x]-bld.FnatvelossList[x]-bld.FsolList[x] + bld.Um*Aside*(Tsin_bfr-bld.ToutList[x]))
            Tsem_m_list.append(Tsin)
            Tsin_bfr = Tsin 
        # ---- Plot Data ----
        month_txt = GetMonthtxt(Month)
        hours = mdates.HourLocator(byhour=range(0,24,4))
        hours_fmt = mdates.DateFormatter('%H:00')
        # ---- Figure 1
        fig1, ax1 = plt.subplots()
        Tinp = bld.Convert_timestep(3600, Tinp)
        TinpEP = bld.Convert_timestep(3600, TinpEP)
        # ax1.plot_date(timList, bld.TairList, 'k', label='Indoor air temperature\n(ISO13790, self built)')
        ax1.plot_date(timList, Tinp, 'c', label='Building operative temperature\n(ISO 13790, Ballarini et al., 2019)')
        # ax1.plot_date(timList, bld.TmList, 'r', label='Building temperature\n(ISO13790, self built)')
        # ax1.plot_date(timList, bld.TsList, 'c', label='Building wall surface temperature\n(ISO13790, self-built)')
        ax1.plot_date(timList, bld.TopList, 'g', label='Building operative temperature\n(ISO13790, self-built)')
        ax1.plot_date(timList, TinpEP, 'm', label='Building operative temperature\n(EnergyPlus, Ballarini et al., 2019)')
        # ax1.plot_date(timList, TinEP, 'k', label='Building operative temperature\n(EnergyPlus, self-built)')
        ax1.plot_date(timList, Tsem_m_list, 'y--', label='Building temperature (SEM)')
        # ax1.plot_date(timList, bld.TheList, 'b', label='Outdoor temperature')
        ax1.plot_date(timList, bld.ToutList, 'b', label='Outdoor temperature')
        ax11= ax1.twinx()
        ax11.plot_date(timList, bld.FsolList, 'r--', label='Indoor solar gain load')
        ax11.plot_date(timList, bld.QhvacList, 'm--', label='HVAC load')
        #------
        ax1.set_title(label=bld.Type+" building", fontdict={'fontweight':'heavy', 'fontsize':8})
        ax1.set_xlabel("Time of day")
        ax1.set_ylabel("Temperature ($^o$C)")
        ax11.set_ylabel("Power Load (W)")
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines11, labels11 = ax11.get_legend_handles_labels()
        ax1.legend(lines1 + lines11, labels1 + labels11, loc=(1,1), fontsize=8).set_title(title=month_txt, prop={'weight':'heavy', 'size':8})
        ax1.set_ylim(ymin=23, ymax=27)
        ax11.set_ylim(ymin=0, ymax=10500) #bld.Max_watt_h+500
        ax1.autoscale(enable=True, axis='x', tight=True)
        ax1.xaxis.set_major_locator(hours)
        ax1.xaxis.set_major_formatter(hours_fmt)
        # rotates and right aligns the x labels, and moves the bottom of the
        # axes up to make room for them
        fig1.autofmt_xdate(rotation=0, ha='left')
        fig1.tight_layout(pad=0)
        ax1.grid(True)
        print("Time Constant: "+ str(round(bld.TimeConstant,1))+ " hrs")
    elif mode == 2:
        bld = Building(
            clim_file='Kozani.xlsx', 
            bld_type=ChosenType, 
            bld_tinit= 18,
            time_interval=dt,
        )
        # ---- Add building elements
        # ---- Opaques
        bld.AddElement(attr='Wall', l=10, h=3.6, o='North', ul=0.365)
        bld.AddElement(attr='Wall', l=10, h=3.6, o='East', ul=0.365)
        bld.AddElement(attr='Wall', l=10, h=3.6, o='South', ul=0.365)
        bld.AddElement(attr='Wall', l=10, h=3.6, o='West', ul=0.365)
        bld.AddElement(attr='Roof', l=10, h=10, ul=0.305)
        bld.AddElement(attr='Floor', l=10, h=10, ul=0)
        bld.ToutmeasList = [13.5, 12, 10.5, 10.45, 10.4, 10.3, 10.2, 10.2, 10.7, 12.1, 15, 18, 20, 21, 21.8, 21.95, 21.6, 20.5, 19, 17, 15.4, 14, 13.4, 12.6, 13.5] #Ballarini et al. (2019)
        bld.InitParamsISO(
            Month, 
            Day,
            HC_schedule=[[8*3600-12*15*60, 16*3600]],
            HC_maxload=[2000],
            people_schedule=[[8*3600, 16*3600]],
            people_num=[4],
            Tsetpoints=[[20],[26]], 
            )
        # ---- Run (Choose between Update and RunSim, they do the same)
        bld.RunSimThermostat('heating', dT = 0.5)
        timList = TimeList(dt, bld.step)
        # ---- Plot Data ----
        month_txt = GetMonthtxt(Month)
        hours = mdates.HourLocator(byhour=range(0,24,4))
        hours_fmt = mdates.DateFormatter('%H:00')
        # ---- Figure 1
        fig1, ax1 = plt.subplots()
        ax1.plot_date(timList, bld.TairList, 'k', label='Indoor air temperature\n(ISO13790, self built)')
        #------
        ax1.set_title(label=bld.Type+" building", fontdict={'fontweight':'heavy', 'fontsize':8})
        ax1.set_xlabel("Time of day")
        ax1.set_ylabel("Temperature ($^o$C)")
        lines1, labels1 = ax1.get_legend_handles_labels()
        ax1.legend(lines1, labels1, loc=(1,1), fontsize=8).set_title(title=month_txt, prop={'weight':'heavy', 'size':8})
        ax1.set_ylim(ymin=16, ymax=22)
        ax1.autoscale(enable=True, axis='x', tight=True)
        ax1.xaxis.set_major_locator(hours)
        ax1.xaxis.set_major_formatter(hours_fmt)
        # rotates and right aligns the x labels, and moves the bottom of the
        # axes up to make room for them
        fig1.autofmt_xdate(rotation=0, ha='left')
        fig1.tight_layout(pad=0)
        ax1.grid(True)
        print("Time Constant: "+ str(round(bld.TimeConstant,1))+ " hrs")
        T = bld.TairList