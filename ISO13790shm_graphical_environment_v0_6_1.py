# -*- coding: utf-8 -*-
"""
Created on Thu Feb 19 13:45:56 2026

@author: Leonidas Zouloumis

Description: This program is the graphical environment for the ISO EN 13790 —
simple hourly method simulator\n

Changelog:\n

v0.6.1 - February 10th 2026 (operable)\n
Added the ability to use schedules on certain days of the week


v0.6.0 - October 10th 2025 (operable)\n
Added heating/cooling schedule configuration. The occupancy and 
heating/cooling schedule are now separately configurable.\n
Heating/cooling schedule can be copied from occupancy 
using the respective button\n
Excel sheet results can now be output in different timesteps:
per minute and per hour

v0.5.1 - September 17th 2025 (operable)\n
The user can now get an excel file regarding 
the required heating/ cooling loads of the simulated building \n
WARNING: Keeping the excel file open while simulating will not update it \n
Loads are output in an excel file and are timestep-dependent
plot was replaced by plot due to method deprecation \n
Fixed bug where leaving some empty input boxes of air supply caused crashes \n
"Save plot location" changed to "Saved ploy & and data" \n
Menu option text was slightly changed (Both save and load functions)

v0.5.0 - April 3rd 2024 (operable)\n
Added fresh air supply for each hour of the day

v0.4.1 — Nov 10th 2023 (operable)\n
Added the ability to modify 
the latitude of the location where the building is situated\n

v0.4.0 — Oct 31st 2023 (operable)\n
Created buildings ccan now be saved in excel format, 
and be loaded at any time\n
Added top menu\n
Added save building option (operable)\n
Added load building option (operable)\n
Added reset option (work in progress)\n
'Add element' button is now disabled, unless text is written beside it\n
Fixed a bug where the orientation of a frame element would revert to None,
if "Cancel" button was pressed when choosing "Attached to" value\n

v0.3.0 — Sep 27th 2023 (operable)\n
Custom simulation mode added as a tabgroup in initialization parameter tab\n
Simulation can now be conducted for any time period. 
The option to suppress daily plots during custom simulation is now available\n

v0.2.0 — Sep 19th 2023 (operable)\n
Changes were made in the elements and properties tab\n
Added vertical separator line\n
Added a right column, containing the following entities:\n
    -Climate file location\n
    -Initial settings\n
    -Thermostat settings\n
    -Occupancy schedule\n
    -Save plots\n
    -Simulation\n

v0.1.2 — Sep 12th 2023\n
'Change value' button is now operable\n
Added a left column, containing the following entities:\n
    -Building weight category\n
    -Building weight\n
    -Element properties\n
Tabs are fully functional. A building can now be constructed from the interface
\n

v0.1.1 — Sep 6th 2023\n
Modified the properties tab (It has now tighter layout)\n

v0.1.0 — Sep 5th 2023\n
Core window created
"""

import PySimpleGUI as sg
from ISO13790shm import Building, Opaque
from matplotlib import pyplot as plt
from TimeLib import TimeList, GetMonthtxt
import matplotlib.dates as mdates
import numpy as np
import pandas as pd
import os.path

class Simulator():
    
    def __init__(self):
        # ---- For making decisions
        self.event = None
        self.values = None
        self.last_plot_num = 0
        self.path = ''
        self.weekly_cnt = 0
        # ---- For finding the selected element in the simulated building
        self.sel_indx = None
        self.sel_bld_elem = None
        self.sel_opq_indx = None
        self.sel_opq = None
        # ---- Lists used for storing values of the building elements
        self.OpaqueList = []
        self.FrameList = []
        self.FrameListAttachedto = []
        self.FramesonOpaques = []
        self.SelectedItem = ''
        self.SimulatedBuilding = Building()
        self.NoSimReasonsList = [
            '– No building weight category selected\n',
            (
                '– At least a floor, a roof and four '
                +'walls (North, East, South, West) must exist\n'
            ),
            '– Not all windows and doors are attached to opaque elements\n',
            '– Not all windows and doors are oriented\n',
            '– No climate file selected\n',
            '– No initialization parameters selected\n',
            '– No thermostat settings selected\n',
            '– Occupancy schedule is not configured\n',
            '– No path for plot saving is selected\n',
            '– Building location latitude must be set\n'
            ]
        self.Ready = [0]*len(self.NoSimReasonsList)
        self.OccupancySchedule = []
        self.HeatCoolSchedule = []
        self.VdotList=[None]*24
        # ---- Creation of the window that hosts the layout
        Col_left = [
           [self.ElementListboxTab()],
           [self.PropertiesTab()] 
           ]
        Col_right = [
            [self.ClimateFileTab()],
            [self.InitialSettingsTab(), self.ThermostatSettingsTab()],
            [self.OccupancyScheduleTab()],
            [
                
                    self.SavePlotTab(),self.SimulationTab()
                ]
           ]
        self.Window = sg.Window(
            title='ISO13790—single hourly model simulator', 
            layout=[
                [
                    sg.Column(
                        layout=Col_left, 
                        element_justification='center'
                        ),
                    sg.VerticalSeparator(color='Gray'),
                    sg.Column(
                        layout=Col_right,
                        element_justification='center'
                        )
                    ],
                [
                    self.TopMenu()
                    ]
                ],
            resizable=False,
            )
        # ---- Lists used in plots
        self.FigList = []
        self.AxList = []
        self.plot_cnt = 0
        # ---- Useful output data
        self.Eheat = 0 # [kWh/m2]
        self.Ecool = 0 # [kWh/m2]
        self.Eheat_monthList = [] # [kWh/m2]
        self.Ecool_monthList = [] # [kWh/m2]
        # ---- Tick mark and sum list for excel output
        self.tm_mins = 1
        self.sum_list = [0] * 5
    
    def AddElementtoLists(self):
        window = sg.Window(
            title='Choose the type of the element to add',
            layout=self.ChooseElementTypeTab()
            )
        event, values = window.read()
        chosen = None
        if (event == 'Wall' or event == 'Floor' or event == 'Roof'):
            chosen = 'OPAQUE_LIST'
        elif (event == 'Window' or event == 'Door'):
            chosen = 'FRAME_LIST'
        window.close()
        if chosen =='OPAQUE_LIST':
            self.OpaqueList.append(self.values['ELEMENT_INPUT'])
            self.FramesonOpaques.append([])
            self.Window[chosen].update(self.OpaqueList)
            self.ChangeAttachedToValues()
            self.SimulatedBuilding.AddElement(attr=event)
        elif chosen =='FRAME_LIST':
            self.FrameList.append(self.values['ELEMENT_INPUT'])
            self.FrameListAttachedto.append('None')
            self.Window[chosen].update(self.FrameList)
            self.SimulatedBuilding.AddElement(attr=event)
        self.Window['ELEMENT_INPUT'].update(value='')
    
    def CalculateAsol(self, element):
        if isinstance(element, Opaque):
            element.Asol = (element.abso
                            *element.ExternalRes
                            *element.U
                            *element.Area
                            )
        else:
            element.Asol = (1 - element.Ff)*element.ggl*element.Area
    
    def CalculateElemArea(self, element):
        element.Area = element.Height * element.Length
        if (
                element.Attr == 'Wall' 
                or element.Attr == 'Roof'
                or element.Attr == 'Floor'
                ):
            for i in element.AttachedFrames:
                element.Area -= i.Area
            self.SimulatedBuilding.FloorArea = 0
            for i in self.SimulatedBuilding.OpaqueList:
                if i.Attr == 'Floor':
                    self.SimulatedBuilding.FloorArea += i.Length * i.Height
        else:
            frame_indx = self.SimulatedBuilding.FrameList.index(element)
            if self.Window[
                'Attached to_VALUE_f'
                ].DisplayText[
                    -len(
                        self.FrameListAttachedto[frame_indx]
                        ):
                    ]!= 'None':
                opq_indx = self.OpaqueList.index(
                    self.Window['Attached to_VALUE_f'].DisplayText[
                        -len(self.FrameListAttachedto[frame_indx]):
                            ]
                    )
                opq = self.SimulatedBuilding.OpaqueList[opq_indx]
                opq.Area = opq.Height * opq.Length
                for i in opq.AttachedFrames:
                    opq.Area -= i.Area
    
    def ChangeAttachedToValues(self):
        opaque_list_temp = ['None']
        previous_value = self.values['Attached to_SETTER_f']
        for i in self.OpaqueList:
            opaque_list_temp.append(i)
        self.Window['Attached to_SETTER_f'].update(
            values=opaque_list_temp,
            value=previous_value
            )
        
    def CheckReady1(self):
        elems = [0]*3
        oris = [0] * 4
        all_elems = 0
        all_oris = 0
        i = 0
        while i < len(self.OpaqueList) and not(all_elems*all_oris):
            if self.SimulatedBuilding.OpaqueList[i].Attr == 'Roof':
                elems[0] = 1
            elif self.SimulatedBuilding.OpaqueList[i].Attr == 'Floor':
                elems[1] = 1
            elif (
                    self.SimulatedBuilding.OpaqueList[i].Attr == 'Wall' 
                    and self.SimulatedBuilding.OpaqueList[
                        i
                        ].Orientation_num > 0
                    ):
                elems[2] = 1
                oris[
                    self.SimulatedBuilding.OpaqueList[i].Orientation_num-1
                    ] = self.SimulatedBuilding.OpaqueList[
                        i
                        ].Orientation_num > 0
            all_elems = sum(elems) == len(elems)
            all_oris = sum(oris) == len(oris)
            i += 1
        return all_elems*all_oris
    
    def CheckReady2(self):
        all_attached = 1
        i = 0
        while i < len(self.FrameListAttachedto) and all_attached:
            all_attached = self.FrameListAttachedto[i] != 'None'
            i += 1
        return all_attached
    
    def CheckReady3(self):
        all_oriented = 1
        i = 0
        while i < len(self.FrameList) and all_oriented:
            all_oriented = (
                self.SimulatedBuilding.FrameList[i].Orientation != 'None'
                )
            i += 1
        return all_oriented
        
    def ChooseElementTypeTab(self):
        layout = [
            [
                sg.Button('Wall'),
                sg.Button('Roof'),
                sg.Button('Floor'),
                sg.Button('Window'),
                sg.Button('Door')
                ]
            ]
        return layout
    
    def ClimateFileTab(self):
        frame = sg.Column(
            layout=[
                [
                    sg.Frame(
                        title='Climate file (.xlsx) location ',
                        title_location='n',
                        layout=[
                            [
                                sg.Input(
                                    default_text='',
                                    key='CLIMATE_FILE_PATH',
                                    size=(40,5),
                                    enable_events=True
                                    ),
                                sg.FileBrowse(
                                    button_text='Open File...',
                                    key='CLIMATE_FILE_PATH_SEARCH',
                                    file_types=[
                                        ('Excel files (.xlsx)','*.xlsx'),
                                        ],
                                    
                                    )
                                ]
                            ],
                        )
                    ]
                ],
            )
        return frame
                
    def ConfigureSchedules(self):
        current_color = self.Window[self.event].ButtonColor[1]
        new_color = 'dark red'
        if current_color == 'dark red':
            new_color = 'Green'
        self.Window[self.event].update(button_color=new_color)
        self.UpdateSchedules()
        
    
        
    
    def ConfirmCancelPropertyValue(self):
        prop = (
            self.event[:-9] * (self.event[-10:-1] != '_CONFIRM_')
            + self.event[:-10] * (self.event[-10:-1] == '_CONFIRM_')
            )
        of_listbox = self.event[-1]
        try:
            final_value = float(self.values[prop+'_SETTER_'+of_listbox])
        except ValueError:
            final_value = -10000
        if prop == 'Orientation':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.Orientation = self.values[
                    prop+'_SETTER_'
                    +of_listbox
                    ]
                self.sel_bld_elem.SetOrientation_num()    
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.Orientation))
                        ) 
                    + str(self.sel_bld_elem.Orientation)
                ),
                visible=True
                )
            for i in self.sel_bld_elem.AttachedFrames:
                i.Orientation = self.values[
                    prop+'_SETTER_'
                    +of_listbox
                    ]
                i.SetOrientation_num()
        elif prop == 'Length':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.Length = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.Length))
                        ) 
                    + str(self.sel_bld_elem.Length)
                ),
                visible=True
                )
            self.CalculateElemArea(self.sel_bld_elem)
            self.CalculateAsol(self.sel_bld_elem)
        elif prop == 'Height/Width':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.Height = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.Height))
                        ) 
                    + str(self.sel_bld_elem.Height)
                ),
                visible=True
                )
            self.CalculateElemArea(self.sel_bld_elem)
            self.CalculateAsol(self.sel_bld_elem)
        elif prop == 'Heat loss factor':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.U = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.U))
                        ) 
                    + str(self.sel_bld_elem.U)
                ),
                visible=True
                )
            self.CalculateAsol(self.sel_bld_elem)
        elif prop == 'Shading factor':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.Fsh = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.Fsh))
                        ) 
                    + str(self.sel_bld_elem.Fsh)
                ),
                visible=True
                )
        elif prop == 'External resistance':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.ExternalRes = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.ExternalRes))
                        ) 
                    + str(self.sel_bld_elem.ExternalRes)
                ),
                visible=True
                )
            self.CalculateAsol(self.sel_bld_elem)
        elif prop == 'Absorptivity factor':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.abso = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.abso))
                        ) 
                    + str(self.sel_bld_elem.abso)
                ),
                visible=True
                )
            self.CalculateAsol(self.sel_bld_elem)
        elif prop == 'Attached to':
            if self.event[-10:-1] == '_CONFIRM_':
                if self.values[prop+'_SETTER_'+of_listbox] != 'None':
                    self.opq_indx = self.OpaqueList.index(
                        self.values[prop+'_SETTER_'+of_listbox]
                        )
                    temp = self.FrameListAttachedto[self.sel_indx]
                    self.FrameListAttachedto[self.sel_indx] = self.values[
                        prop+'_SETTER_'+of_listbox
                        ]
                    if temp != 'None':
                        previous_opq_indx = self.OpaqueList.index(
                            temp
                            )
                        self.SimulatedBuilding.RemoveFramefromOpaque(
                                self.sel_indx, 
                                previous_opq_indx
                                )
                    self.SimulatedBuilding.AssignFrametoOpaque(
                            self.sel_indx, 
                            self.opq_indx
                            )
                    self.SimulatedBuilding.FrameList[
                        self.sel_indx
                        ].Orientation = self.SimulatedBuilding.OpaqueList[
                            self.opq_indx
                            ].Orientation
                    self.SimulatedBuilding.FrameList[
                        self.sel_indx
                        ].SetOrientation_num()
                    self.FramesonOpaques[self.opq_indx].append(self.sel_indx)
                elif self.Window[
                        prop+'_VALUE_'+ of_listbox
                        ].DisplayText[
                            -len(
                                self.FrameListAttachedto[self.sel_indx]
                                ):
                            ]!= 'None':
                    self.opq_indx = self.OpaqueList.index(
                        self.Window[prop+'_VALUE_'+of_listbox].DisplayText[
                            -len(self.FrameListAttachedto[self.sel_indx]):
                                ]
                        )
                    self.FrameListAttachedto[self.sel_indx] = self.values[
                        prop+'_SETTER_'+of_listbox
                        ]
                    self.FramesonOpaques[self.opq_indx].remove(self.sel_indx)
                    self.SimulatedBuilding.RemoveFramefromOpaque(
                        self.sel_indx, 
                        self.opq_indx
                        )
                    for i in self.FramesonOpaques[self.opq_indx]:
                        self.SimulatedBuilding.AssignFrametoOpaque(
                                i, 
                                self.opq_indx
                                )
                opq_elem = self.SimulatedBuilding.OpaqueList[self.opq_indx]
                self.CalculateElemArea(opq_elem)
                self.CalculateAsol(opq_elem)
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(self.FrameListAttachedto[self.sel_indx])
                        ) 
                    + self.FrameListAttachedto[self.sel_indx]
                ),
                visible=True
                )
            self.Window['Orientation_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.Orientation))
                        ) 
                    + str(self.sel_bld_elem.Orientation)
                ),
                visible=True
                )
        elif prop == 'Glazing factor':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.ggl = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.ggl))
                        ) 
                    + str(self.sel_bld_elem.ggl)
                ),
                visible=True
                )
            self.CalculateAsol(self.sel_bld_elem)
        elif prop == 'Frame factor':
            if self.event[-10:-1] == '_CONFIRM_':
                self.sel_bld_elem.Ff = final_value
            self.Window[prop+'_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.Ff))
                        ) 
                    + str(self.sel_bld_elem.Ff)
                ),
                visible=True
                )
            self.CalculateAsol(self.sel_bld_elem)
        # ----
        self.Window[prop+'_CHANGE_'+of_listbox].update(
            text='Change value', 
            visible=True
            )
        self.Window[prop+'_SETTER_'+of_listbox].update(
            visible=False
            )
        self.Window[prop+'_CONFIRM_'+of_listbox].update(
            visible=False
            )
        self.Window[prop+'_CANCEL_'+of_listbox].update(
            visible=False
            ) 
    
    def CreatePlots(self, day=1, month=1, finished=0):
        timList = TimeList(
            self.SimulatedBuilding.dt, 
            self.SimulatedBuilding.step
            )
        # month_txt = GetMonthtxt(self.values['MONTH_VALUE'])
        hours = mdates.HourLocator(byhour=range(0,24,4))
        hours_fmt = mdates.DateFormatter('%H:00')
        # ---- Plot data
        the_title = (
            self.SimulatedBuilding.Type+" building, "
            + str(day)+'/'
            + str(month)+',\n '
            + 'max HC power: '
            + self.values['MAX_HC_VALUE']
            + ' W'
            )
        annual_title = (
            self.SimulatedBuilding.Type+" building,\n"
            + 'max HC power: '
            + self.values['MAX_HC_VALUE']
            + ' W'
            )
        max_power = max(
            max(self.SimulatedBuilding.FnatvelossList),
            max(self.SimulatedBuilding.FhlossList),
            max(self.SimulatedBuilding.QhvacList),
            max(self.SimulatedBuilding.FsolList),
            float(self.values['MAX_HC_VALUE'])
            )
        temp_limits = (0, 30)
        min_temp = min(
            min(self.SimulatedBuilding.TairList),
            min(self.SimulatedBuilding.TmList),
            min(self.SimulatedBuilding.ToutList),
            temp_limits[0]
            )
        max_temp = max(
            max(self.SimulatedBuilding.TairList),
            max(self.SimulatedBuilding.TmList),
            max(self.SimulatedBuilding.ToutList),
            temp_limits[1]
            )
        self.Eheat = float(0)
        self.Ecool = float(0)
        for i in self.SimulatedBuilding.QhvacList:
            if i>0:
                self.Eheat += i * self.SimulatedBuilding.dt / 3600 / 1000
            else:
                self.Ecool -= i * self.SimulatedBuilding.dt / 3600 / 1000
        legend_title_daily = (
            'Daily consumption\n'
            + 'Heating: ' 
            + str(round(self.Eheat, 2))
            + ' kWh\n'
            + 'Cooling: ' 
            + str(round(self.Ecool, 2))
            + ' kWh\n'
            )
        legend_title_total = (
            'Annual consumption\n'
            + 'Heating: ' 
            + str(round(sum(self.Eheat_monthList), 2))
            + ' kWh/m$^2$\n'
            + 'Cooling: ' 
            + str(round(sum(self.Ecool_monthList), 2))
            + ' kWh/m$^2$\n'
            )
        if (
                (
                    not(self.values['SUPPRESS_DAILY_PLOTS'])
                    or self.values['INIT_TAB'] == 'DAILY_TAB'
                    )
                and not(finished)
                ):
            self.CreatePlotsDaily(
                timList, 
                the_title, 
                legend_title_daily, 
                min_temp, 
                max_temp, 
                hours, 
                hours_fmt, 
                max_power
                )
        if self.values['INIT_TAB'] == 'CUSTOM_TAB' and finished:
            self.CreatePlotsTotalEnergy(
                annual_title, 
                legend_title_total, 
                )
            
    def CreatePlotsDaily(
            self, 
            timList, 
            the_title, 
            legend_title, 
            min_temp, 
            max_temp, 
            hours, 
            hours_fmt, 
            max_power
            ):
        # ---- Figure 1
        fig1, ax1 = plt.subplots()
        ax1.plot(
            timList, 
            self.SimulatedBuilding.TairList, 
            'r', 
            label='Indoor air temperature'
            )
        ax1.plot(
            timList, 
            self.SimulatedBuilding.TmList, 
            'b', 
            label='Building temperature'
            )
        ax1.plot(
            timList, 
            self.SimulatedBuilding.ToutList, 
            'y', 
            label='Outdoor temperature'
            )
        # ----
        ax1.set_title(
            label= the_title,
            fontdict={'fontweight':'heavy', 'fontsize':8}
            )
        ax1.set_xlabel("Time of day")
        ax1.set_ylabel("Temperature ($^o$C)")
        lines1, labels1 = ax1.get_legend_handles_labels()
        ax1.legend(
            lines1, 
            labels1, 
            loc=(1.02,0.5), 
            fontsize=6
            ).set_title(title=legend_title, prop={'weight':'heavy', 'size':6})
        ax1.set_ylim()
        ax1.set_ylim(ymin=min_temp, ymax=max_temp)
        ax1.autoscale(enable=True, axis='x', tight=True)
        ax1.xaxis.set_major_locator(hours)
        ax1.xaxis.set_major_formatter(hours_fmt)
        # rotates and right aligns the x labels, and moves the bottom of the
        # axes up to make room for them
        fig1.autofmt_xdate(rotation=0, ha='left')
        fig1.tight_layout(pad=1)
        ax1.grid(True)
        # ---- Figure 2
        fig2, ax2 = plt.subplots()
        ax2.plot(
            timList, 
            self.SimulatedBuilding.QhvacList, 
            'r', 
            label='Heating(+)/Cooling(-) power'
            )
        ax2.plot(
            timList, 
            self.SimulatedBuilding.FhlossList, 
            'b', 
            label='Building thermal conductivity\nheat losses(+)/gains(-)'
            )
        ax2.plot(
            timList, 
            self.SimulatedBuilding.FsolList, 
            'y', 
            label='Building solar thermal gains'
            )
        ax2.plot(
            timList, 
            self.SimulatedBuilding.FintList, 
            'k', 
            label='Building internal thermal gains'
            )
        ax2.plot(
            timList, 
            self.SimulatedBuilding.FnatvelossList, 
            'g', 
            label='Building ventilation\nheat losses(+)/gains(-)'
            )
        # ---- 
        ax2.set_title(
            label=the_title, 
            fontdict={'fontweight':'heavy', 'fontsize':8}
            )
        ax2.set_xlabel("Time of day")
        ax2.set_ylabel("Power (W)")
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax2.legend(
            lines2, 
            labels2, 
            loc=(1.02,0.5), 
            fontsize=6
            ).set_title(title=legend_title, prop={'weight':'heavy', 'size':6})
        ax2.set_ylim()
        ax2.set_ylim(
            ymin=-max_power, 
            ymax=max_power
            )
        ax2.autoscale(enable=True, axis='x', tight=True)
        ax2.xaxis.set_major_locator(hours)
        ax2.xaxis.set_major_formatter(hours_fmt)
        # rotates and right aligns the x labels, and moves the bottom of the
        # axes up to make room for them
        fig2.autofmt_xdate(rotation=0, ha='left')
        fig2.tight_layout(pad=1)
        ax2.grid(True)
        # ----
        self.FigList.append(fig1)
        self.AxList.append(ax1)
        self.FigList.append(fig2)
        self.AxList.append(ax2)
        plt.close('all')
        
    def CreatePlotsTotalEnergy(
            self,
            the_title, 
            legend_title, 
            ):
        x = np.arange(len(self.Ecool_monthList))
        w = 0.1
        # ---- Figure 3
        fig3, ax3 = plt.subplots()
        ax3.bar(
            x,
            self.Eheat_monthList,
            color='r',
            label='Heating energy',
            width=w,
            )
        ax3.bar(
            x+w,
            self.Ecool_monthList,
            color='b',
            label='Cooling energy',
            width=w,
            )
        # ---- 
        ax3ticks = self.Window[
            'FROM_MONTH_COMBO'
            ].Values[self.month_st-1:self.month_f]
        if self.month_st == self.month_f:
            ax3ticks = [self.Window[
                'FROM_MONTH_COMBO'
                ].Values[self.month_st-1]]
        ax3.set_xticks(
            x+w/2,
            ax3ticks
            )
        ax3.set_title(
            label=the_title, 
            fontdict={'fontweight':'heavy', 'fontsize':8}
            )
        ax3.set_xlabel("Months")
        ax3.set_ylabel("Energy consumption (kWh/m$^2$)")
        lines3, labels3 = ax3.get_legend_handles_labels()
        ax3.legend(
            lines3, 
            labels3, 
            loc=(1.02,0.5), 
            fontsize=6
            ).set_title(title=legend_title, prop={'weight':'heavy', 'size':6})
        ax3.autoscale(enable=True, axis='x', tight=True)
        fig3.tight_layout(pad=1)
        ax3.grid(axis='y')
        self.FigList.append(fig3)
        self.AxList.append(ax3)
    
    def CreatePropertyDefault(self, key_text, unit):
        prop_line = [
            sg.Text(' '*20, key=key_text+'_DEFAULT'),
            ]
        return prop_line        
    
    def CreatePropertyFrame(self, key_text, unit):
        prop_line = [
            sg.Text(
                key_text+': ', 
                key=key_text+'_GETTER_f',
                size=self.max_length_features
                ),
            sg.pin(
                sg.Text(
                    ' ', 
                    key=key_text+'_VALUE_f',
                    size=9
                    )
                ),
            sg.pin(
                self.ValuePropertyFrame(key_text)
                ),
            sg.Text(
                unit, 
                key=key_text+'_UNIT_f',
                size=self.max_length_units
                ),
            sg.Button(
                'Change value', 
                key=key_text+'_CHANGE_f',
                size=15
                ),
            sg.Button(
                'Confirm', 
                key=key_text+'_CONFIRM_f', 
                visible=False
                ),
            sg.Button(
                'Cancel', 
                key=key_text+'_CANCEL_f', 
                visible=False
                )
            ]
        return prop_line
    
    def CreatePropertyOpaque(self, key_text, unit):
        prop_line = [
            sg.Text(
                key_text+': ', 
                key=key_text+'_GETTER_o',
                size=self.max_length_features
                ),
            sg.pin(
                sg.Text(
                    ' ', 
                    key=key_text+'_VALUE_o',
                    size=9
                    )
                ),
            sg.pin(
                self.ValuePropertyOpaque(key_text)
                ),
            sg.Text(
                unit, 
                key=key_text+'_UNIT_o',
                size=self.max_length_units
                ),
            sg.Button(
                'Change value', 
                key=key_text+'_CHANGE_o',
                size=15
                ),
            sg.Button(
                'Confirm', 
                key=key_text+'_CONFIRM_o', 
                visible=False
                ),
            sg.Button(
                'Cancel', 
                key=key_text+'_CANCEL_o', 
                visible=False
                )
            ]
        return prop_line
   
    def DisableAddButton(self):
        self.Window['ADD'].update(disabled=self.values['ELEMENT_INPUT'] == '')
   
    def ElementListboxTab(self):
        bld = sg.Column(
            layout=[
                [
                    sg.Frame(
                        title='Building weight category',
                        title_location='n',
                        layout= [
                            [
                                sg.Combo(
                                    [
                                        'None',
                                        'Very light',
                                        'Light',
                                        'Medium',
                                        'Heavy',
                                        'Very heavy'
                                        ],
                                    default_value='None',
                                    readonly=True,
                                    enable_events=True,
                                    key='INPUT_WEIGHT_CATEGORY',
                                    )
                                ]
                            ],
                        element_justification='center'
                        ),
                    sg.Frame(
                        title='Latitude of the building location',
                        title_location='n',
                        layout= [
                            [
                                sg.Input(
                                    enable_events=True,
                                    key='BUILDING_LATITUDE',
                                    size = 10
                                    ),
                                sg.Text(
                                    text = '\u2070'
                                    )
                                ]
                            ],
                        element_justification='center'
                        )
                    ],
                [
                    sg.Frame(
                        title='List of building elements',
                        title_location='n',
                        layout = [
                            [
                                sg.Frame(
                                    title='Walls, Floors, Roofs',
                                    title_location='n',
                                    layout= [
                                        [
                                            sg.Listbox(
                                                values=[], 
                                                size=(20,5),
                                                key='OPAQUE_LIST',
                                                enable_events=True
                                                )
                                            ]
                                        ]
                                    ),
                                sg.Frame(
                                    title='Windows, Doors',
                                    title_location='n',
                                    layout= [
                                        [
                                            sg.Listbox(
                                                values=[], 
                                                size=(20,5),
                                                key='FRAME_LIST',
                                                enable_events=True
                                                )
                                            ]
                                        ]
                                    )
                                ],
                            [
                                sg.Input(
                                    key='ELEMENT_INPUT', 
                                    do_not_clear=True,
                                    enable_events=True
                                    ), 
                                sg.Button(
                                    button_text='Add Element', 
                                    key='ADD',
                                    disabled=True,
                                    )
                                ]                                        
                            ]
                        )
                    ],
                ],
            element_justification='center'
            )
        return bld
    
    def FreshAirSupply(self):
        lt = []
        for i in range(24):
            lt.append(
                sg.Input(
                    default_text='',
                    pad=1,
                    key='FRESH_AIR_IN_'+str(i),
                    size=4,
                    enable_events=True
                    )
                )
        frame = sg.Frame(
            title= 'Fresh Air Supply [m\u00b3/h]',
            layout = [lt],
            title_location='n',
            element_justification='center'
            )
        col = sg.Col(layout=[[frame]], justification='center')
        return col
    
    def GetMonth(self, key_str):
        if self.values[key_str] == 'Jan':
            month = 1
        elif self.values[key_str] == 'Feb':
            month = 2
        elif self.values[key_str] == 'Mar':
            month = 3
        elif self.values[key_str] == 'Apr':
            month = 4
        elif self.values[key_str] == 'May':
            month = 5
        elif self.values[key_str] == 'Jun':
            month = 6
        elif self.values[key_str] == 'Jul':
            month = 7
        elif self.values[key_str] == 'Aug':
            month = 8
        elif self.values[key_str] == 'Sep':
            month = 9
        elif self.values[key_str] == 'Oct':
            month = 10
        elif self.values[key_str] == 'Nov':
            month = 11
        else:
            month = 12
        return month
    
    def GetMonthsCustom(self):
        starting_month = self.GetMonth('FROM_MONTH_COMBO')
        finishing_month = self.GetMonth('TO_MONTH_COMBO')
        return starting_month, finishing_month
    
    def HCScheduleBar(self):
        Col_list = []
        for i in range(49):
            if (i%2)==0:
                upper_element = sg.Text(
                    text=' '*(len(str(i//2))==1)+str(i//2),
                    font=('Helvetica',10),
                    pad=0
                    )
                lower_element = sg.Text(
                    text='',
                    pad=0
                    )
            else:
                upper_element = sg.Text(
                    text='',
                    pad=0
                    )
                lower_element = sg.Button(
                    button_color='Dark Red',
                    key="HC_SCHED_BUTTON_"+str(i//2),
                    size=1,
                    pad=0,
                    )
            Col = sg.Column(
                layout= [
                    [upper_element],
                    [lower_element]
                    ],
                pad=0,
                element_justification='center'
                )
            Col_list.append(Col)
        two_texts = [
            sg.Text(
                text='On',
                background_color='Green'
                ),
            sg.Text(
                text='Off',
                background_color='Dark Red'
                )
            ]
        OccupancyToHCButton = [
            sg.Button(
                button_text = 'Copy from Occupancy',
                key = 'HCcopiesOccupancy',
                enable_events=True
                )
            ]
        frame = sg.Frame(
            title='Heating/Cooling schedule',
            title_location = 'n',
            layout = [
                Col_list,
                two_texts,
                OccupancyToHCButton
                ],
            element_justification='center',
            )
        col_final = sg.Column(layout=[[frame]],justification='center') 
        return col_final
    
    def HidePropertiesDetails(self):
        self.Window['COLUMN_DEFAULT'].update(visible=True)
        self.Window['COLUMN_OPAQUE'].update(
            visible=False
            )
        self.Window['COLUMN_FRAME'].update(
            visible=False
            )
    
    def HorSeparator(self, length=1):
        return [sg.Text('_'*length)]
    
    def InitialSettingsTab(self):
        frame1 = sg.Column(
            layout=[
                [sg.VPush()],
                [sg.VPush()],
                [sg.VPush()],
                [
                    sg.Column(
                        layout=[
                            [
                                sg.Text(
                                    text='Day of month:',
                                    size=15
                                    ),
                                ],
                            [
                                sg.Text(
                                    text='Month:',
                                    size=15
                                    ),
                                ],
                            ]
                        ),
                    sg.Column(
                        layout=[
                            [
                                sg.Input(
                                    size=10,
                                    enable_events=True,
                                    key='DAY_VALUE'
                                    ),
                                ],
                            [
                                sg.Input(
                                    size=10,
                                    enable_events=True,
                                    key='MONTH_VALUE'
                                    ),
                                ],
                            ]
                        ),
                    ],
                ],
            )
        frame2 = sg.Column(
            layout=[
                [
                    sg.Column(
                        layout=[
                            [sg.Text(text='From:', size=5)],
                            [sg.Text(text='To:', size=5)]
                            ]
                        ),
                    sg.Column(
                        layout=[
                            [
                                sg.Input(
                                    key='FROM_DAY_INPUT', 
                                    size=2,
                                    enable_events=True
                                    )
                                ],
                            [
                                sg.Input(
                                    key='TO_DAY_INPUT', 
                                    size=2,
                                    enable_events=True
                                    )
                                ]
                            ]
                        ),
                    sg.Column(
                        layout=[
                            [
                                sg.Combo(
                                    values=(
                                        'Jan',
                                        'Feb',
                                        'Mar',
                                        'Apr',
                                        'May',
                                        'Jun',
                                        'Jul',
                                        'Aug',
                                        'Sep',
                                        'Oct',
                                        'Nov',
                                        'Dec'
                                        ),
                                    size=4,
                                    readonly=True,
                                    key='FROM_MONTH_COMBO',
                                    enable_events=True
                                    )
                                ],
                            [
                                sg.Combo(
                                    values=(
                                        'Jan',
                                        'Feb',
                                        'Mar',
                                        'Apr',
                                        'May',
                                        'Jun',
                                        'Jul',
                                        'Aug',
                                        'Sep',
                                        'Oct',
                                        'Nov',
                                        'Dec'
                                        ),
                                    size=4,
                                    readonly=True,
                                    key='TO_MONTH_COMBO',
                                    enable_events=True
                                    )
                                ]
                            ]
                        )
                    ],
                [
                    sg.Column(
                        layout=[
                            [
                                sg.Checkbox(
                                    text='Suppress daily plots',
                                    key='SUPPRESS_DAILY_PLOTS'
                                    )
                                ]
                            ]
                        )
                    ]
                ]
            )
        tab1 = sg.Tab(
            title='Daily simulation mode',
            layout=[[frame1]],
            element_justification='center',
            key='DAILY_TAB',
            )
        tab2 = sg.Tab(
            title='Custom period simulation mode',
            layout=[[frame2]],
            element_justification='center',
            key='CUSTOM_TAB',
            )
        tab = sg.TabGroup(
            layout=[[tab1, tab2]],
            key='INIT_TAB',
            tab_background_color='Dark gray',
            enable_events=True
            )
        frame = sg.Column(
            layout=[
                [
                    sg.Frame(
                        title='Initialization parameters',
                        title_location='n',
                        layout = [
                            [
                                sg.Text(
                                    text='Initial temperature:',
                                    size=15
                                    ),
                                sg.Input(
                                    size=10,
                                    enable_events=True,
                                    key='INIT_TEMP_VALUE'
                                    ),
                                sg.Text(
                                    text='\u2070C',
                                    )
                                ],
                            [tab]
                            ],
                        element_justification='center'
                        )
                    ]
                ],
            element_justification = 'center'
            )
        return frame
           
    def LoadClimateFile(self):
        self.SimulatedBuilding.fl = self.values['CLIMATE_FILE_PATH'][:]
    
    def OccupancyScheduleBar(self):
        Col_list = []
        for i in range(49):
            if (i%2)==0:
                upper_element = sg.Text(
                    text=' '*(len(str(i//2))==1)+str(i//2),
                    font=('Helvetica',10),
                    pad=0
                    )
                lower_element = sg.Text(
                    text='',
                    pad=0
                    )
            else:
                upper_element = sg.Text(
                    text='',
                    pad=0
                    )
                lower_element = sg.Button(
                    button_color='Dark Red',
                    key="OCC_SCHED_BUTTON_"+str(i//2),
                    size=1,
                    pad=0,
                    )
            Col = sg.Column(
                layout= [
                    [upper_element],
                    [lower_element]
                    ],
                pad=0,
                element_justification='center'
                )
            Col_list.append(Col)
        two_texts = [
            sg.Text(
                text='Occupied',
                background_color='Green'
                ),
            sg.Text(
                text='Unoccupied',
                background_color='Dark Red'
                )
            ]
        frame = sg.Frame(
            title='Occupancy schedule',
            title_location = 'n',
            layout = [
                Col_list,
                two_texts
                ],
            element_justification='center',
            )
        col_final = sg.Column(layout=[[frame]],justification='center') 
        return col_final
    
    def CopyOccupancyToHeating(self):
        for i in range(24): 
            new_color = self.Window['OCC_SCHED_BUTTON_'+str(i)].ButtonColor[1]
            self.Window['HC_SCHED_BUTTON_'+str(i)].update(button_color=new_color)
        self.UpdateSchedules()
    
    def OccupancyScheduleTab(self):
        frame = sg.Column(
            layout=[
                [
                    sg.Frame(
                        title = 'Scheduling settings',
                        title_location = 'n',
                        layout= [
                            [
                                self.OccupancyScheduleBar()
                                ],
                            [
                                self.HCScheduleBar()
                                ],
                            [
                                self.FreshAirSupply()
                                ],
                            [
                                sg.Column(
                                    layout=[
                                        [
                                            sg.Text(
                                                text='No. of occupants:',
                                                size=20
                                                ),
                                            ],
                                        [
                                            sg.Text(
                                                text=('Max heating/'
                                                      +'cooling power:'),
                                                size=20
                                                ),
                                            ],
                                        ]
                                    ),
                                sg.Column(
                                    layout=[
                                        [
                                            sg.Input(
                                                size=10,
                                                enable_events=True,
                                                key='NO_OCCUPANT_VALUE'
                                                ),
                                            ],
                                        [
                                            sg.Input(
                                                size=10,
                                                enable_events=True,
                                                key='MAX_HC_VALUE'
                                                ),
                                            ],
                                        ]
                                    ),
                                sg.Column(
                                    layout=[
                                        [
                                            sg.Text(
                                                text='',
                                                )
                                            ],
                                        [
                                            sg.Text(
                                                text='W ',
                                                )
                                            ],         
                                        ]
                                    ),
                                sg.Column(
                                    layout=[
                                        [
                                            sg.Text(
                                                text='Weekly operation pattern',
                                                key='WEEKLY_OCC_TEXT',
                                                visible=False
                                                )
                                            ],
                                        [
                                            sg.Combo(
                                                values=[
                                                    'Whole week',
                                                    '6/7 days of week',
                                                    '5/7 days of week'
                                                    ],
                                                default_value='Whole week',
                                                key='WEEKLY_OCC_PATTERN',
                                                readonly=True,
                                                visible=False
                                                )
                                            ]
                                        ]
                                    ) 
                                ]
                            ]
                        )
                    ]
                ]
            )
        return frame
    
    def OpenFile(self):
        openfile = sg.popup_get_file(
            'Hey!', 
            no_window=True,
            file_types=[
                ('Excel files (.xlsx)','*.xlsx'),
                ]
            )
        General_df = None
        Opaque_df = None
        Frame_df = None
        if openfile != '':
            General_df = pd.read_excel(
                io=openfile,
                sheet_name='General info'
                )
            Opaque_df = pd.read_excel(
                io=openfile,
                sheet_name='Opaque list'
                )
            Frame_df = pd.read_excel(
                io=openfile,
                sheet_name='Frame list'
                )
            # ---- Change values that are invalid
            Opaque_df['Orientation'] =  Opaque_df[
                'Orientation'
                ].replace(
                    to_replace=np.float64('nan'),
                    value='None'
                    )
            check_str = Opaque_df['Attached frame indices'].apply(type)==str
            if check_str.any():
                At_fr_list = Opaque_df[
                'Attached frame indices'
                ].str.split(pat=',').to_list()
            else:
                At_fr_list = Opaque_df[
                'Attached frame indices'
                ].to_list()
            for j in range(len(At_fr_list)):
                if type(At_fr_list[j]) == list:
                    At_fr_list[j] = pd.to_numeric(
                        At_fr_list[j], 
                        errors='coerce', 
                        downcast='integer'
                        ).tolist()
                elif (
                        type(At_fr_list[j]) == float 
                        and not pd.isna(At_fr_list[j])
                        ):
                    At_fr_list[j] = [int(At_fr_list[j])]
                else:  
                    At_fr_list[j] = ['None']
            # ---- Reset Building and properties
            self.HidePropertiesDetails()
            self.SimulatedBuilding = Building()
            self.LoadClimateFile()
            # ---- Get the opaque and frame characteristics
            self.OpaqueList = Opaque_df['Element name'].to_list()
            self.FrameList = Frame_df['Element name'].to_list()
            for i in range(len(Frame_df['Element name'])):
                self.SimulatedBuilding.AddElement(
                    attr=Frame_df['Attribute'][i],
                    l=Frame_df['Length'][i],
                    h=Frame_df['Height/Width'][i],
                    o=Frame_df['Orientation'][i],
                    f=Frame_df['Shading factor'][i],
                    ul=Frame_df['Heat loss factor'][i],
                    g=Frame_df['Glazing factor'][i],
                    ffr=Frame_df['Frame factor'][i]
                    )
            for i in range(len(Opaque_df['Element name'])):
                self.SimulatedBuilding.AddElement(
                    attr=Opaque_df['Attribute'][i],
                    l=Opaque_df['Length'][i],
                    h=Opaque_df['Height/Width'][i],
                    o=Opaque_df['Orientation'][i],
                    f=Opaque_df['Shading factor'][i],
                    ul=Opaque_df['Heat loss factor'][i],
                    a=Opaque_df['Absorptivity factor'][i],
                    r=Opaque_df['External resistance'][i],
                    )
            # ---- Fix frame attachments
            self.FrameListAttachedto = ['None'] * len(self.FrameList)
            for j in range(len(At_fr_list)):
                for k in At_fr_list[j]:
                    if k != 'None':
                        self.SimulatedBuilding.AssignFrametoOpaque(
                            k, 
                            j
                            )
                        self.FrameListAttachedto[k] = self.OpaqueList[j]
                    self.FramesonOpaques.append(At_fr_list[j])
            self.Window['OPAQUE_LIST'].update(self.OpaqueList)
            self.Window['FRAME_LIST'].update(self.FrameList)
            self.ChangeAttachedToValues()
            # ---- Get the general characteristics ----
            self.Window[
                'INPUT_WEIGHT_CATEGORY'
                ].update(value= General_df['Weight Category'][0])
            self.SimulatedBuilding.Type = General_df['Weight Category'][0]
            self.SimulatedBuilding.SetBldType()
            self.values[
                'INPUT_WEIGHT_CATEGORY'
                ] = General_df['Weight Category'][0]
            # ----
            self.Window[
                'BUILDING_LATITUDE'
                ].update(
                    value= General_df['Location latitude (⁰)'][0]
                    )
            self.SimulatedBuilding.Latitude = General_df[
                'Location latitude (⁰)'
                ][0]
            self.values[
                'BUILDING_LATITUDE'
                ] = General_df['Location latitude (⁰)'][0]
        return General_df, Opaque_df, Frame_df
    
    def PropertiesDetails(self):
        self.PropFeaturesList = [
            'Orientation',
            'Length',
            'Height/Width',
            'Heat loss factor',
            'Shading factor',
            'Attached to',
            'External resistance',
            'Absorptivity factor',
            'Glazing factor',
            'Frame factor'
            ]
        len_list=[]
        for i in self.PropFeaturesList:
            len_list.append(len(i))
        self.max_length_features = max(len_list)
        self.PropFeaturesUnits = [
            ' ',
            'm',
            'm',
            'W/(m\u00b2·K)',
            ' ',
            ' ',
            '(m\u00b2·K)/W',
            ' ',
            ' ',
            ' '
            ]
        len_list=[]
        for i in self.PropFeaturesUnits:
            len_list.append(len(i))
        self.max_length_units = max(len_list)
        i = list(range(len(self.PropFeaturesList)))
        j = i[:]
        k = i[:]
        j.remove(5)
        j.remove(8)
        j.remove(9)
        k.remove(6)
        k.remove(7)
        prop_det1 = []
        prop_det2 = []
        prop_det3 = []
        for i1 in i:
            feature_title = self.PropFeaturesList[i1]
            unit = self.PropFeaturesUnits[i1]
            prop_det1.append(self.CreatePropertyDefault(feature_title,unit))
        for j1 in j:
            feature_title = self.PropFeaturesList[j1]
            unit = self.PropFeaturesUnits[j1]
            prop_det2.append(self.CreatePropertyOpaque(feature_title,unit))
        for k1 in k:
            feature_title = self.PropFeaturesList[k1]
            unit = self.PropFeaturesUnits[k1]
            prop_det3.append(self.CreatePropertyFrame(feature_title,unit))
        prop_row_d = sg.Column(prop_det1, key='COLUMN_DEFAULT')
        prop_row_o = sg.Column(prop_det2, key='COLUMN_OPAQUE', visible=False)
        prop_row_f = sg.Column(prop_det3, key='COLUMN_FRAME', visible=False)
        return [prop_row_d, prop_row_o, prop_row_f]
    
    def PropertiesTab(self):
        prop = sg.Column(
            layout=[
                [
                    sg.Frame(
                        title='Properties of selected element',
                        title_location='n',
                        layout=[
                            [
                                sg.Text(
                                    text='Selected Element:', 
                                    key='ELEMENT_GETTER_TITLE',
                                    )
                                ],
                            self.PropertiesDetails()
                            ],
                        element_justification='center',
                        size=(500,320)
                        )
                    ]
                ]
            )
        return prop
    
    def RevertPropertiesButtons(self, prop, of_lstbox):
            if (
                    (
                        of_lstbox == 'f' and (
                            prop != 'Orientation' and 
                            prop != 'Absorptivity factor' and
                            prop != 'External resistance'
                            )
                        ) or 
                    (
                        of_lstbox == 'o' and(
                            prop != 'Glazing factor' and
                            prop != 'Frame factor' and
                            prop != 'Attached to'
                            )
                        )
                    ):
                self.Window[prop+'_SETTER_'+of_lstbox].update(
                    visible=False
                    )
                self.Window[prop+'_CONFIRM_'+of_lstbox].update(
                    visible=False
                    )
                self.Window[prop+'_CANCEL_'+of_lstbox].update(
                    visible=False
                    )
                self.Window[prop+'_CHANGE_'+of_lstbox].update(
                    text='Change Value',
                    visible=True
                )
    
    def Run(self):
        self.Update()
        tabby = None
        while self.event != sg.WIN_CLOSED:
            if ((self.event == 'OPAQUE_LIST' 
                  and len(self.values['OPAQUE_LIST'])>0)
                  or (self.event == 'FRAME_LIST' 
                  and len(self.values['FRAME_LIST'])>0)
                ):
                self.ShowListboxProperties(self.event)
            elif self.event == 'ADD':
                self.AddElementtoLists()
            elif self.event == 'INPUT_WEIGHT_CATEGORY':
                self.SetBldWeightCategory(self.values['INPUT_WEIGHT_CATEGORY'])
            elif self.event == 'BUILDING_LATITUDE':
                self.SetBldLatitude(self.values['BUILDING_LATITUDE'])
            elif self.event[-9:-1] == '_CHANGE_':
                self.SetPropertyValue()
            elif (self.event[-10:-1] == '_CONFIRM_' 
                  or self.event[-9:-1] == '_CANCEL_'
                  ):
                self.ConfirmCancelPropertyValue()
            elif self.event == 'CLIMATE_FILE_PATH':
                self.LoadClimateFile()
            elif self.event == 'INIT_TAB':
                self.Window['WEEKLY_OCC_PATTERN'].update(
                    visible=self.values['INIT_TAB'] == 'CUSTOM_TAB'
                    )
                self.Window['WEEKLY_OCC_TEXT'].update(
                    visible=self.values['INIT_TAB'] == 'CUSTOM_TAB'
                    )
            elif self.event[:9]=='OCC_SCHED' or self.event[:8]=='HC_SCHED':
                self.ConfigureSchedules()
            elif self.event == 'HCcopiesOccupancy':
                self.CopyOccupancyToHeating()
            elif self.event[:9] == 'FRESH_AIR':
                self.SetAirSupply()
            elif self.event == 'SAVE_PLOT_PATH':
                self.path = self.values['SAVE_PLOT_PATH']
                excel_name = 'Temperatures & Thermal loads'
                self.excel_namepath = self.path+'/'+excel_name+'.xlsx'
            elif self.event == 'DATA_EXCEL_OUTPUT':
                self.Window['DATA_TICKMARK_TEXT'].update(
                    visible=self.values['DATA_EXCEL_OUTPUT']
                    )
                self.Window['DATA_TICKMARK_MINUTES'].update(
                    visible=self.values['DATA_EXCEL_OUTPUT']
                    )
            elif self.event == 'SIMULATE_BUTTON':
                if os.path.isfile(self.excel_namepath):
                    try:
                        os.remove(self.excel_namepath)
                    except:
                        pass
                if self.values['INIT_TAB'] == 'DAILY_TAB':
                    self.RunSimDaily()
                else:
                    self.RunSimCustom()
            elif self.event == 'Save building to Excel file...':
                tabby = self.SaveFile()
            elif self.event == 'Load building from Excel file...':
                tabby = self.OpenFile()
            self.SimulationReadyCheck()
            self.DisableAddButton()
            self.Update()
        self.Terminate()
        return self.SimulatedBuilding, tabby
    
    def RunSimCustom(self):
        day = int(self.values['FROM_DAY_INPUT'])
        day_f = int(self.values['TO_DAY_INPUT'])
        (self.month_st, self.month_f) = self.GetMonthsCustom() 
        self.weekly_cnt = 0
        month = self.month_st
        Current_DoY = self.SimulatedBuilding.DoYCalc(month, day)
        Last_DoY = self.SimulatedBuilding.DoYCalc(self.month_f, day_f)
        bld_tinit = None
        self.Eheat_monthList = []
        self.Ecool_monthList = []
        current_month = month
        self.Eheat_monthList.append(0)
        self.Ecool_monthList.append(0)
        fin = 0
        while Current_DoY <= Last_DoY and not(fin):
            if current_month != month:
                self.Eheat_monthList[-1] = round(self.Eheat_monthList[-1],2)
                self.Eheat_monthList.append(0)
                self.Ecool_monthList[-1] = round(self.Ecool_monthList[-1],2)
                self.Ecool_monthList.append(0)
                current_month = month
            self.RunSimDaily(day, month, bld_tinit)
            self.Eheat_monthList[-1] += (
                self.Eheat/self.SimulatedBuilding.FloorArea
                )
            self.Ecool_monthList[-1] += (
                self.Ecool/self.SimulatedBuilding.FloorArea
                )
            [day, month] = self.SimulatedBuilding.NextDay(month=month, day=day)
            Current_DoY = self.SimulatedBuilding.DoYCalc(month, day)
            fin = Current_DoY == 1
            bld_tinit = self.SimulatedBuilding.tm1
            self.weekly_cnt = (self.weekly_cnt+1)*(self.weekly_cnt<6) + 0*(self.weekly_cnt>=6)
        self.weekly_cnt = 8
        self.CreatePlots(day, month, finished=1)
        self.SavePlots()
        
    def RunSimDaily(self, day=None, month=None, bld_tinit=None):
        if day == None: day = int(self.values['DAY_VALUE'])
        if month == None: month = int(self.values['MONTH_VALUE']) 
        date_str = str(day)+'/'+str(month)
        self.SimulatedBuilding.ResetClock(
            Month=month,
            Day=day
            )
        self.SimulatedBuilding.ResetLists()
        if bld_tinit == None:
            bld_tinit = float(self.values['INIT_TEMP_VALUE'])
        self.SimulatedBuilding.Temp_iv = bld_tinit
        self.SimulatedBuilding.tm1 = self.SimulatedBuilding.Temp_iv
        self.SimulatedBuilding.Tairprev = bld_tinit
        self.SimulatedBuilding.Tsprev = bld_tinit
        HC_Sched, Occ_Sched, Vdotlist =self.SetWeeklySchedule()
        self.SimulatedBuilding.VdotList_handvalues = Vdotlist
        self.SimulatedBuilding.InitParamsISO(
            month=month,
            day=day,
            HC_schedule=HC_Sched,
            people_schedule=Occ_Sched,
            people_num = [
                int(self.values['NO_OCCUPANT_VALUE'])
                ]*len(self.OccupancySchedule),
            HC_maxload = [
                float(self.values['MAX_HC_VALUE'])
                ]*len(self.HeatCoolSchedule),
            Tsetpoints = [
                [
                    float(self.values['HEAT_SET_VALUE'])
                    ]*len(self.HeatCoolSchedule),
                [
                    float(self.values['COOL_SET_VALUE'])
                    ]*len(self.HeatCoolSchedule),
                ]                                                                     
            )
        self.SimulatedBuilding.RunSim()
        self.CreatePlots(day, month)
        self.SavePlots()
        if self.values['DATA_EXCEL_OUTPUT']:
            try:
                self.SaveLoadsExcel(date_str)
            except:
                pass
    
    def SaveFile(self):
        savefile = sg.popup_get_file(
            message='Hey!', 
            no_window=True, 
            save_as=True,
            file_types=[
                ('Excel files (.xlsx)','*.xlsx'),
                ]
            )
        # General building data
        gen_list = {
            'Weight Category':self.SimulatedBuilding.Type,
            'Location latitude (\u2070)':self.SimulatedBuilding.Latitude
            }
        General_df = pd.DataFrame(
            data=gen_list, 
            index = [0]
            )
        # Opaque List
        opq_clm_list = [
            'Element name',
            'Attribute',
            'Orientation',
            'Length',
            'Height/Width',
            'Heat loss factor',
            'Shading factor',
            'External resistance',
            'Absorptivity factor',
            'Attached frame indices'
            ]
        opqlist_values = []
        # print(opqlist_values)
        for i in range(len(self.SimulatedBuilding.OpaqueList)):
            y = self.SimulatedBuilding.OpaqueList[i]
            temp_str=''
            if len(y.AttachedFrames_indx) > 0:
                temp_str = str(y.AttachedFrames_indx[0])
                for x in y.AttachedFrames_indx[1:]:
                    temp_str += ',' + str(x)
            opqlist_values.append(
                [
                    self.OpaqueList[i],
                    y.Attr,
                    y.Orientation,
                    y.Length,
                    y.Height,
                    y.U,
                    y.Fsh,
                    y.ExternalRes,
                    y.abso,
                    temp_str
                    ]
                )
        Opaque_df = pd.DataFrame(data=opqlist_values, columns=opq_clm_list)
        # Frame List
        frm_clm_list = [
            'Element name',
            'Attribute',
            'Orientation',
            'Length',
            'Height/Width',
            'Heat loss factor',
            'Shading factor',
            'Glazing factor',
            'Frame factor'
            ]
        frmlist_values = []
        for i in range(len(self.SimulatedBuilding.FrameList)):
            y = self.SimulatedBuilding.FrameList[i]
            frmlist_values.append(
                [
                    self.FrameList[i],
                    y.Attr,
                    y.Orientation,
                    y.Length,
                    y.Height,
                    y.U,
                    y.Fsh,
                    y.ggl,
                    y.Ff
                    ]
                )
        Frame_df = pd.DataFrame(data=frmlist_values, columns=frm_clm_list)
        if savefile != '':
            with pd.ExcelWriter(savefile) as writer:
                General_df.to_excel(
                    excel_writer=writer,
                    sheet_name='General info',
                    index=False
                    )
                Opaque_df.to_excel(
                    excel_writer=writer,
                    sheet_name="Opaque list"
                    )
                Frame_df.to_excel(
                    excel_writer=writer,
                    sheet_name="Frame list"
                    )
        return self.SimulatedBuilding.Type, Opaque_df, Frame_df
    
    def SaveLoadsExcel(self, date_string = '1/1'):
        cols = [
            'Indoor air temperature [oC]',
            'Outdoor air temperature [oC]',
            'Required heating (+)/cooling(-) load [W]',
            'Thermal conductivity heat loss(+)/gains(-) [W]',
            'Ventilation loss(+)/gains(-) [W]',
            'Solar gains [W]',
            'Internal gains [W]'
            ]
        data_pd = pd.DataFrame(
            columns = cols
            )
        time = []
        self.tm_mins = 1
        if self.values['DATA_TICKMARK_MINUTES'] == 'per 1 hour':
            self.tm_mins = 60
        for i in range(86400//(self.SimulatedBuilding.dt*self.tm_mins)):
            minutes_in_day = i*self.SimulatedBuilding.dt//60*self.tm_mins
            time.append(
                '{:02d}:{:02d}'.format(
                    *divmod(
                        minutes_in_day, 60
                        )
                    )
                )
        date = [date_string] * len(time)
        datetime = list(zip(date,time))
        mindex = pd.MultiIndex.from_tuples(datetime, names=['date', 'time'])
        # read all values; temperatures are instantaneous
        TairList = []
        TheList = []
        # loads are summed per intervals (e.g. 1 hour contains hourly sum)
        QhvacList = [] #0
        FhlossList = [] #1
        FnatvelossList = [] #2
        FsolList = [] #3
        FintList = [] #4
        cnt = 0
        for i in range(len(self.SimulatedBuilding.QhvacList)):
            self.sum_list[0] += self.SimulatedBuilding.QhvacList[i]
            self.sum_list[1] += self.SimulatedBuilding.FhlossList[i]
            self.sum_list[2] += self.SimulatedBuilding.FnatvelossList[i]
            self.sum_list[3] += self.SimulatedBuilding.FsolList[i]
            self.sum_list[4] += self.SimulatedBuilding.FintList[i]
            got_it = 0
            cnt += 1
            if cnt == self.tm_mins:
                cnt = 0
                got_it = 1
                # ---- thermal loads
                QhvacList.append(self.sum_list[0]/self.tm_mins)
                FhlossList.append(self.sum_list[1]/self.tm_mins)
                FnatvelossList.append(self.sum_list[2]/self.tm_mins)
                FsolList.append(self.sum_list[3]/self.tm_mins)
                FintList.append(self.sum_list[4]/self.tm_mins)
                self.sum_list = [0] * 5   
            if i%self.tm_mins == 0:
                # ---- air temperatures (indoor, outdoor)
                TairList.append(self.SimulatedBuilding.TairList[i])
                TheList.append(self.SimulatedBuilding.TheList[i])
        if not got_it:
             # ---- thermal loads
             QhvacList.append(self.sum_list[0]/self.tm_mins)
             FhlossList.append(self.sum_list[1]/self.tm_mins)
             FnatvelossList.append(self.sum_list[2]/self.tm_mins)
             FsolList.append(self.sum_list[3]/self.tm_mins)
             FintList.append(self.sum_list[4]/self.tm_mins)
             self.sum_list = [0] * 5
        data_dict = {
            cols[0]: TairList,
            cols[1]: TheList,
            cols[2]: QhvacList,
            cols[3]: FhlossList,
            cols[4]: FnatvelossList,
            cols[5]: FsolList,
            cols[6]: FintList
            }
        data_pd = pd.DataFrame(data_dict, index=mindex)
        headers = True
        mode_chosen = 'w'
        ise = None
        first_time = 1
        if os.path.isfile(self.excel_namepath):
            headers = False
            mode_chosen = 'a'
            ise = 'overlay'
            first_time = 0
        writer = pd.ExcelWriter(
                self.excel_namepath,
                engine='openpyxl',
                mode=mode_chosen,
                if_sheet_exists=ise
                )
        if first_time:
            strow = 0
        else:
            strow = writer.sheets['Sheet1'].max_row
        data_pd.to_excel(
            excel_writer=writer,
            sheet_name='Sheet1',
            float_format='%.2f',
            header=headers,
            startrow=strow,
            index=True,
            merge_cells=True
            )
        writer.close()
    
    def SavePlots(self):
        for i in range(len(self.FigList)):
            self.plot_cnt += 1
            self.FigList[i].savefig(
                fname=self.path+'/SimFile_'+str(self.plot_cnt)+'.png',
                dpi=1200
                )
        self.FigList = []
        self.AxList = []
    
    def SavePlotTab(self):
        frame = sg.Frame(
            title='Saved plot & data location',
            title_location='n',
            layout=[
                [
                    sg.Column(
                        layout = [
                            [
                                sg.Input(
                                    default_text='',
                                    key='SAVE_PLOT_PATH',
                                    size=(40,5),
                                    enable_events=True,
                                    disabled=True
                                    ),
                                sg.FolderBrowse(
                                    button_text='Open Folder...',
                                    key='SAVE_PLOT_PATH_SEARCH',
                                    ),
                                ],
                            [
                                sg.Checkbox(
                                    text='Output results to excel file',
                                    key='DATA_EXCEL_OUTPUT',
                                    enable_events=True,
                                    ),
                                ],
                            [
                                sg.Text(
                                    text= 'Data tick Marks: ',
                                    key='DATA_TICKMARK_TEXT',
                                    visible=False
                                    ),
                                sg.Combo(
                                    values = [
                                        'per 1 minute',
                                        'per 1 hour'
                                        ],
                                    default_value = 'per 1 minute',
                                    key='DATA_TICKMARK_MINUTES',
                                    readonly=True,
                                    visible=False
                                    )
                                ]
                            ]
                        )
                    ]
                ]
            )
        return frame
    
    def SetAirSupply(self):
        indx = int(self.event.split('_')[-1])
        # print('The value is '+self.values[self.event])
        try:
            self.VdotList[indx] = float(self.values[self.event])
        except ValueError:
            self.VdotList[indx] = None
        # print(self.VdotList)
        

    def SetBldLatitude(self, Latitude):
        self.SimulatedBuilding.Latitude = Latitude
    
    def SetBldWeightCategory(self, WeightCategory):
        self.SimulatedBuilding.Type = WeightCategory
        self.SimulatedBuilding.SetBldType()
    
    def SetPropertyValue(self):
        prop = self.event[:-9]
        of_listbox = self.event[-1]
        self.Window[prop+'_VALUE_'+of_listbox].update(visible=False)
        self.Window[prop+'_CHANGE_'+of_listbox].update(text='', visible=False)
        self.Window[prop+'_SETTER_'+of_listbox].update(visible=True)
        self.Window[prop+'_CONFIRM_'+of_listbox].update(visible=True)
        self.Window[prop+'_CANCEL_'+of_listbox].update(visible=True) 
    
    def SetWeeklySchedule(self):
        WeekPatternOn = [0] * 7
        if self.values['WEEKLY_OCC_PATTERN'] == 'Whole week':
            WeekPatternOn = [1] * 7
        elif self.values['WEEKLY_OCC_PATTERN'] == '6/7 days of week':
            WeekPatternOn = [1,1,1,1,1,1,0]
        elif self.values['WEEKLY_OCC_PATTERN'] == '5/7 days of week':
            WeekPatternOn = [1,1,1,1,1,0,0]
        HCschedday = [[0, 0]]
        OccSchedday = [[0, 0]]
        Vdotlist = [0] * len(self.VdotList)
        if WeekPatternOn[self.weekly_cnt]:
            HCschedday = self.HeatCoolSchedule
            OccSchedday = self.OccupancySchedule
            Vdotlist = self.VdotList
        return HCschedday, OccSchedday, Vdotlist
    
    def ShowListboxProperties(self, listbox_key):
        self.SelectedItem = self.values[listbox_key][0]
        self.Window['ELEMENT_GETTER_TITLE'].update(
            'Selected Element: '
            + str(self.SelectedItem)
            )
        self.ShowPropertiesDetails() 
    
    def ShowPropertiesDetails(self):
        self.Window['COLUMN_DEFAULT'].update(visible=False)
        self.Window['COLUMN_OPAQUE'].update(
            visible=(self.event == 'OPAQUE_LIST')
            )
        self.Window['COLUMN_FRAME'].update(
            visible=(self.event == 'FRAME_LIST')
            )
        self.sel_indx = self.Window[self.event].Values.index(
            self.values[self.event][0]
            )
        if self.event == 'OPAQUE_LIST':
            self.sel_bld_elem = self.SimulatedBuilding.OpaqueList[
                self.sel_indx
                ]
            of_listbox='o'
            self.Window['External resistance_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.ExternalRes))
                        ) 
                    + str(self.sel_bld_elem.ExternalRes)
                ),
                visible=True
                )
            self.Window['Absorptivity factor_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.abso))
                        ) 
                    + str(self.sel_bld_elem.abso)
                ),
                visible=True
                )
        else:
            self.sel_bld_elem = self.SimulatedBuilding.FrameList[
                self.sel_indx
                ]
            of_listbox='f'
            self.Window['Glazing factor_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.ggl))
                        ) 
                    + str(self.sel_bld_elem.ggl)
                ),
                visible=True
                )
            self.Window['Frame factor_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(str(self.sel_bld_elem.Ff))
                        ) 
                    + str(self.sel_bld_elem.Ff)
                ),
                visible=True
                )
            self.Window['Attached to_VALUE_'+of_listbox].update(
                value=(
                    ' '*(
                        self.max_length_units
                        -len(self.FrameListAttachedto[self.sel_indx])
                        ) 
                    + self.FrameListAttachedto[self.sel_indx]
                ),
                visible=True
                )
            self.Window['Orientation_CHANGE_'+of_listbox].update(
                visible=False
                )
        self.Window['Orientation_VALUE_'+of_listbox].update(
            value=(
                ' '*(
                    self.max_length_units
                    -len(str(self.sel_bld_elem.Orientation))
                    ) 
                + str(self.sel_bld_elem.Orientation)
            ),
            visible=True
            )
        self.RevertPropertiesButtons('Orientation', of_listbox)
        self.Window['Length_VALUE_'+of_listbox].update(
            value=(
                ' '*(
                    self.max_length_units
                    -len(str(self.sel_bld_elem.Length))
                    ) 
                + str(self.sel_bld_elem.Length)
            ),
            visible=True
            )
        self.RevertPropertiesButtons('Length', of_listbox)
        self.Window['Height/Width_VALUE_'+of_listbox].update(
            value=(
                ' '*(
                    self.max_length_units
                    -len(str(self.sel_bld_elem.Height))
                    ) 
                + str(self.sel_bld_elem.Height)
            ),
            visible=True
            )
        self.RevertPropertiesButtons('Height/Width', of_listbox)
        self.Window['Heat loss factor_VALUE_'+of_listbox].update(
            value=(
                ' '*(
                    self.max_length_units
                    -len(str(self.sel_bld_elem.U))
                    ) 
                + str(self.sel_bld_elem.U)
            ),
            visible=True
            )
        self.RevertPropertiesButtons('Heat loss factor', of_listbox)
        self.Window['Shading factor_VALUE_'+of_listbox].update(
            value=(
                ' '*(
                    self.max_length_units
                    -len(str(self.sel_bld_elem.Fsh))
                    ) 
                + str(self.sel_bld_elem.Fsh)
            ),
            visible=True
            )    
        self.RevertPropertiesButtons('Shading factor', of_listbox)
        self.RevertPropertiesButtons('Absorptivity factor', of_listbox)
        self.RevertPropertiesButtons('External resistance', of_listbox)
        self.RevertPropertiesButtons('Glazing factor', of_listbox)
        self.RevertPropertiesButtons('Frame factor', of_listbox)
        self.RevertPropertiesButtons('Attached to', of_listbox)
    
    def SimulationReadyCheck(self):
        self.Ready[0] = self.values['INPUT_WEIGHT_CATEGORY'] != 'None'
        self.Ready[1] = self.CheckReady1()
        self.Ready[2] = self.CheckReady2()
        self.Ready[3] = self.CheckReady3()
        self.Ready[4] = self.values['CLIMATE_FILE_PATH'] != ''
        if self.values['INIT_TAB'] == 'DAILY_TAB':
            self.Ready[5] = (
                self.values['INIT_TEMP_VALUE'] != ''
                and self.values['DAY_VALUE'] != ''
                and self.values['MONTH_VALUE'] != ''
                )
        else:
            self.Ready[5] = (
                self.values['INIT_TEMP_VALUE'] != ''
                and self.values['FROM_MONTH_COMBO'] != ''
                and self.values['TO_MONTH_COMBO'] != ''
                and self.values['FROM_DAY_INPUT'] != ''
                and self.values['TO_DAY_INPUT'] != ''
                )
        self.Ready[6] = (
            self.values['HEAT_SET_VALUE'] != ''
            and self.values['COOL_SET_VALUE'] != ''
            )
        self.Ready[7] = (
            self.values['NO_OCCUPANT_VALUE'] != ''
            and self.values['MAX_HC_VALUE'] != ''
            )
        self.Ready[8] = self.values['SAVE_PLOT_PATH'] != ''
        self.Window['SIMULATE_BUTTON'].update(
            disabled=(len(self.Ready)!=sum(self.Ready))
            )
        self.Ready[9] = self.values['BUILDING_LATITUDE'] != ''
        all_text = ''
        for i in range(len(self.NoSimReasonsList)):
            all_text += self.NoSimReasonsList[i] * (not(self.Ready[i]))
        all_text=all_text[:-1]
        self.Window['SIM?_TITLE'].update(
            visible=(len(self.Ready)!=sum(self.Ready))
            )
        self.Window['SIM?_REASONS'].update(
            value=all_text,
            visible=(len(self.Ready)!=sum(self.Ready))
            )
    
    def SimulationTab(self):
        all_text = ''
        for i in range(len(self.NoSimReasonsList)):
            if i != 2 and i != 3:
                all_text += self.NoSimReasonsList[i] * (not(self.Ready[i]))
        all_text=all_text[:-1]
        frame = sg.Column(
            layout = [
                [
                    sg.Frame(
                        title='',
                        layout=[
                            [
                                sg.Button(
                                    button_text='Simulate!',
                                    key='SIMULATE_BUTTON',
                                    disabled=True
                                    )
                                ],
                            [
                                sg.Text(
                                    text=(
                                        'Cannot initiate simulation '
                                        + 'due to '
                                        + 'the following reasons:'
                                        ),
                                    font=('Helvetica', 10, 'bold'),
                                    key='SIM?_TITLE'
                                    )
                                ],
                            [
                                sg.Text(
                                    text=(all_text),                  
                                    font=('Helvetica', 10),
                                    key='SIM?_REASONS',
                                    background_color='Dark Red'
                                    )
                                ]
                            ],
                        element_justification='center'
                        )    
                    ]
                ]
            )
        return frame
    
    def TopMenu(self):
        menu_def = [
            [
                'File', 
                ['Save building to Excel file...', 'Load building from Excel file...']
                ],
            ]
        menu = sg.Menu(menu_definition=menu_def)
        return menu
    
    def ThermostatSettingsTab(self):
        frame = sg.Column(
            layout=[
                [
                    sg.Frame(
                        title='Thermostat settings',
                        title_location='n',
                        layout=[
                            [
                                sg.Column(
                                    layout=[
                                        [
                                            sg.Text(
                                                text='Heating Setpoint:',
                                                size=15
                                                ),
                                            ],
                                        [
                                            sg.Text(
                                                text='Cooling Setpoint:',
                                                size=15
                                                ),
                                            ]
                                        ]
                                    ),
                                sg.Column(
                                    layout=[
                                        [
                                            sg.Input(
                                                size=10,
                                                enable_events=True,
                                                key='HEAT_SET_VALUE'
                                                ),
                                            ],
                                        [
                                            sg.Input(
                                                size=10,
                                                enable_events=True,
                                                key='COOL_SET_VALUE'
                                                ),
                                            ],
                                        [
                                            
                                            ]
                                        ]
                                    ),
                                sg.Column(
                                    layout=[
                                        [
                                            sg.Text(
                                                text='\u2070C   '
                                                )
                                            ],
                                        [
                                            sg.Text(
                                                text='\u2070C   '
                                                )
                                            ]                                            
                                        ]
                                    ),
                                ]
                            ]
                        )
                    ]
                ]
            )
        return frame
    
    def Terminate(self):
        self.Window.close()
    
    def Update(self):
        # ---- Operation
        self.event, self.values = self.Window.read()
        print(self.event,'\n', self.values)
    
    def UpdateSchedules(self):
        self.OccupancySchedule = []
        self.HeatCoolSchedule = []
        for i in range(24):
            target_color_occ = self.Window['OCC_SCHED_BUTTON_'+str(i)].ButtonColor[1]
            if target_color_occ == 'green':
                self.OccupancySchedule.append([i*3600, (i+1)*3600])
            target_color_hc = self.Window['HC_SCHED_BUTTON_'+str(i)].ButtonColor[1]
            if target_color_hc == 'green':
                self.HeatCoolSchedule.append([i*3600, (i+1)*3600])
    
    def ValuePropertyFrame(self, keytext):
        prop = sg.Input(
            key=keytext+'_SETTER_f', 
            visible=False,
            size=10,
            do_not_clear=False
            )
        if keytext == 'Orientation':
            prop = sg.Combo(
                values=[
                    'None',
                    'North',
                    'South',
                    'East',
                    'West'
                ],
                default_value='None',
                key=keytext+'_SETTER_f',
                readonly=True,
                size=9,
                visible=False
                )
        elif keytext == 'Attached to':
            opaque_list_temp = ['None']
            for i in self.OpaqueList:
                opaque_list_temp.append(i)
            prop = sg.Combo(
                values = opaque_list_temp,
                default_value='None',
                key=keytext+'_SETTER_f',
                readonly=True,
                size=9,
                visible=False,
                enable_events=True
                )
        return prop
    
    def ValuePropertyOpaque(self, keytext):
        prop = sg.Input(
            key=keytext+'_SETTER_o', 
            visible=False,
            size=10,
            do_not_clear=False
            )
        if keytext == 'Orientation':
            prop = sg.Combo(
                values=[
                    'None',
                    'North',
                    'South',
                    'East',
                    'West'
                ],
                default_value='None',
                key=keytext+'_SETTER_o',
                readonly=True,
                size=9,
                visible=False
                )
        return prop

        
if __name__ == "__main__":
    sim = Simulator()
    (FinalBuilding, saved_file) = sim.Run()