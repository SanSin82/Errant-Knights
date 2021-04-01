import myNotebook as nb
import sys
import json
import requests
from config import config
from theme import theme
import webbrowser
import os.path
import gspread
from gspread_formatting import *
from os import path

try:
    # Python 2
    import Tkinter as tk
    import ttk
except ModuleNotFoundError:
    # Python 3
    import tkinter as tk
    from tkinter import ttk

this = sys.modules[__name__]  # For holding module globals
this.VersionNo = "1"
this.TodayData = {}
this.DataIndex = 0
this.Status = "Active"
this.TickTime = ""
this.cred = ''  # google sheet service account cred's path to file
this.SystemFaction = tk.StringVar()
this.MasterPriority = tk.StringVar()
this.MasterFaction = tk.StringVar()
this.MasterWork = tk.StringVar()
this.MasterGoal = tk.StringVar()
this.MasterCZFaction = tk.StringVar()

style = ttk.Style()

style.theme_create('dark', settings={
    ".": {
        "configure": {
            "background": '#000000', # All except tabs
            "font": '#000000'
        }
    },
    "TNotebook": {
        "configure": {
            "background":'#000000', # Your margin color
            "tabmargins": [2, 5, 0, 0], # margins: left, top, right, separator
        }
    },
    "TNotebook.Tab": {
        "configure": {
            "background": '#ffffff', # tab color when not selected
            "padding": [10, 2], # [space between text and horizontal tab-button border, space between text and vertical tab_button border]
            "font":'#000000'
        },
        "map": {
            "background": [("selected", '#ffffff')], # Tab color when selected
            "expand": [("selected", [1, 1, 1, 0])] # text margins
        }
    }
})

def plugin_prefs(parent, cmdr, is_beta):
    """
   Return a TK Frame for adding to the EDMC settings dialog.
   """

    frame = nb.Frame(parent)
    nb.Label(frame, text="Errant Knights v" + this.VersionNo).grid(column=0, sticky=tk.W)
    """
   reset = nb.Button(frame, text="Reset Counter").place(x=0 , y=290)
   """
    nb.Checkbutton(frame, text="Make Errant Knights Active", variable=this.Status, onvalue="Active",
                   offvalue="Paused").grid()
    return frame

def plugin_start(plugin_dir):
    """
   Load this plugin into EDMC
   """
    this.Dir = plugin_dir
    this.cred = os.path.join(this.Dir, "service_account.json")
    file = os.path.join(this.Dir, "Today Data.txt")
    
    if path.exists(file):
        with open(file) as json_file:
            this.TodayData = json.load(json_file)
            z = len(this.TodayData)
            for i in range(1, z + 1):
                x = str(i)
                this.TodayData[i] = this.TodayData[x]
                del this.TodayData[x]

    this.LastTick = tk.StringVar(value=config.get("XLastTick"))
    this.TickTime = tk.StringVar(value=config.get("XTickTime"))
    this.Status = tk.StringVar(value=config.get("XStatus"))
    this.DataIndex = tk.IntVar(value=config.get("XIndex"))
    this.StationFaction = tk.StringVar(value=config.get("XStation"))

    # this.LastTick.set("12")

    #  tick check and counter reset
    response = requests.get('https://elitebgs.app/api/ebgs/v5/ticks')  # get current tick and reset if changed
    tick = response.json()
    this.CurrentTick = tick[0]['_id']
    this.TickTime = tick[0]['time']
    print(this.LastTick.get())
    print(this.CurrentTick)
    if this.LastTick.get() != this.CurrentTick:
        this.LastTick.set(this.CurrentTick)
        this.TodayData = {}
        print("Tick auto reset happened")
    # create google sheet
    google_sheet_int()

    return "Errant Knights v2"


def plugin_start3(plugin_dir):
    return plugin_start(plugin_dir)


def plugin_stop():
    """
    EDMC is closing
    """
    save_data()

    print("Farewell cruel world!")


def plugin_app(parent):
    """    Create a frame for the EDMC main window    """
    this.frame = tk.Frame(parent)

    title = tk.Label(this.frame, text="Errant Knights v" + this.VersionNo)
    title.grid(row=0, column=0, sticky=tk.W)

    tk.Button(this.frame, text='ORDERS', command=display_data).grid(row=0, column=1, padx=3)
    tk.Label(this.frame, text="Status:").grid(row=1, column=0, sticky=tk.W)
    tk.Label(this.frame, text="Last Tick:").grid(row=2, column=0, sticky=tk.W)
    this.Status_Label = tk.Label(this.frame, text=this.Status.get()).grid(row=1, column=1, sticky=tk.W)
    this.TimeLabel = tk.Label(this.frame, text=tick_format(this.TickTime)).grid(row=2, column=1, sticky=tk.W)
    tk.Label(this.frame, text="Controlling Faction:").grid(row=3, column=0, sticky=tk.W)
    this.Controlling_Label = tk.Label(this.frame, textvariable=this.SystemFaction).grid(row=3, column=1, sticky=tk.W)
    theme.update(this.frame)
    tk.Label(this.frame, text="Priority:").grid(row=4, column=0, sticky=tk.W)
    this.MasterPriority_Label = tk.Label(this.frame, textvariable=this.MasterPriority).grid(row=4, column=1, sticky=tk.W)
    theme.update(this.frame)
    tk.Label(this.frame, text="Faction:").grid(row=5, column=0, sticky=tk.W)
    this.MasterFaction_Label = tk.Label(this.frame, textvariable=this.MasterFaction).grid(row=5, column=1, sticky=tk.W)
    theme.update(this.frame)
    tk.Label(this.frame, text="Work:").grid(row=6, column=0, sticky=tk.W)
    this.MasterWork_Label = tk.Label(this.frame, textvariable=this.MasterWork).grid(row=6, column=1, sticky=tk.W)
    theme.update(this.frame)
    tk.Label(this.frame, text="Goal:").grid(row=7, column=0, sticky=tk.W)
    this.MasterGoal_Label = tk.Label(this.frame, textvariable=this.MasterGoal).grid(row=7, column=1, sticky=tk.W)
    theme.update(this.frame)
    tk.Label(this.frame, text="CZFaction:").grid(row=8, column=0, sticky=tk.W)
    this.MasterCZFaction_Label = tk.Label(this.frame, textvariable=this.MasterCZFaction).grid(row=8, column=1,
                                                                                            sticky=tk.W)
    theme.update(this.frame)
    tk.Button(this.frame, text='CZ HIGH', command=high_cz).grid(row=10, column=0, padx=3)
    tk.Button(this.frame, text='CZ MED', command=med_cz).grid(row=10, column=1, padx=3)
    tk.Button(this.frame, text='CZ LOW', command=low_cz).grid(row=10, column=2, padx=3)

    return this.frame


def journal_entry(cmdr, is_beta, system, station, entry, index):
    if this.Status.get() != "Active":
        print('Paused')
        return

    if entry['event'] == 'Docked':
        this.StationFaction.set(entry['StationFaction']['Name'])  # set station controlling faction name

        try:
            gc = gspread.service_account(filename=this.cred)
            sh = gc.open("DAILY UPDATE")
            worksheet = sh.worksheet("Orders")
            system = this.TodayData[this.DataIndex.get()][0]['System']
            cell1 = worksheet.find(system)
            systemrow = cell1.row
            pcell = worksheet.cell(systemrow, 2).value
            fcell = worksheet.cell(systemrow, 3).value
            wcell = worksheet.cell(systemrow, 4).value
            gcell = worksheet.cell(systemrow, 5).value
            czcell = worksheet.cell(systemrow, 6).value
            this.MasterPriority.set(pcell)
            this.MasterFaction.set(fcell)
            this.MasterWork.set(wcell)
            this.MasterGoal.set(gcell)
            this.MasterCZFaction.set(czcell)
        except gspread.exceptions.CellNotFound:
            this.MasterPriority.set('NONE')
            this.MasterFaction.set('NONE')
            this.MasterWork.set('NONE')
            this.MasterGoal.set('NONE')
            this.MasterCZFaction.set('NONE')

    if entry['event'] == 'Location':
        try:
            if this.SystemFaction != entry['SystemFaction']['Name']:
                this.SystemFaction.set(entry['SystemFaction']['Name'])
        except KeyError:
            this.SystemFaction.set('Unpopulated')

        #  tick check and counter reset
        response = requests.get('https://elitebgs.app/api/ebgs/v5/ticks')  # get current tick and reset if changed
        tick = response.json()
        this.CurrentTick = tick[0]['_id']
        this.TickTime = tick[0]['time']
        print(this.TickTime)
        if this.LastTick.get() != this.CurrentTick:
            this.LastTick.set(this.CurrentTick)
            this.TodayData = {}
            this.TimeLabel = tk.Label(this.frame, text=tick_format(this.TickTime)).grid(row=2, column=1, sticky=tk.W)
            theme.update(this.frame)
            print("Tick auto reset happened")
            # set up new sheet at tick reset
        google_sheet_int()
        # today data creation
        x = len(this.TodayData)
        try:
            if x >= 1:
                for y in range(1, x + 1):
                    if entry['StarSystem'] == this.TodayData[y][0]['System']:
                        this.DataIndex.set(y)
                        print('system in data')
                        sheet_insert_new_system(y)
                        return
                this.TodayData[x + 1] = [{'System': entry['StarSystem'], 'SystemAddress': entry['SystemAddress'],
                                          'Factions': []}]
                this.DataIndex.set(x + 1)

                for i in entry['Factions']:
                    if i['Name'] != "Pilots' Federation Local Branch":
                        inf = i['Influence'] * 100
                        state = ''
                        pendingstate = ''
                        try:
                            for z in i['ActiveStates']:
                                state = state + z['State'] + ' '
                        except KeyError:
                            state = 'None'
                        try:
                            for z in i['PendingStates']:
                                pendingstate = pendingstate + z['State'] + ' '
                        except KeyError:
                            pendingstate = 'None'

                        this.TodayData[x + 1][0]['Factions'].append({'Faction': i['Name'], 'INF': inf, 'State': state,
                                                                     'PendingState': pendingstate, 'Bounties': 0,
                                                                     'Bonds': 0, 'TradeProfit': 0, 'BMProfit': 0,
                                                                     'MissionPoints': 0, 'MissionFailed': 0, 'CartData': 0,
                                                                     'Murders': 0, 'Fines&Bounties': 0, 'CZ High': 0,
                                                                     'CZ Med': 0,
                                                                     'CZ Low': 0})
            else:
                this.TodayData = {1: [{'System': entry['StarSystem'],
                                       'SystemAddress': entry['SystemAddress'], 'Factions': []}]}
                this.DataIndex.set(1)
                for i in entry['Factions']:
                    if i['Name'] != "Pilots' Federation Local Branch":
                        inf = i['Influence'] * 100
                        state = ''
                        pendingstate = ''
                        try:
                            for z in i['ActiveStates']:
                                state = state + z['State'] + ' '
                        except KeyError:
                            state = 'None'
                        try:
                            for z in i['PendingStates']:
                                pendingstate = pendingstate + z['State'] + ' '
                        except KeyError:
                            pendingstate = 'None'

                        this.TodayData[x + 1][0]['Factions'].append({'Faction': i['Name'], 'INF': inf, 'State': state,
                                                                     'PendingState': pendingstate, 'Bounties': 0,
                                                                     'Bonds': 0, 'TradeProfit': 0, 'BMProfit': 0,
                                                                     'MissionPoints': 0, 'MissionFailed': 0, 'CartData': 0,
                                                                     'Murders': 0, 'Fines&Bounties': 0, 'CZ High': 0,
                                                                     'CZ Med': 0,
                                                                     'CZ Low': 0})

            sheet_insert_new_system(x + 1)  # insert data into google sheet

            try:
                gc = gspread.service_account(filename=this.cred)
                sh = gc.open("DAILY UPDATE")
                worksheet = sh.worksheet("Orders")
                system = this.TodayData[this.DataIndex.get()][0]['System']
                cell1 = worksheet.find(system)
                systemrow = cell1.row
                pcell = worksheet.cell(systemrow, 2).value
                fcell = worksheet.cell(systemrow, 3).value
                wcell = worksheet.cell(systemrow, 4).value
                gcell = worksheet.cell(systemrow, 5).value
                czcell = worksheet.cell(systemrow, 6).value
                this.MasterPriority.set(pcell)
                this.MasterFaction.set(fcell)
                this.MasterWork.set(wcell)
                this.MasterGoal.set(gcell)
                this.MasterCZFaction.set(czcell)
            except gspread.exceptions.CellNotFound:
                this.MasterPriority.set('NONE')
                this.MasterFaction.set('NONE')
                this.MasterWork.set('NONE')
                this.MasterGoal.set('NONE')
                this.MasterCZFaction.set('NONE')

        except KeyError:
            try:
                gc = gspread.service_account(filename=this.cred)
                sh = gc.open("DAILY UPDATE")
                worksheet = sh.worksheet("Orders")
                system = this.TodayData[this.DataIndex.get()][0]['System']
                cell1 = worksheet.find(system)
                systemrow = cell1.row
                pcell = worksheet.cell(systemrow, 2).value
                fcell = worksheet.cell(systemrow, 3).value
                wcell = worksheet.cell(systemrow, 4).value
                gcell = worksheet.cell(systemrow, 5).value
                czcell = worksheet.cell(systemrow, 6).value
                this.MasterPriority.set(pcell)
                this.MasterFaction.set(fcell)
                this.MasterWork.set(wcell)
                this.MasterGoal.set(gcell)
                this.MasterCZFaction.set(czcell)
            except gspread.exceptions.CellNotFound:
                this.MasterPriority.set('NONE')
                this.MasterFaction.set('NONE')
                this.MasterWork.set('NONE')
                this.MasterGoal.set('NONE')
                this.MasterCZFaction.set('NONE')

    if entry['event'] == 'FSDJump':  # get factions at jump, load into today data, check tick and reset if needed

        try:
            if this.SystemFaction != entry['SystemFaction']['Name']:
                this.SystemFaction.set(entry['SystemFaction']['Name'])
        except KeyError:
            this.SystemFaction.set('Unpopulated')

        #  tick check and counter reset
        response = requests.get('https://elitebgs.app/api/ebgs/v5/ticks')  # get current tick and reset if changed
        tick = response.json()
        this.CurrentTick = tick[0]['_id']
        this.TickTime = tick[0]['time']
        print(this.TickTime)
        if this.LastTick.get() != this.CurrentTick:
            this.LastTick.set(this.CurrentTick)
            this.TodayData = {}
            this.TimeLabel = tk.Label(this.frame, text=tick_format(this.TickTime)).grid(row=2, column=1, sticky=tk.W)
            theme.update(this.frame)
            print("Tick auto reset happened")
            # set up new sheet at tick reset
        google_sheet_int()
        # today data creation
        x = len(this.TodayData)
        try:
            if x >= 1:
                for y in range(1, x + 1):
                    if entry['StarSystem'] == this.TodayData[y][0]['System']:
                        this.DataIndex.set(y)
                        print('system in data')
                        sheet_insert_new_system(y)
                        return
                this.TodayData[x + 1] = [{'System': entry['StarSystem'], 'SystemAddress': entry['SystemAddress'],
                                          'Factions': []}]
                this.DataIndex.set(x + 1)

                for i in entry['Factions']:
                    if i['Name'] != "Pilots' Federation Local Branch":
                        inf = i['Influence'] * 100
                        state = ''
                        pendingstate = ''
                        try:
                            for z in i['ActiveStates']:
                                state = state + z['State'] + ' '
                        except KeyError:
                            state = 'None'
                        try:
                            for z in i['PendingStates']:
                                pendingstate = pendingstate + z['State'] + ' '
                        except KeyError:
                            pendingstate = 'None'

                        this.TodayData[x + 1][0]['Factions'].append({'Faction': i['Name'], 'INF': inf, 'State': state,
                                                                     'PendingState': pendingstate, 'Bounties': 0,
                                                                     'Bonds': 0, 'TradeProfit': 0, 'BMProfit': 0,
                                                                     'MissionPoints': 0, 'MissionFailed': 0, 'CartData': 0,
                                                                     'Murders': 0, 'Fines&Bounties': 0, 'CZ High': 0,
                                                                     'CZ Med': 0,
                                                                     'CZ Low': 0})
            else:
                this.TodayData = {1: [{'System': entry['StarSystem'],
                                       'SystemAddress': entry['SystemAddress'], 'Factions': []}]}
                this.DataIndex.set(1)
                for i in entry['Factions']:
                    if i['Name'] != "Pilots' Federation Local Branch":
                        inf = i['Influence'] * 100
                        state = ''
                        pendingstate = ''
                        try:
                            for z in i['ActiveStates']:
                                state = state + z['State'] + ' '
                        except KeyError:
                            state = 'None'
                        try:
                            for z in i['PendingStates']:
                                pendingstate = pendingstate + z['State'] + ' '
                        except KeyError:
                            pendingstate = 'None'

                        this.TodayData[x + 1][0]['Factions'].append({'Faction': i['Name'], 'INF': inf, 'State': state,
                                                                     'PendingState': pendingstate, 'Bounties': 0,
                                                                     'Bonds': 0, 'TradeProfit': 0, 'BMProfit': 0,
                                                                     'MissionPoints': 0, 'MissionFailed': 0, 'CartData': 0,
                                                                     'Murders': 0, 'Fines&Bounties': 0, 'CZ High': 0,
                                                                     'CZ Med': 0,
                                                                     'CZ Low': 0})

            sheet_insert_new_system(x + 1)  # insert data into google sheet

            try:
                gc = gspread.service_account(filename=this.cred)
                sh = gc.open("DAILY UPDATE")
                worksheet = sh.worksheet("Orders")
                system = this.TodayData[this.DataIndex.get()][0]['System']
                cell1 = worksheet.find(system)
                systemrow = cell1.row
                pcell = worksheet.cell(systemrow, 2).value
                fcell = worksheet.cell(systemrow, 3).value
                wcell = worksheet.cell(systemrow, 4).value
                gcell = worksheet.cell(systemrow, 5).value
                czcell = worksheet.cell(systemrow, 6).value
                this.MasterPriority.set(pcell)
                this.MasterFaction.set(fcell)
                this.MasterWork.set(wcell)
                this.MasterGoal.set(gcell)
                this.MasterCZFaction.set(czcell)
            except gspread.exceptions.CellNotFound:
                this.MasterPriority.set('NONE')
                this.MasterFaction.set('NONE')
                this.MasterWork.set('NONE')
                this.MasterGoal.set('NONE')
                this.MasterCZFaction.set('NONE')

        except KeyError:
            try:
                gc = gspread.service_account(filename=this.cred)
                sh = gc.open("DAILY UPDATE")
                worksheet = sh.worksheet("Orders")
                system = this.TodayData[this.DataIndex.get()][0]['System']
                cell1 = worksheet.find(system)
                systemrow = cell1.row
                pcell = worksheet.cell(systemrow, 2).value
                fcell = worksheet.cell(systemrow, 3).value
                wcell = worksheet.cell(systemrow, 4).value
                gcell = worksheet.cell(systemrow, 5).value
                czcell = worksheet.cell(systemrow, 6).value
                this.MasterPriority.set(pcell)
                this.MasterFaction.set(fcell)
                this.MasterWork.set(wcell)
                this.MasterGoal.set(gcell)
                this.MasterCZFaction.set(czcell)
            except gspread.exceptions.CellNotFound:
                this.MasterPriority.set('NONE')
                this.MasterFaction.set('NONE')
                this.MasterWork.set('NONE')
                this.MasterGoal.set('NONE')
                this.MasterCZFaction.set('NONE')

    if entry['event'] == 'RedeemVoucher' and entry['Type'] == 'bounty':  # bounties collected
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in entry['Factions']:
            for x in range(0, t):
                if z['Faction'] == this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Faction']:
                    this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Bounties'] += z['Amount']
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = x
                    data = z['Amount']
                    sheet_commit_data(system, index, 'Bounty', data)
        save_data()

    if entry['event'] == 'RedeemVoucher' and entry['Type'] == 'bond':  # bonds collected
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in entry['Factions']:
            for x in range(0, t):
                if z['Faction'] == this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Faction']:
                    this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Bonds'] += z['Amount']
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = x
                    data = z['Amount']
                    sheet_commit_data(system, index, 'Bonds', data)
        save_data()

    try:
        if entry['event'] == 'MarketSell':  # bmTrade Profit
            t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
            for z in range(0, t):
                if entry['BlackMarket']:
                    if this.StationFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                        cost = entry['Count'] * entry['AvgPricePaid']
                        bmprofit = entry['TotalSale'] - cost
                        this.TodayData[this.DataIndex.get()][0]['Factions'][z]['BMProfit'] += bmprofit
                        system = this.TodayData[this.DataIndex.get()][0]['System']
                        index = z
                        data = bmprofit
                        sheet_commit_data(system, index, 'BMTrade', data)
            save_data()
    except KeyError:
        if entry['event'] == 'MarketSell':  # bmTrade Profit
            t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
            for z in range(0, t):
                if this.StationFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                    cost = entry['Count'] * entry['AvgPricePaid']
                    profit = entry['TotalSale'] - cost
                    this.TodayData[this.DataIndex.get()][0]['Factions'][z]['TradeProfit'] += profit
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = z
                    data = profit
                    sheet_commit_data(system, index, 'Trade', data)
            save_data()

    if entry['event'] == 'MissionCompleted':  # get mission influence value
        fe = entry['FactionEffects']
        print("mission completed")
        for i in fe:
            fe3 = i['Faction']
            print(fe3)
            fe4 = i['Influence']
            for x in fe4:
                fe6 = x['SystemAddress']
                inf = len(x['Influence'])
                for y in this.TodayData:
                    if fe6 == this.TodayData[y][0]['SystemAddress']:
                        t = len(this.TodayData[y][0]['Factions'])
                        system = this.TodayData[y][0]['System']
                        for z in range(0, t):
                            if fe3 == this.TodayData[y][0]['Factions'][z]['Faction']:
                                this.TodayData[y][0]['Factions'][z]['MissionPoints'] += inf
                                sheet_commit_data(system, z, 'Mission', inf)
            save_data()
        try:
            gc = gspread.service_account(filename=this.cred)
            sh = gc.open("DAILY UPDATE")
            worksheet = sh.worksheet("Mission")
            id = str(entry['MissionID'])
            cell1 = worksheet.find(id)
            missionrow = cell1.row
            worksheet.delete_row(missionrow)
        except:
            gspread.exceptions.CellNotFound

    if entry['event'] == 'SellExplorationData' or entry['event'] == "MultiSellExplorationData":  # get carto data value
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in range(0, t):
            if this.StationFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                this.TodayData[this.DataIndex.get()][0]['Factions'][z]['CartData'] += entry['TotalEarnings']
                system = this.TodayData[this.DataIndex.get()][0]['System']
                index = z
                data = entry['TotalEarnings']
                sheet_commit_data(system, index, 'Expo', data)
        save_data()

    if entry['event'] == 'CommitCrime' and entry['CrimeType'] == 'murder':  # crime murder needs tested
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in range(0, t):
            if entry['Faction'] == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Murders'] += 1
                system = this.TodayData[this.DataIndex.get()][0]['System']
                index = z
                data = 1
                sheet_commit_data(system, index, 'Murders', data)
        save_data()

    try:
        if entry['event'] == 'CommitCrime':  # bounties collected
            t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
            for z in range(0, t):
                if entry['Faction'] == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                    this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Fines&Bounties'] += entry['Bounty']
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = z
                    data = entry['Bounty']
                    sheet_commit_data(system, index, 'Fines&Bounties', data)
            save_data()
    except KeyError:
        if entry['event'] == 'CommitCrime':  # bounties collected
            t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
            for z in range(0, t):
                if entry['Faction'] == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                    this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Fines&Bounties'] += entry['Fine']
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = z
                    data = entry['Fine']
                    sheet_commit_data(system, index, 'Fines&Bounties', data)
            save_data()

    if entry['event'] == 'MissionAccepted':  # get mission influence value
        gc = gspread.service_account(filename=this.cred)
        sh = gc.open("DAILY UPDATE")
        worksheet = sh.worksheet("Mission")
        str_list = list(filter(None, worksheet.col_values(1)))
        next_row = (len(str_list)+1)
        id = entry['MissionID']
        system = this.TodayData[this.DataIndex.get()][0]['System']
        faction = entry['Faction']
        inf = len(entry['Influence'])
        worksheet.update_cell(next_row, 1, id)
        worksheet.update_cell(next_row, 2, system)
        worksheet.update_cell(next_row, 3, faction)
        worksheet.update_cell(next_row, 4, inf)

    if entry['event'] == 'MissionFailed':
        gc = gspread.service_account(filename=this.cred)
        sh = gc.open("DAILY UPDATE")
        worksheet = sh.worksheet("Mission")
        id = str(entry['MissionID'])
        cell1 = worksheet.find(id)
        missionrow = cell1.row
        scell = worksheet.cell(missionrow, 2).value
        fcell = worksheet.cell(missionrow, 3).value
        icell = worksheet.cell(missionrow, 4).value
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in range(0, t):
            if fcell == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                this.TodayData[this.DataIndex.get()][0]['Factions'][z]['MissionFailed'] += int(icell)
                system = scell
                index = z
                data = int(icell)
                sheet_commit_data(system, index, 'MissionFailed', data)
            save_data()
        worksheet.delete_row(missionrow)

    if entry['event'] == 'SupercruiseEntry':
        try:
            gc = gspread.service_account(filename=this.cred)
            sh = gc.open("DAILY UPDATE")
            worksheet = sh.worksheet("Orders")
            system = entry['StarSystem']
            cell1 = worksheet.find(system)
            systemrow = cell1.row
            pcell = worksheet.cell(systemrow, 2).value
            fcell = worksheet.cell(systemrow, 3).value
            wcell = worksheet.cell(systemrow, 4).value
            gcell = worksheet.cell(systemrow, 5).value
            czcell = worksheet.cell(systemrow, 6).value
            this.MasterPriority.set(pcell)
            this.MasterFaction.set(fcell)
            this.MasterWork.set(wcell)
            this.MasterGoal.set(gcell)
            this.MasterCZFaction.set(czcell)
        except gspread.exceptions.CellNotFound:
            this.MasterPriority.set('NONE')
            this.MasterFaction.set('NONE')
            this.MasterWork.set('NONE')
            this.MasterGoal.set('NONE')
            this.MasterCZFaction.set('NONE')

    if entry['event'] == 'SupercruiseExit':
        try:
            gc = gspread.service_account(filename=this.cred)
            sh = gc.open("DAILY UPDATE")
            worksheet = sh.worksheet("Orders")
            system = entry['StarSystem']
            cell1 = worksheet.find(system)
            systemrow = cell1.row
            pcell = worksheet.cell(systemrow, 2).value
            fcell = worksheet.cell(systemrow, 3).value
            wcell = worksheet.cell(systemrow, 4).value
            gcell = worksheet.cell(systemrow, 5).value
            czcell = worksheet.cell(systemrow, 6).value
            this.MasterPriority.set(pcell)
            this.MasterFaction.set(fcell)
            this.MasterWork.set(wcell)
            this.MasterGoal.set(gcell)
            this.MasterCZFaction.set(czcell)
        except gspread.exceptions.CellNotFound:
            this.MasterPriority.set('NONE')
            this.MasterFaction.set('NONE')
            this.MasterWork.set('NONE')
            this.MasterGoal.set('NONE')
            this.MasterCZFaction.set('NONE')


def high_cz():
    t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
    for z in range(0, t):
        if this.MasterCZFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
            this.TodayData[this.DataIndex.get()][0]['Factions'][z]['CZ High'] += 1
            system = this.TodayData[this.DataIndex.get()][0]['System']
            index = z
            data = 1
            sheet_commit_data(system, index, 'CZ High', data)
    save_data()


def med_cz():
    t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
    for z in range(0, t):
        if this.MasterCZFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
            this.TodayData[this.DataIndex.get()][0]['Factions'][z]['CZ Med'] += 1
            system = this.TodayData[this.DataIndex.get()][0]['System']
            index = z
            data = 1
            sheet_commit_data(system, index, 'CZ Med', data)
    save_data()


def low_cz():
    t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
    for z in range(0, t):
        if this.MasterCZFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
            this.TodayData[this.DataIndex.get()][0]['Factions'][z]['CZ Low'] += 1
            system = this.TodayData[this.DataIndex.get()][0]['System']
            index = z
            data = 1
            sheet_commit_data(system, index, 'CZ Low', data)
    save_data()


def version_tuple(version):
    try:
        ret = tuple(map(int, version.split(".")))
    except:
        ret = (0,)
    return ret


def tick_format(TickTime):
    datetime1 = TickTime.split('T')
    x = datetime1[0]
    z = datetime1[1]
    y = x.split('-')
    if y[1] == "01":
        month = "Jan"
    elif y[1] == "02":
        month = "Feb"
    elif y[1] == "03":
        month = "March"
    elif y[1] == "04":
        month = "April"
    elif y[1] == "05":
        month = "May"
    elif y[1] == "06":
        month = "June"
    elif y[1] == "07":
        month = "July"
    elif y[1] == "08":
        month = "Aug"
    elif y[1] == "09":
        month = "Sep"
    elif y[1] == "10":
        month = "Oct"
    elif y[1] == "11":
        month = "Nov"
    elif y[1] == "12":
        month = "Dec"
    date1 = y[2] + " " + month
    time1 = z[0:5]
    datetimetick = time1 + ' UTC ' + date1
    return datetimetick


def save_data():
    config.set('XLastTick', this.CurrentTick)
    config.set('XTickTime', this.TickTime)
    config.set('XStatus', this.Status.get())
    config.set('XIndex', this.DataIndex.get())
    config.set('XStation', this.StationFaction.get())

    file = os.path.join(this.Dir, "Today Data.txt")
    with open(file, 'w') as outfile:
        json.dump(this.TodayData, outfile)


def google_sheet_int():
    # start google sheet data store
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("DAILY UPDATE")
    try:
        worksheet = sh.worksheet("Today")
    except:
        worksheet = sh.add_worksheet(title="Today", rows="30000", cols="16")
        worksheet.update('A1', '# of Systems')
        worksheet.update('B1', 0)
        set_column_width(worksheet, 'A', 270)


def sheet_insert_new_system(index):
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("DAILY UPDATE")
    worksheet = sh.worksheet("Today")
    factionname = []
    factioninf = []
    factionstate = []
    factionpendingstate = []
    systemfaction = this.SystemFaction.get()
    system = this.TodayData[index][0]['System']
    mpriority = '=iferror(vlookup(indirect("b"&row(),true),Master!$A:$F,2,false),"NONE")'
    mfaction = '=iferror(vlookup(indirect("b"&row(),true),Master!$A:$F,3,false),"NONE")'
    mwork = '=iferror(vlookup(indirect("b"&row(),true),Master!$A:$F,4,false),"NONE")'
    mgoal = '=iferror(vlookup(indirect("b"&row(),true),Master!$A:$F,5,false),"NONE")'
    mcz = '=iferror(vlookup(indirect("b"&row(),true),Master!$A:$F,6,false),"NONE")'
    try:
        cell = worksheet.find(system)
    except gspread.exceptions.CellNotFound:
        z = len(this.TodayData[index][0]['Factions'])
        for x in range(0, z):
            factionname.append([this.TodayData[index][0]['Factions'][x]['Faction']])
            factioninf.append([this.TodayData[index][0]['Factions'][x]['INF']])
            factionstate.append([this.TodayData[index][0]['Factions'][x]['State']])
            factionpendingstate.append([this.TodayData[index][0]['Factions'][x]['PendingState']])
        no_of_systems = int(worksheet.acell('B1').value)
        if no_of_systems == 0:
            no_of_systems += 1
            worksheet.update('B1', no_of_systems)
            # worksheet.update('A2:P3', [['System', system],['Faction', 'Mission +', 'Trade', 'Bounties',
            # 'Carto Data']])
            worksheet.batch_update([{'range': 'A2:P3',
                                     'values': [['System', system, 'System Control', systemfaction,
                                                 'Priority', mpriority,
                                                 'Faction', mfaction,
                                                 'Work', mwork,
                                                 'Goal', mgoal,
                                                 'CZFaction', mcz],
                                                ['Faction', 'INF', 'State', 'PendingState',
                                                 'Bounties', 'Bonds', 'Mission +', 'Mission Failed',
                                                 'Trade', 'BMTrade', 'Carto Data', 'Murders',
                                                 'Fines&Bounties', 'CZ High', 'CZ Med', 'CZ Low']]},
                                    {'range': 'A4:A11', 'values': factionname},
                                    {'range': 'B4:B11', 'values': factioninf},
                                    {'range': 'C4:C11', 'values': factionstate},
                                    {'range': 'D4:D11', 'values': factionpendingstate},
                                    {'range': 'E4:P11',
                                     'values': [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]}],
                                   value_input_option='USER_ENTERED')
        else:
            row = no_of_systems * 11 + 2
            no_of_systems += 1
            worksheet.update('B1', no_of_systems)
            range1 = 'A' + str(row) + ':P' + str(row + 1)
            range2 = 'A' + str(row + 2) + ':A' + str(row + 10)
            range3 = 'B' + str(row + 2) + ':B' + str(row + 10)
            range4 = 'C' + str(row + 2) + ':C' + str(row + 10)
            range5 = 'D' + str(row + 2) + ':D' + str(row + 10)
            range6 = 'E' + str(row + 2) + ':P' + str(row + 10)
            worksheet.batch_update([{'range': range1,
                                     'values': [['System', system, 'System Control', systemfaction,
                                                 'Priority', mpriority,
                                                 'Faction', mfaction,
                                                 'Work', mwork,
                                                 'Goal', mgoal,
                                                 'CZFaction', mcz],
                                                ['Faction', 'INF', 'State', 'PendingState',
                                                 'Bounties', 'Bonds', 'Mission +', 'Mission Failed',
                                                 'Trade', 'BMTrade', 'Carto Data', 'Murders',
                                                 'Fines&Bounties', 'CZ High', 'CZ Med', 'CZ Low']]},
                                    {'range': range2, 'values': factionname},
                                    {'range': range3, 'values': factioninf},
                                    {'range': range4, 'values': factionstate},
                                    {'range': range5, 'values': factionpendingstate},
                                    {'range': range6,
                                     'values': [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]}],
                                   value_input_option='USER_ENTERED')


def sheet_commit_data(system, index, event, data):
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("DAILY UPDATE")
    worksheet = sh.worksheet("Today")
    cell1 = worksheet.find(system)
    factionrow = cell1.row + 2 + index
    if event == "INF":
        cell = worksheet.cell(factionrow, 2).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 2, total)

    if event == "State":
        cell = worksheet.cell(factionrow, 3).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 3, total)

    if event == "PendingState":
        cell = worksheet.cell(factionrow, 4).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 4, total)

    if event == "Bounty":
        cell = worksheet.cell(factionrow, 5).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 5, total)

    if event == "Bonds":
        cell = worksheet.cell(factionrow, 6).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 6, total)

    if event == "Mission":
        cell = worksheet.cell(factionrow, 7).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 7, total)

    if event == "MissionFailed":
        cell = worksheet.cell(factionrow, 8).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 8, total)

    if event == "Trade":
        cell = worksheet.cell(factionrow, 9).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 9, total)

    if event == "BMTrade":
        cell = worksheet.cell(factionrow, 10).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 10, total)

    if event == "Expo":
        cell = worksheet.cell(factionrow, 11).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 11, total)

    if event == "Murders":
        cell = worksheet.cell(factionrow, 12).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 12, total)

    if event == "Fines&Bounties":
        cell = worksheet.cell(factionrow, 13).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 13, total)

    if event == "CZ High":
        cell = worksheet.cell(factionrow, 14).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 14, total)

    if event == "CZ Med":
        cell = worksheet.cell(factionrow, 15).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 15, total)

    if event == "CZ Low":
        cell = worksheet.cell(factionrow, 16).value
        total = int(cell) + data
        worksheet.update_cell(factionrow, 16, total)


def display_data():
    color='#fa8100'
    form = tk.Toplevel(this.frame)
    form.title("Errant Knights Orders")
    style.theme_use('dark')
    form.geometry("900x900")
    tab_parent = ttk.Notebook(form)
    tab = ttk.Frame(tab_parent)
    tab_parent.add(tab, text="ORDERS")
    systemLabel = tk.Label(tab, text="System",fg=color, bg='black').grid(row=0, column=0)
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("DAILY UPDATE")
    worksheet = sh.worksheet("Orders")
    acell = worksheet.row_values(2)
    bcell = worksheet.row_values(3)
    ccell = worksheet.row_values(4)
    dcell = worksheet.row_values(5)
    ecell = worksheet.row_values(6)
    fcell = worksheet.row_values(7)
    gcell = worksheet.row_values(8)
    hcell = worksheet.row_values(9)
    icell = worksheet.row_values(10)
    jcell = worksheet.row_values(11)
    kcell = worksheet.row_values(12)
    lcell = worksheet.row_values(13)
    mcell = worksheet.row_values(14)
    ncell = worksheet.row_values(15)
    ocell = worksheet.row_values(16)
    pcell = worksheet.row_values(17)
    qcell = worksheet.row_values(18)
    rcell = worksheet.row_values(19)
    scell = worksheet.row_values(20)
    tcell = worksheet.row_values(21)
    ucell = worksheet.row_values(22)
    vcell = worksheet.row_values(23)
    wcell = worksheet.row_values(24)
    xcell = worksheet.row_values(25)
    ycell = worksheet.row_values(26)
    zcell = worksheet.row_values(27)

    system_a = tk.Label(tab, text= acell,fg=color, bg='black').grid(row=1, column=0, sticky=tk.E)
    system_b = tk.Label(tab, text= bcell,fg=color, bg='black').grid(row=2, column=0, sticky=tk.E)
    system_c = tk.Label(tab, text= ccell,fg=color, bg='black').grid(row=3, column=0, sticky=tk.E)
    system_d = tk.Label(tab, text= dcell,fg=color, bg='black').grid(row=4, column=0, sticky=tk.E)
    system_e = tk.Label(tab, text= ecell,fg=color, bg='black').grid(row=5, column=0, sticky=tk.E)
    system_f = tk.Label(tab, text= fcell,fg=color, bg='black').grid(row=6, column=0, sticky=tk.E)
    system_g = tk.Label(tab, text= gcell,fg=color, bg='black').grid(row=7, column=0, sticky=tk.E)
    system_h = tk.Label(tab, text= hcell,fg=color, bg='black').grid(row=8, column=0, sticky=tk.E)
    system_i = tk.Label(tab, text= icell,fg=color, bg='black').grid(row=9, column=0, sticky=tk.E)
    system_j = tk.Label(tab, text= jcell,fg=color, bg='black').grid(row=10, column=0, sticky=tk.E)
    system_k = tk.Label(tab, text= kcell,fg=color, bg='black').grid(row=11, column=0, sticky=tk.E)
    system_l = tk.Label(tab, text= lcell,fg=color, bg='black').grid(row=12, column=0, sticky=tk.E)
    system_m = tk.Label(tab, text= mcell,fg=color, bg='black').grid(row=13, column=0, sticky=tk.E)
    system_n = tk.Label(tab, text= ncell,fg=color, bg='black').grid(row=14, column=0, sticky=tk.E)
    system_o = tk.Label(tab, text= ocell,fg=color, bg='black').grid(row=15, column=0, sticky=tk.E)
    system_p = tk.Label(tab, text= pcell,fg=color, bg='black').grid(row=16, column=0, sticky=tk.E)
    system_q = tk.Label(tab, text= qcell,fg=color, bg='black').grid(row=17, column=0, sticky=tk.E)
    system_r = tk.Label(tab, text= rcell,fg=color, bg='black').grid(row=18, column=0, sticky=tk.E)
    system_s = tk.Label(tab, text= scell,fg=color, bg='black').grid(row=19, column=0, sticky=tk.E)
    system_t = tk.Label(tab, text= tcell,fg=color, bg='black').grid(row=20, column=0, sticky=tk.E)
    system_u = tk.Label(tab, text= ucell,fg=color, bg='black').grid(row=21, column=0, sticky=tk.E)
    system_v = tk.Label(tab, text= vcell,fg=color, bg='black').grid(row=22, column=0, sticky=tk.E)
    system_w = tk.Label(tab, text= wcell,fg=color, bg='black').grid(row=23, column=0, sticky=tk.E)
    system_x = tk.Label(tab, text= xcell,fg=color, bg='black').grid(row=24, column=0, sticky=tk.E)
    system_y = tk.Label(tab, text= ycell,fg=color, bg='black').grid(row=25, column=0, sticky=tk.E)
    system_z = tk.Label(tab, text= zcell,fg=color, bg='black').grid(row=26, column=0, sticky=tk.E)

    tab_parent.pack(expand=1, fill='both')
