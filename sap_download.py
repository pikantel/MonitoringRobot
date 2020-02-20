import win32com.client
from win32com import *
from win32api import *
from win32com.client import *
import datetime
import json
from prepare_data import clear_path

#-Includes--------------------------------------------------------------
from sys import *

#-Sub Main--------------------------------------------------------------

list_of_partners = ['X5312_OMPT', 'X4114_OMPT', 'X4300_OMPT', 'X4545_OMPT',
                    'X4116_OMPT', 'X4695_OMPT', 'X4674_OMPT', 'X5005_OMPT',
                    'Y4180_OMPT', 'X2258_FSF', 'X2258_OOH', 'X1295_UBF',
                    'X2378_OMPT', 'X4337_OMPT', 'X4300_OMPT', 'X2599_OMPT',
                    'X9540_OMPT', 'X9532_OMPT', 'X2862_OMPT']


def sap_connect():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return
    return session


def read_from_txt():
    f = open(r'C:\Temp\output\SAP\data.txt', "r")
    lines = f.readlines()
    result = []
    for x in lines[8:]:
        result.append(x.split('\t')[1])
    f.close()
    return result


class SapDownload:

    def __init__(self, date_from, date_to, start_hour, finish_hour, path):
        self.date_from = date_from
        self.date_to = date_to
        self.start_hour = start_hour
        self.finish_hour = finish_hour
        self.path = path

    def download_data(self):
        session = sap_connect()
        session.StartTransaction("SQ00")
        if session.findById("wnd[0]/usr/txtRS38R-WSTEXT").text != "Standard Area (Client-specific)":
            session.findById("wnd[0]/mbar/menu[5]/menu[0]").Select()
            session.findById("wnd[1]/usr/radRAD1").Select()
            session.findById("wnd[1]/tbar[0]/btn[2]").Press()

        if session.findById("wnd[0]/titl").text != "Query from User Group " + "O2C_MACRO" + ": Initial Screen":
            session.findById("wnd[0]/mbar/menu[1]/menu[7]").Select()
            session.findById("wnd[1]/tbar[0]/btn[29]").Press()
            session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").Press()
            session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").Press()
            session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "O2C_MACRO"
            session.findById("wnd[3]/tbar[0]/btn[0]").Press()
            session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
            session.findById("wnd[1]/tbar[0]/btn[0]").Press()

        session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "OMPT_MONITOR"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = self.date_from
        session.findById("wnd[0]/usr/ctxtSP$00002-HIGH").text = self.date_to
        session.findById("wnd[0]/usr/ctxtSP$00003-LOW").text = self.start_hour
        session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").text = self.finish_hour
        session.findById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").press()
        j = 0
        for i in range(len(list_of_partners)):  # add idoc partners
            try:
                session.findById(
                    f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                    f"ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                    f"txtRSCSEL_255-SLOW_I[1,{i}]").text = list_of_partners[i]

            # while the next line of the list in SAP is not visible, scolldown once

            except:
                j += 1
                session.findById(
                    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                    "ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = j
                session.findById(
                    f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                    f"ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                    f"txtRSCSEL_255-SLOW_I[1,{i-j}]").text = list_of_partners[i]

        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/txtSP$00001-LOW").caretPosition = 9
        session.findById("wnd[0]/tbar[1]/btn[8]").Press()
        session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&PC")
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.path
        clear_path(self.path)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = 'data.txt'
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

    def download_from_edid4(self):
        session = sap_connect()
        session.StartTransaction("se16n")
        session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EDID4"
        session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").columns.elementAt(4).width = 5
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()

        j = 0
        list_of_idocs = read_from_txt()
        num_of_rounds = int(len(list_of_idocs))/5

        while num_of_rounds > 0:
            session.findById("wnd[1]/tbar[0]/btn[7]").press()
            num_of_rounds -= 1

        for i in range(len(list_of_idocs)):
            try:
                session.findById(f"wnd[1]/usr/tblSAPLSE16NMULTI_TC/"
                                 f"ctxtGS_MULTI_SELECT-LOW[1,{i}]").text = list_of_idocs[i]

            # while the next line of the list in SAP is not visible, scolldown once

            except:
                j += 1
                session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC").verticalScrollbar.position = j
                session.findById(f"wnd[1]/usr/tblSAPLSE16NMULTI_TC/"
                                 f"ctxtGS_MULTI_SELECT-LOW[1,{i-j}]").text = list_of_idocs[i]

        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").text = "E1EDK02"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "data2.txt"
        session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 18
        session.findById("wnd[1]/tbar[0]/btn[0]").press()



