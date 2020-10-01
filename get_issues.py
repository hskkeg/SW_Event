from redminelib import Redmine
from src.List import *
from openpyxl import Workbook
from openpyxl.styles import Font, Side, Border

# single line comment
"""
  Multi lines comment
"""

"""
    09.22   22896 - [BR][IDB] HMC AE 21MY M 등록
    09.22   30344 - [IDB] TML Kanger2.0 T-car 用
    09.22   29350 - [IDB] WEIMA APE-5 Proto 用
    09.22   30554 - [100KR] SEM N7 DV S/W 用
    09.22   17123 - [BR] [IDB] JK1 내수 SOP Event
    09.24   29768 - [IDB] FAW N117 Proto (Mando Internal)
    09.24   29769 - [IDB] NIO Force Proto (Mando Internal)
    09.25   30607 - [BR] [100CN] SV51-01 GB6(MoC Si Base/Prime) SOP R/C
    09.22   29387 - [RCU] WEIMA APE-5 Proto 用
"""
ED4_WORK_LIST = [22896, 30344, 29350, 30554, 17123, 29768, 29769, 30607, 29387]
REG_EVENT_LIST = ['AE 21MY M','TML Kanger2.0 T-car','IDB APE-5 Proto','N7 DV','JK1 SOP','N117 Proto','Force Proto','SV51-01 GB6 RC','RCU APE-5 Proto']

def PM_Redmine_Issue_List(Input):
    redmine = Redmine('http://191.1.11.178', username='sk.hahm', password='dbsguswls22@')
    ED4_TitleList = []
    ED4_MemberName = []
    ED4_Tracker_List = []
    ED4_URL_List = []
    s = redmine.issue.get(Input)

    for i in s.children:
        if ('Mandatory' in redmine.issue.get(i.id).tracker.name) | ('Category' in redmine.issue.get(i.id).tracker.name) | ('Carry_Over' in redmine.issue.get(i.id).tracker.name):
            i = redmine.issue.get(i.id)
            for j in i.children:
                try:
                    Name = redmine.issue.get(j.id).assigned_to.name
                    Title = j.subject
                    URL = j.url
                    Tracker = str(redmine.issue.get(i.id).tracker)

                    #if ('/' in Tracker):
                        #Tracker = Tracker.replace("/", "")

                    for Change in ED4_Find_Member():
                        if Name in Change[0]:
                            ED4_TitleList.append(Title)
                            ED4_MemberName.append(Change[-1])
                            ED4_Tracker_List.append(Tracker)
                            ED4_URL_List.append(URL)
                            print(URL)
                except:
                    pass

        else:
            try:
                Name = redmine.issue.get(i.id).assigned_to.name
                Title = i.subject
                URL = i.url
                Tracker = str(redmine.issue.get(i.id).tracker)

                #if ('/' in Tracker):
                    #Tracker = Tracker.replace("/", "")


                for Change in ED4_Find_Member():
                    if Name in Change[0]:
                        ED4_TitleList.append(Title)
                        ED4_MemberName.append(Change[-1])
                        ED4_Tracker_List.append(Tracker)
                        ED4_URL_List.append(URL)
                        print(URL)
            except:
                pass

    for idx in range(len(ED4_WORK_LIST)):
        print(idx)

        if(Input == ED4_WORK_LIST[idx]):
            WSNum = 'ws'+str(idx)
            WSNum = wb.create_sheet()
            WSNum.title = str(REG_EVENT_LIST[idx])
            Export_Excel(WSNum, ED4_TitleList, ED4_MemberName, ED4_Tracker_List, ED4_URL_List)

def ED4_Member():
    ED4member = []
    for i in ED4_MemberList():
        ED4member.append(i.split('/'))
    return ED4member

def ED4_Find_Member():
    ED4findmember = []
    for j in ED4_Member():
        ED4findmember.append(j)
    return ED4findmember

def ED4_Work_List_CW39():
    for j in range(len(ED4_WORK_LIST)):
        PM_Redmine_Issue_List(ED4_WORK_LIST[j])

def Export_Excel(SheetNum, TiltleList, MemberName, Tracker, URL):
    SheetNum.cell(1, 1, '일감')
    SheetNum.cell(1, 2, '담당자')
    SheetNum.cell(1, 3, '유형')
    SheetNum.cell(1, 4, '레드마인 주소')
    for j in range(len(TiltleList)):
        SheetNum.cell(j + 2, 1, TiltleList[j])
        SheetNum.cell(j + 2, 2, MemberName[j])
        SheetNum.cell(row=j + 2, column=3).value = Tracker[j]
        SheetNum.cell(j + 2, 4, URL[j])
    #SheetNum.column_dimensions.auto_size = True
    SheetNum.column_dimensions['A'].width = 95
    SheetNum.column_dimensions['B'].width = 9
    SheetNum.column_dimensions['C'].width = 15
    SheetNum.column_dimensions['D'].width = 32


wb = Workbook()     # create work book
ED4_Work_List_CW39()
wb.save("CW39.xlsx")

# PM_Redmine_Issue_List(17111)