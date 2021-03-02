import openpyxl
import datetime
import pandas as pd

from Syne_TestReportMapping import Syne_Date_Hours_Mapping, Resource_Task_TotalHour_mapping, get_AllDates_FromSyne


def getWeekDay(SyneTimesheetDate):
    monthToNum={'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
    weekdayToName ={0:'Mon',1:'Tue',2:'Wed',3:'Thu',4:'Fri',5:'Sat',6:'Sun'}
    dateSplit = SyneTimesheetDate.split("-")
    int_day=int(dateSplit[0])
    monthSub = dateSplit[1][0:3];
    int_month = monthToNum[monthSub]
    int_Year= int(dateSplit[2])
    weekday = datetime.date(day=int_day, month=int_month, year=int_Year).weekday()
    return weekdayToName[weekday]

def Emp_count(str):
    counts = dict()
    words = str.split()
    for word in words:
        if word in counts:
            counts[word] += 1
        else:
            counts[word] = 1
    return counts


def TSRunner():
    inputExcel1 ="C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\Syne Jan Timesheet.xlsx"
    inputExcel2 ="C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\Client timesheet report daily.csv"
    outputExcel ="C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\output.xlsx"

    dfSyneExcel = pd.read_excel(inputExcel1,"new sheet")
    dfClientExcel = pd.read_csv(inputExcel2)
    TimesheetDetail = dfSyneExcel.columns[0]
    columncount = len(dfSyneExcel.columns)
    TotalDays = columncount-5
    headerRow= ['','EMP ID','RESOURCE','PROJECT','TASK','TOTAL']
    weakdayRow = ['', '', '', '', '', '']
    #create header Row
    for day in range(1,TotalDays+1):
        headerRow.append(str(day))
    #create Weekday Row
    for i in range(5,TotalDays+5):
        weakdayRow.append(getWeekDay(dfSyneExcel.iat[2, i]))
    #get all unique employee Id's
    print(dfSyneExcel[TimesheetDetail].unique())
    print(dfSyneExcel[TimesheetDetail].count())
    empId=""
    empIdNameMapping = {}
    for index, row in dfSyneExcel.iterrows():
        if isinstance(row[0], int):
            print(str(row[0])+" ---- "+row[1])
            syneReportDetails = {}
            syneReportDetails['ResourceName'] = row[1]
            syneReportDetails['Project'] = row[2]
            task = {}
            taskType = {}
            for i in range(3, len(dfSyneExcel)):
                attendance = {}
                if dfSyneExcel.loc[i][0]== row[0]:
                    attendance['total'] = dfSyneExcel.loc[i][4]
                    for day in range(5, TotalDays + 5):
                        attendance[day - 4] = dfSyneExcel.loc[i][day]
                    taskType[dfSyneExcel.loc[i][3]] = attendance
                syneReportDetails['Task'] = taskType
            empIdNameMapping[row[0]] = syneReportDetails
    print(empIdNameMapping)

    outputData = [['', TimesheetDetail], [''], [''], headerRow, weakdayRow]
    ListAllData = []
    for empId, empDetails in empIdNameMapping.items():
        l1=[]
        EmpId_TaskHour = Resource_Task_TotalHour_mapping(inputExcel1, str(empId))
        noOfTask = EmpId_TaskHour[str(empId)].keys()
        getAllDates= get_AllDates_FromSyne(inputExcel1)
        #count=0
        counter = 0
        for task_key in noOfTask:
            l2 = []
            if counter == 0:
                l2.append('Syne')
                l2.append(empId)
                l2.append(empDetails['ResourceName'])
                l2.append(empDetails['Project'])
                counter = counter + 1
            else:
                l2.append('')
                l2.append('')
                l2.append('')
                l2.append('')

            l2.append(task_key)
            l2.append(EmpId_TaskHour[str(empId)][task_key])
            for day in getAllDates:
                GetDate_Hour_Mapping = Syne_Date_Hours_Mapping(inputExcel1, str(empId), day)
                l2.append(GetDate_Hour_Mapping[day][task_key])
            outputData.append(l2)

        outputData.append(l1)


    Outputdf1 = pd.DataFrame(outputData)
    Outputdf1.to_excel("C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\output.xlsx",sheet_name='Sheet_name_1', header=False,index=False)



TSRunner()


