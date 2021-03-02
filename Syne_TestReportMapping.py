import openpyxl
import datetime
import pandas as pd

def EmpId_Name_Mapping(inputExcel):
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    empIdNameMapping = {}
    for index, row in dfSyneExcel.iterrows():
        if isinstance(row[0], int):
            empIdNameMapping[str(row[0])]= str(row[1])
    return empIdNameMapping


def get_AllDates_FromSyne(inputExcel):
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    AllDates=[]
    for i in range(5, len(dfSyneExcel.columns)):
        AllDates.append(dfSyneExcel.values[2][i])
    return AllDates

def Resource_Task_TotalHour_mapping(inputExcel, EmpID):
    headerRow = ['Billable', 'Leave', 'Public Holiday', 'Training', 'Admin/Other']
    Syne_TaskHourMap = {}
    taskHour_dict = {}
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")

    for i in range(1, len(dfSyneExcel)):
        if(str(dfSyneExcel.values[i][0])==EmpID):
            taskName = dfSyneExcel.values[i][3]
            TotalHours = dfSyneExcel.values[i][4]
            taskHour_dict[taskName]= TotalHours
        Syne_TaskHourMap[EmpID] = taskHour_dict
    return Syne_TaskHourMap

def Syne_Date_Hours_Mapping(inputExcel, EmpID, Date):
    Syne_date_taskHourMap = {}
    TaskDayHour_dict = {}
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    dateColNo = 0
    for j in range(1, len(dfSyneExcel.columns)):
        if (str(dfSyneExcel.values[2][j]) == Date):
            dateColNo = j
            break
    for i in range(1, len(dfSyneExcel)):
        if (str(dfSyneExcel.values[i][0]) == EmpID):
            task = dfSyneExcel.values[i][3]
            hourValue = dfSyneExcel.values[i][dateColNo]
            TaskDayHour_dict[task] = hourValue
        Syne_date_taskHourMap[Date] = TaskDayHour_dict
    return Syne_date_taskHourMap




inputExcel = "C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\Syne Jan Timesheet.xlsx"
empId_NameMap = EmpId_Name_Mapping(inputExcel)
All_dates = get_AllDates_FromSyne(inputExcel)
EmpId_TaskHour = Resource_Task_TotalHour_mapping(inputExcel, '11')
Task_DateHour = Syne_Date_Hours_Mapping(inputExcel, '321', '06-JAN-2021')

print(empId_NameMap)
print(All_dates)
print(EmpId_TaskHour)
print(Task_DateHour)

