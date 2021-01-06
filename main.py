from helper import getEmployeesData, daysCurrentMonth, getEmployeesToCome, createExcelFromDictionary

# Data to fill
employeesCsvPath = 'C:\\Users\\Mariejo\\PycharmProjects\\pythonProject2\\employees_data.txt'
resultExcelPath = 'C:\\Users\\Mariejo\\Desktop\\attendanceSheet1.xlsx';
employeesPercentagePerDay = 0.4

oddEvenDate = {
    "Monday": 1,
    "Tuesday": 0,
    "Wednesday": 1,
    "Thursday": 0,
    "Friday": 1,
    "Saturday": 0,
    "Sunday": 1
}

employeesData = getEmployeesData(employeesCsvPath)
fullDates = daysCurrentMonth()
schedule = getEmployeesToCome(employeesData, employeesPercentagePerDay, fullDates, oddEvenDate)
createExcelFromDictionary(schedule, resultExcelPath)
