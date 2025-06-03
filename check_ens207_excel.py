import openpyxl

wb = openpyxl.load_workbook('course_assignments.xlsx')
ws = wb.active
for row in ws.iter_rows(min_row=2, values_only=True):
    print(f'{row[0]}: Enrollment={row[3]}')
