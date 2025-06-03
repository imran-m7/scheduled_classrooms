import openpyxl

wb = openpyxl.load_workbook('course_assignments.xlsx')
ws = wb.active

for code in ['CS511.1', 'MBA535.1']:
    found = False
    for row in ws.iter_rows(min_row=2, values_only=True):
        course, room1, time1, room2, time2, *_ = row
        if course == code:
            print(f'{code}:')
            if room1:
                print(f'  Time 1: {time1}, Room: {room1}')
            if room2:
                print(f'  Time 2: {time2}, Room: {room2}')
            found = True
    if not found:
        print(f'{code}: Not found in Excel output')
