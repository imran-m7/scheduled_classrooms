import openpyxl

wb = openpyxl.load_workbook('course_assignments.xlsx')
ws = wb.active

total = 0
assigned = 0
unassigned = 0
infeasible = 0
other_status = {}

for row in ws.iter_rows(min_row=2, values_only=True):
    total += 1
    status = str(row[8]) if row[8] else ''
    if status.startswith('Assigned'):
        assigned += 1
    elif 'Unassigned' in status:
        unassigned += 1
    elif 'Infeasible' in status:
        infeasible += 1
    else:
        other_status[status] = other_status.get(status, 0) + 1

print(f"Total rows in Excel: {total}")
print(f"Assigned: {assigned}")
print(f"Unassigned: {unassigned}")
print(f"Infeasible: {infeasible}")
if other_status:
    print("Other statuses:")
    for k, v in other_status.items():
        print(f"  {k}: {v}")
