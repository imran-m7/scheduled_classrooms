import csv
from collections import defaultdict
from docx import Document
import pulp
import os
import openpyxl

# File paths
COURSES_CSV = 'AcilanDersler.csv'
ROOMS_CSV = 'Class Quotas  E-Campus.csv'
SCHEDULE_DOCX = '2025spring_schedule_march_28_1515.docx'
GRADUATE_DOCX = 'graduate.docx'

# 1. Parse course enrollments
def load_course_enrollments(csv_path):
    enrollments = {}
    code_counts = defaultdict(int)
    encodings = ['utf-8-sig', 'cp1254', 'latin1']
    for enc in encodings:
        try:
            with open(csv_path, encoding=enc) as f:
                reader = csv.reader(f)
                header = next(reader)
                # Find indices
                code_idx = None
                exist_idx = None
                for i, col in enumerate(header):
                    if 'Course Code' in col:
                        code_idx = i
                    if 'Existing' in col:
                        exist_idx = i
                if code_idx is None or exist_idx is None:
                    continue
                for row in reader:
                    if len(row) > max(code_idx, exist_idx):
                        base_code = row[code_idx]
                        try:
                            n = int(row[exist_idx])
                        except (ValueError, IndexError):
                            continue
                        if base_code:
                            code_counts[base_code] += 1
                            sectioned_code = f"{base_code}.{code_counts[base_code]}"
                            enrollments[sectioned_code] = n
            break
        except UnicodeDecodeError:
            continue
    return enrollments

# 2. Parse room capacities
def load_room_capacities(csv_path):
    capacities = {}
    encodings = ['utf-8-sig', 'cp1254', 'latin1']
    for enc in encodings:
        try:
            with open(csv_path, encoding=enc) as f:
                reader = csv.DictReader(f)
                for row in reader:
                    name = row.get('Name')
                    cap = row.get('Teaching Capacity')
                    if name and cap:
                        try:
                            capacities[name] = int(cap)
                        except ValueError:
                            continue
            break
        except UnicodeDecodeError:
            continue
    return capacities

# 3. Parse course schedule from DOCX
def load_course_schedule(docx_path):
    schedule = []
    doc = Document(docx_path)
    for table in doc.tables:
        headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
        # Try to find column indices for course code, time, and room
        code_idx = next((i for i, h in enumerate(headers) if 'code' in h), 0)
        time_idx = next((i for i, h in enumerate(headers) if 'time' in h or 'hour' in h), 2)
        room_idx = next((i for i, h in enumerate(headers) if 'room' in h or 'venue' in h), 3)
        for row in table.rows[1:]:
            cells = [cell.text.strip() for cell in row.cells]
            if len(cells) > max(code_idx, time_idx, room_idx):
                course_code = cells[code_idx]  # Keep section suffix (e.g., CS101.1)
                time = cells[time_idx]
                room = cells[room_idx]
                if course_code and time:
                    schedule.append({'course_code': course_code, 'time': time, 'room': room})
    return schedule

def main():
    # Load data
    enrollments_raw = load_course_enrollments(COURSES_CSV)
    capacities = load_room_capacities(ROOMS_CSV)
    schedule_main = load_course_schedule(SCHEDULE_DOCX)
    schedule_grad = load_course_schedule(GRADUATE_DOCX)
    schedule = schedule_main + schedule_grad

    # Remove known duplicate: ELT571.1 (keep only one entry with the same code and time)
    seen = set()
    deduped_schedule = []
    for s in schedule:
        key = (s['course_code'], s['time'])
        if key == ("ELT571.1", s['time']):
            if key in seen:
                continue
        if key not in seen:
            deduped_schedule.append(s)
            seen.add(key)
    schedule = deduped_schedule

    # Helper: get enrollment for a sectioned course code
    def get_enrollment(code):
        if code in enrollments_raw:
            return enrollments_raw[code]
        base = code.split('.')[0]
        return enrollments_raw.get(base, None)

    # Build sets (only include courses with enrollment info)
    courses = [s['course_code'] for s in schedule if get_enrollment(s['course_code']) is not None]
    rooms = list(capacities.keys())
    times = list(set(s['time'] for s in schedule))
    course_time = {s['course_code']: s['time'] for s in schedule}

    # Decision variables: x[c, r] = 1 if course c assigned to room r
    x = pulp.LpVariable.dicts('assign', ((c, r) for c in courses for r in rooms), cat='Binary')

    # Model
    prob = pulp.LpProblem('ClassroomAssignment', pulp.LpMinimize)

    # Objective: Minimize total unused seat-hours
    prob += pulp.lpSum([
        x[c, r] * (capacities[r] - get_enrollment(c))
        for c in courses for r in rooms if capacities[r] >= get_enrollment(c)
    ])

    # Constraints
    # 1. Each course assigned to exactly one room (with enough capacity)
    for c in courses:
        prob += pulp.lpSum([x[c, r] for r in rooms if capacities[r] >= get_enrollment(c)]) == 1

    # 2. No overlapping courses in the same room at the same time
    for r in rooms:
        for t in times:
            prob += pulp.lpSum([
                x[c, r] for c in courses if course_time[c] == t and capacities[r] >= get_enrollment(c)
            ]) <= 1

    # Solve
    prob.solve()

    # Output results
    print('Status:', pulp.LpStatus[prob.status])
    assigned_courses = 0
    total_unused_seat_hours = 0
    for c in courses:
        assigned = False
        for r in rooms:
            if pulp.value(x[c, r]) == 1:
                total_unused_seat_hours += capacities[r] - get_enrollment(c)
                assigned = True
        if assigned:
            assigned_courses += 1

    print(f"\nTotal assigned courses: {assigned_courses} out of {len(courses)}")
    print(f"Total unused seat-hours: {total_unused_seat_hours}")

    # List all unassigned courses
    print('\n--- Unassigned Courses (not assigned to any room) ---')
    for c in courses:
        assigned = any(pulp.value(x[c, r]) == 1 for r in rooms)
        if not assigned:
            print(f'Course {c} (enrollment: {get_enrollment(c)})')

    # Diagnostic: print infeasible courses (no room large enough)
    print('\n--- Infeasible Courses (no room large enough) ---')
    for c in courses:
        if all(get_enrollment(c) > capacities[r] for r in rooms):
            print(f'Course {c} (enrollment: {get_enrollment(c)})')

    # Output results to Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Assignments'
    ws.append(['Course Code', 'Assigned Room', 'Time', 'Enrollment', 'Room Capacity', 'Assignment Status'])

    assigned_courses = 0
    for c in courses:
        assigned_room = None
        for r in rooms:
            if pulp.value(x[c, r]) == 1:
                assigned_room = r
                break
        enrollment = get_enrollment(c)
        time = course_time[c]
        if assigned_room:
            ws.append([c, assigned_room, time, enrollment, capacities[assigned_room], 'Assigned'])
            assigned_courses += 1
        else:
            # Check if infeasible (no room large enough)
            infeasible = all(enrollment > capacities[r] for r in rooms)
            status = 'Infeasible' if infeasible else 'Unassigned'
            ws.append([c, '', time, enrollment, '', status])

    wb.save('course_assignments.xlsx')
    print(f"\nResults saved to course_assignments.xlsx. Total assigned courses: {assigned_courses} out of {len(courses)}")

if __name__ == '__main__':
    main()