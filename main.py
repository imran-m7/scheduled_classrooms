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

    # Special case: merge all ENS207-3.* and ENS207-6.* into ENS207
    print("DEBUG: ENS207-3.* and ENS207-6.* enrollments before merge:")
    for k in enrollments_raw:
        if k.startswith('ENS207-3.') or k.startswith('ENS207-6.'):
            print(f"{k}: {enrollments_raw[k]}")
    ens207_total = 0
    ens207_times = set()
    ens207_rooms = set()
    to_delete = []
    # Only sum the first nonzero ENS207-3.* and ENS207-6.*
    ens207_3_found = False
    ens207_6_found = False
    for k in sorted(enrollments_raw.keys()):
        if k.startswith('ENS207-3.') and not ens207_3_found and enrollments_raw[k] > 0:
            ens207_total += enrollments_raw[k]
            ens207_3_found = True
            to_delete.append(k)
        elif k.startswith('ENS207-6.') and not ens207_6_found and enrollments_raw[k] > 0:
            ens207_total += enrollments_raw[k]
            ens207_6_found = True
            to_delete.append(k)
        elif k.startswith('ENS207-3.') or k.startswith('ENS207-6.'):
            to_delete.append(k)
    for k in to_delete:
        if k in enrollments_raw:
            del enrollments_raw[k]
    if ens207_total > 0:
        enrollments_raw['ENS207'] = ens207_total
    print(f"DEBUG: ENS207 merged total enrollment: {ens207_total}")

    # Special case: merge ENS209-3 and ENS209-6 into ENS209, and map ENS209-3/6.* in schedule to ENS209
    print("DEBUG: ENS209-3 and ENS209-6 enrollments before merge:")
    ens209_total = 0
    ens209_3_found = False
    ens209_6_found = False
    to_delete_209 = []
    for k in sorted(enrollments_raw.keys()):
        if k.startswith('ENS209-3') and not ens209_3_found and enrollments_raw[k] > 0:
            print(f"{k}: {enrollments_raw[k]}")
            ens209_total += enrollments_raw[k]
            ens209_3_found = True
            to_delete_209.append(k)
        elif k.startswith('ENS209-6') and not ens209_6_found and enrollments_raw[k] > 0:
            print(f"{k}: {enrollments_raw[k]}")
            ens209_total += enrollments_raw[k]
            ens209_6_found = True
            to_delete_209.append(k)
        elif k.startswith('ENS209-3') or k.startswith('ENS209-6'):
            to_delete_209.append(k)
    for k in to_delete_209:
        if k in enrollments_raw:
            del enrollments_raw[k]
    if ens209_total > 0:
        enrollments_raw['ENS209'] = ens209_total
    print(f"DEBUG: ENS209 merged total enrollment: {ens209_total}")

    # Special case: merge ARCH216.1 and ARCH216-6.1 into ARCH216
    print("DEBUG: ARCH216.1 and ARCH216-6.1 enrollments before merge:")
    arch216_total = 0
    arch216_1_found = False
    arch216_6_found = False
    to_delete_arch216 = []
    for k in sorted(enrollments_raw.keys()):
        if k == 'ARCH216.1' and not arch216_1_found and enrollments_raw[k] > 0:
            print(f"{k}: {enrollments_raw[k]}")
            arch216_total += enrollments_raw[k]
            arch216_1_found = True
            to_delete_arch216.append(k)
        elif k == 'ARCH216-6.1' and not arch216_6_found and enrollments_raw[k] > 0:
            print(f"{k}: {enrollments_raw[k]}")
            arch216_total += enrollments_raw[k]
            arch216_6_found = True
            to_delete_arch216.append(k)
        elif k == 'ARCH216.1' or k == 'ARCH216-6.1':
            to_delete_arch216.append(k)
    for k in to_delete_arch216:
        if k in enrollments_raw:
            del enrollments_raw[k]
    if arch216_total > 0:
        enrollments_raw['ARCH216'] = arch216_total
    print(f"DEBUG: ARCH216 merged total enrollment: {arch216_total}")

    schedule_main = load_course_schedule(SCHEDULE_DOCX)
    schedule_grad = load_course_schedule(GRADUATE_DOCX)
    schedule = schedule_main + schedule_grad

    # --- Ensure POLS304.1 is present in the schedule if in enrollments but missing from schedule ---
    if 'POLS304.1' in enrollments_raw:
        found_pols3041 = any(s['course_code'] == 'POLS304.1' for s in schedule)
        if not found_pols3041:
            print('Adding missing POLS304.1 to schedule with time Wed. 12:00-14:50 and room to be assigned by model.')
            schedule.append({'course_code': 'POLS304.1', 'time': 'Wed. 12:00-14:50', 'room': ''})

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

    # --- Ensure POLS304.1 is present in the schedule if in enrollments but missing from schedule ---
    if 'POLS304.1' in enrollments_raw:
        found_pols3041 = any(s['course_code'] == 'POLS304.1' for s in schedule)
        if not found_pols3041:
            print('Adding missing POLS304.1 to schedule with time Wed. 12:00-14:50 and room to be assigned by model.')
            schedule.append({'course_code': 'POLS304.1', 'time': 'Wed. 12:00-14:50', 'room': ''})

    # Update schedule: replace all ENS207-3.* and ENS207-6.* with ENS207, collect all times/rooms
    new_schedule = []
    for s in schedule:
        # Map ARCH216.1 and ARCH216-6.1 to ARCH216
        if s['course_code'] in ('ARCH216.1', 'ARCH216-6.1'):
            s['course_code'] = 'ARCH216'
        # Map ENS209-3/6.* to ENS209
        if s['course_code'].startswith('ENS209-3/6.'):
            s['course_code'] = 'ENS209'
        if s['course_code'].startswith('ENS207-3.') or s['course_code'].startswith('ENS207-6.'):
            ens207_times.add(s['time'])
            ens207_rooms.add(s['room'])
            continue  # skip these
        if s['course_code'].startswith('ENS209-3.') or s['course_code'].startswith('ENS209-6.'):
            continue  # skip these (should not appear, but for safety)
        new_schedule.append(s)
    # Add ENS207 for each unique time/room pair
    for t in ens207_times:
        for r in ens207_rooms:
            new_schedule.append({'course_code': 'ENS207', 'time': t, 'room': r})
    # Deduplicate ARCH216 entries (same time/room)
    seen_arch216 = set()
    deduped_schedule = []
    for s in new_schedule:
        if s['course_code'] == 'ARCH216':
            key = (s['course_code'], s['time'], s['room'])
            if key in seen_arch216:
                continue
            seen_arch216.add(key)
        deduped_schedule.append(s)
    schedule = deduped_schedule

    # Helper: get enrollment for a sectioned course code
    def get_enrollment(code):
        if code in enrollments_raw:
            return enrollments_raw[code]
        base = code.split('.')[0]
        if base in enrollments_raw:
            return enrollments_raw[base]
        return None

    # Ensure graduate courses CS600, EE603, ME580, ME605 are present in the schedule before MILP model setup
    grad_courses = ['CS600', 'EE603', 'ME580', 'ME605']
    grad_needed = set()
    for gc in grad_courses:
        code = f'{gc}.1'
        if code in enrollments_raw:
            grad_needed.add(code)
    grad_schedule = load_course_schedule(GRADUATE_DOCX)
    for gc in grad_needed:
        found = any(s['course_code'] == gc for s in schedule)
        if not found:
            # Try to match by base code in grad_schedule
            base_gc = gc.split('.')[0]
            for s in grad_schedule:
                sched_base = s['course_code'].split('.')[0]
                if sched_base == base_gc:
                    # Append with correct section code
                    schedule.append({'course_code': gc, 'time': s['time'], 'room': s['room']})
                    break

    # Debug: Print grad course presence in enrollments and schedule
    print('--- Graduate course debug ---')
    for gc in grad_needed:
        print(f'{gc} in enrollments_raw: {gc in enrollments_raw}, enrollment: {enrollments_raw.get(gc)}')
        found_sched = [s for s in schedule if s['course_code'] == gc]
        print(f'{gc} in schedule: {len(found_sched)} entries')
        for s in found_sched:
            print(f'  time: {s["time"]}, room: {s["room"]}')
    print('--- End grad course debug ---')

    # Build sets (only include courses with enrollment info)
    # If a sectioned code (e.g., POLS304.1) is in the schedule but only the base code (POLS304) is in enrollments, include the sectioned code in courses
    courses = set()
    for s in schedule:
        code = s['course_code']
        base = code.split('.')[0]
        # Add sectioned code if either code or base is in enrollments_raw
        if code in enrollments_raw:
            courses.add(code)
        elif base in enrollments_raw:
            courses.add(code)
    courses = list(courses)
    rooms = list(capacities.keys())
    times = list(set(s['time'] for s in schedule))
    course_time = {s['course_code']: s['time'] for s in schedule}

    # Debug: Check for POLS304.1 in courses and its enrollment
    print('--- POLS304.1 debug ---')
    print('POLS304.1 in courses:', 'POLS304.1' in courses)
    print('POLS304.1 enrollment:', get_enrollment('POLS304.1'))
    print('--- End POLS304.1 debug ---')

    # Diagnostic: Print all schedule entries with base code POLS304
    print('--- POLS304 schedule entries debug ---')
    for s in schedule:
        if s['course_code'].split('.')[0] == 'POLS304':
            print(f"Schedule entry: {s['course_code']} at {s['time']} in {s['room']}")
    print('--- End POLS304 schedule entries debug ---')

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
    print('\n--- Infeasible Courses (no room large enough, before Excel output) ---')
    for c in courses:
        if all(get_enrollment(c) > capacities[r] for r in rooms):
            print(f'Course {c} (enrollment: {get_enrollment(c)})')

    # Output results to Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Assignments'
    ws.append(['Course Code', 'Assigned Room', 'Time', 'Enrollment', 'Room Capacity', 'Assignment Status'])

    assigned_courses = 0
    row_idx = 2  # Excel rows are 1-based, header is row 1
    excel_rows_written = 0
    print('\n--- Excel output course codes ---')
    pols3041_found = False
    for c in courses:
        if c == 'POLS304.1':
            pols3041_found = True
            # Removed verbose print
        # Removed verbose print for all other courses
        try:
            assigned_room = None
            for r in rooms:
                if pulp.value(x[c, r]) == 1:
                    assigned_room = r
                    break
            enrollment = get_enrollment(c)
            time = course_time.get(c, 'MISSING_TIME')
            excel_code = c
            # Special renaming for ENS209, ENS207, and ARCH216 in specific rows
            if row_idx == 131 and c == 'ENS209':
                excel_code = 'ENS209-3/6.1'
            if row_idx == 334 and c == 'ENS207':
                excel_code = 'ENS207-3/6.1'
            if c == 'ARCH216':
                excel_code = 'ARCH216-3/6.1'
            if assigned_room:
                ws.append([excel_code, assigned_room, time, enrollment, capacities[assigned_room], 'Assigned'])
                assigned_courses += 1
            else:
                infeasible = all(enrollment > capacities[r] for r in rooms)
                status = 'Infeasible' if infeasible else 'Unassigned'
                ws.append([excel_code, '', time, enrollment, '', status])
            row_idx += 1
            excel_rows_written += 1
        except Exception as e:
            print(f'Exception for course {c}: {e}')
    print(f'Total courses: {len(courses)}, Excel rows written: {excel_rows_written}')
    print('--- End Excel output course codes ---')
    if not pols3041_found:
        print('WARNING: POLS304.1 was not found in the Excel output loop!')

    wb.save('course_assignments.xlsx')
    print(f"\nResults saved to course_assignments.xlsx. Total assigned courses: {assigned_courses} out of {len(courses)}")

if __name__ == '__main__':
    main()