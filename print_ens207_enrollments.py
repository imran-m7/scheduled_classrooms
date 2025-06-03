import csv
encodings = ['utf-8-sig', 'cp1254', 'latin1']
for enc in encodings:
    try:
        with open('AcilanDersler.csv', encoding=enc) as f:
            reader = csv.reader(f)
            rows = list(reader)[2:]  # skip first two rows
            for row in rows:
                if len(row) > 1 and (row[1].startswith('ENS207-3') or row[1].startswith('ENS207-6')):
                    print(f'{row[1]}: {row[-1]}')
        break
    except UnicodeDecodeError:
        continue
