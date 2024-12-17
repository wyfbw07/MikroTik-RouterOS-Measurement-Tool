import re
import pandas as pd

# Replace 'owping_output.txt' with the path to your output file
with open('30-1.txt', 'r') as f:
    content = f.read()

lines = content.splitlines()

# Parse the input text
sections = []
current_section = None
in_measurements = False
in_summary_stats = False

for line in lines:
    line = line.strip()
    if line.startswith('--- owping statistics from'):
        # Handle the start of a new section
        match = re.match(r'--- owping statistics from \[(.*?)\]:\d+ to \[(.*?)\]:\d+ ---', line)
        if match:
            from_host = match.group(1)
            to_host = match.group(2)
            if current_section is not None and in_summary_stats:
                sections.append(current_section)
                current_section = None
            if current_section is None:
                current_section = {
                    'from': from_host,
                    'to': to_host,
                    'measurements': [],
                    'summary': {}
                }
                in_measurements = True
                in_summary_stats = False
            elif in_measurements:
                # Transition to summary stats
                in_measurements = False
                in_summary_stats = True
    elif line.startswith('SID:'):
        sid = line.split(':', 1)[1].strip()
        current_section['SID'] = sid
    elif in_measurements and line.startswith('seq_no='):
        # Parse measurement line
        match = re.match(r'seq_no=(\d+)\s+delay=([-\d.]+) ms\s+\(([^)]*)\)', line)
        if match:
            seq_no = int(match.group(1))
            delay = float(match.group(2))
            extra_info = match.group(3)
            # Parse extra_info
            sync = ''
            err = ''
            if 'sync' in extra_info:
                sync = 'sync'
            err_match = re.search(r'err=([-\d.]+) ms', extra_info)
            if err_match:
                err = float(err_match.group(1))
            current_section['measurements'].append({
                'seq_no': seq_no,
                'delay': delay,
                'sync': sync,
                'err': err
            })
    elif in_summary_stats:
        # Parse summary stats
        if line.startswith('first:'):
            current_section['summary']['first'] = line.split(':', 1)[1].strip()
        elif line.startswith('last:'):
            current_section['summary']['last'] = line.split(':', 1)[1].strip()
        elif 'sent,' in line and 'lost' in line:
            match = re.match(r'(\d+) sent, (\d+) lost.* (\d+) duplicates', line)
            if match:
                sent = int(match.group(1))
                lost = int(match.group(2))
                duplicates = int(match.group(3))
                current_section['summary']['sent'] = sent
                current_section['summary']['lost'] = lost
                current_section['summary']['duplicates'] = duplicates
        elif line.startswith('one-way delay min/median/max'):
            match = re.match(r'one-way delay min/median/max = ([^ ]+) ms, \(err=([^\)]+)\)', line)
            if match:
                delays = match.group(1).split('/')
                err = float(match.group(2).strip().replace(' ms', ''))
                current_section['summary']['delay_min'] = float(delays[0])
                current_section['summary']['delay_median'] = float(delays[1])
                current_section['summary']['delay_max'] = float(delays[2])
                current_section['summary']['err'] = err
        elif line.startswith('one-way jitter ='):
            match = re.match(r'one-way jitter = ([^ ]+) ms.*', line)
            if match:
                jitter = float(match.group(1))
                current_section['summary']['jitter'] = jitter
        elif line.startswith('TTL'):
            current_section['summary']['TTL'] = line
        elif line.startswith('no reordering') or line.startswith('reordering'):
            current_section['summary']['reordering'] = line
# Append the last section
if current_section is not None:
    sections.append(current_section)

# Prepare data for Excel
measurement_rows = []
summary_rows = []
set_index = 0
all_measurements = []

for section in sections:
    from_host = section['from']
    to_host = section['to']
    SID = section.get('SID', '')

    # Create a DataFrame for measurements in this section
    df_measurements = pd.DataFrame([
        {
            'from_host': from_host,
            'to_host': to_host,
            'SID': SID,
            'seq_no': m['seq_no'],
            'delay (ms)': m['delay'],
            'sync': m['sync'],
            'err (ms)': m['err']
        }
        for m in section['measurements']
    ])
    
    # Track resets in seq_no to create new sheets
    set_index += 1
    all_measurements.append((f'Set_{set_index}', df_measurements))

# Write to Excel with separate sheets
output_path = 'results_with_multiple_sheets.xlsx'
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, df in all_measurements:
        df.to_excel(writer, sheet_name=sheet_name, index=False)