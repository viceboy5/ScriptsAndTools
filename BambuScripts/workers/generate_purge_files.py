import csv, itertools, zipfile, json, io, os, glob

prod_folder = r'C:\Users\Owner\SynologyDrive\WIGGLITEERZ\THEKITCHEN\Experimental\Hank\PurgeTests\Production'
rest_folder = r'C:\Users\Owner\SynologyDrive\WIGGLITEERZ\THEKITCHEN\Experimental\Hank\PurgeTests\The Rest'
template    = r'C:\Users\Owner\SynologyDrive\WIGGLITEERZ\THEKITCHEN\Experimental\Hank\PurgeTests\Purge_Test_Base.3mf'

# Clear both folders
for folder in [prod_folder, rest_folder]:
    os.makedirs(folder, exist_ok=True)
    for f in glob.glob(os.path.join(folder, '*.3mf')):
        os.remove(f)
print('Folders cleared.')

# Load filament colors
colors = {}
with open(r'C:\GitRepos\ScriptsAndTools\BambuScripts\libraries\FilamentLibrary.csv') as f:
    for row in csv.reader(f):
        if len(row) >= 4 and row[0] and row[0] != 'N/A':
            try:
                r, g, b = int(row[1]), int(row[2]), int(row[3])
                colors[row[0]] = '#{:02X}{:02X}{:02X}'.format(r, g, b)
            except:
                pass

filaments = [f for f in colors.keys() if 'Silk' not in f]
print(f'Non-silk filaments: {len(filaments)}')
print(f'Pairs: {len(filaments)*(len(filaments)-1)}  Theoretical min files: {len(filaments)*(len(filaments)-1)//12 + 1}')

# Production priority groups (from xlsx, normalized)
def norm(n):
    return 'Sunlu Mint Green' if n == 'Voxel Sunlu Mint Green' else n

xlsx_groups = [
    ['Eryone Silk (RPG)', 'Sunlu Silk (Blue)', 'Sunlu Silk (Bronze)', 'Sunlu Silk (White)'],
    ['Eryone Silk (RYB)', 'Esun Black', 'Esun Bone White', 'Esun Cold White'],
    ['Eryone Silk (RYB)', 'Esun Peak Green', 'Esun Beige', 'Esun Pink'],
    ['Eryone Silk (RYB)', 'Esun Warm White', 'Esun Black', 'Esun Blue'],
    ['Eryone Silk (RYB)', 'Sunlu Silk (Blue)', 'Esun Black', 'Sunlu Silk (Light Gold)'],
    ['Eryone Silk (RYB)', 'Sunlu Silk (Red)', 'Esun Black', 'Sunlu Silk (White)'],
    ['Esun Beige', 'Sunlu Silk (Blue)', 'Esun Bone White', 'Esun Brown'],
    ['Esun Beige', 'Sunlu Silk (Bronze)', 'Esun Black', 'Esun Yellow'],
    ['Esun Black', 'Esun Light Brown', 'Esun Blue', 'Esun Red'],
    ['Esun Black', 'Esun Magenta', 'Esun Blue', 'Esun Very Perri'],
    ['Esun Black', 'Esun Orange', 'Esun Cold White', 'Esun Dark Blue'],
    ['Esun Black', 'Esun Purple', 'Esun Cold White', 'Esun Magenta'],
    ['Esun Black', 'Esun Silver', 'Esun Bone White', 'Esun Light Blue'],
    ['Esun Black', 'Esun Space Blue', 'Esun Cold White', 'Esun Red'],
    ['Esun Black', 'Sunlu Mint Green', 'Esun Blue', 'Esun Yellow'],
    ['Esun Black', 'Sunlu Silk (Orange)', 'Esun Cold White', 'Sunlu Silk (Pink)'],
    ['Esun Black', 'Sunlu Silk (Purple)', 'Esun Bone White', 'Sunlu Silk (Red Copper)'],
    ['Esun Black', 'Sunlu Silk (Silver)', 'Esun Blue', 'Sunlu Silk (Light Gold)'],
    ['Esun Black', 'Sunlu Mint Green', 'Esun Cold White', 'Voxel Ziro Brown'],
    ['Esun Blue', 'Esun Bone White', 'Esun Brown', 'Esun Light Blue'],
    ['Esun Blue', 'Esun Gray', 'Esun Cold White', 'Esun Light Brown'],
    ['Esun Blue', 'Esun Peak Green', 'Esun Brown', 'Esun Pine Green'],
    ['Esun Bone White', 'Esun Light Brown', 'Esun Pink', 'Esun Red'],
    ['Esun Bone White', 'Esun Warm White', 'Esun Cold White', 'Esun Pink'],
    ['Esun Bone White', 'Esun Yellow', 'Esun Brown', 'Esun Cold White'],
    ['Esun Bone White', 'Sunlu Silk (Black)', 'Esun Cold White', 'Sunlu Silk (Light Gold)'],
    ['Esun Brown', 'Esun Gold', 'Esun Dark Blue', 'Esun Red'],
    ['Esun Brown', 'Esun Olive Green', 'Esun Cold White', 'Esun Very Perri'],
    ['Esun Brown', 'Esun Pink', 'Esun Light Blue', 'Esun Yellow'],
    ['Esun Brown', 'Sunlu Silk (Bronze)', 'Esun Cold White', 'Esun Pine Green'],
    ['Esun Brown', 'Sunlu Silk (Red Copper)', 'Esun Cold White', 'Sunlu Silk (Red)'],
    ['Esun Brown', 'Voxel Ziro Brown', 'Esun Magenta', 'Esun Pink'],
    ['Esun Cold White', 'Esun Jade Green', 'Esun Matcha Green', 'Esun Pine Green'],
    ['Esun Cold White', 'Esun Silver', 'Esun Pink', 'Sunlu Mint Green'],
    ['Esun Cold White', 'Sunlu Silk (Blue)', 'Esun Light Blue', 'Sunlu Silk (Red)'],
    ['Esun Cold White', 'Sunlu Silk (Purple)', 'Esun Light Blue', 'Sunlu Silk (White)'],
    ['Esun Cold White', 'Sunlu Silk (Silver)', 'Esun Dark Blue', 'Esun Light Brown'],
    ['Esun Dark Blue', 'Esun Yellow', 'Esun Orange', 'Esun Peak Green'],
    ['Esun Gold', 'Esun Light Blue', 'Esun Magenta', 'Esun Peak Green'],
    ['Esun Gold', 'Sunlu Silk (Red Copper)', 'Esun Peak Green', 'Sunlu Silk (Light Gold)'],
    ['Esun Gray', 'Esun Silver', 'Esun Peak Green', 'Esun Red'],
    ['Esun Green', 'Esun Peak Green', 'Esun Light Brown', 'Esun Matcha Green'],
    ['Esun Jade Green', 'Esun Yellow', 'Esun Peak Green', 'Esun Purple'],
    ['Esun Light Blue', 'Esun Warm White', 'Eryone Silk (RPG)', 'Sunlu Mint Green'],
    ['Esun Orange', 'Esun Pine Green', 'Esun Pink', 'Esun Red'],
    ['Esun Peak Green', 'Esun Very Perri', 'Esun Pink', 'Esun Yellow'],
    ['Esun Pink', 'Sunlu Silk (Pink)', 'Sunlu Silk (Light Gold)', 'Sunlu Silk (White)'],
    ['Esun Red', 'Esun Yellow', 'Eryone Silk (RPG)', 'Sunlu Silk (Blue)'],
]

prod_sets = {frozenset([norm(n) for n in g]) for g in xlsx_groups}

# Step 1: greedy set cover over ALL C(47,4) groups — no constraints
print('Precomputing C(47,4) groups...')
all_groups = []
for combo in itertools.combinations(filaments, 4):
    pairs = frozenset(itertools.permutations(combo, 2))
    all_groups.append((combo, pairs))
print(f'{len(all_groups)} groups. Running set cover...')

remaining = set(itertools.permutations(filaments, 2))
all_chosen = []  # list of (combo, pairs_frozenset)
while remaining:
    best_combo, best_pairs, best_score = None, None, -1
    for combo, pairs in all_groups:
        score = len(pairs & remaining)
        if score > best_score:
            best_score, best_combo, best_pairs = score, combo, pairs
    if best_score == 0:
        break
    all_chosen.append((best_combo, best_pairs))
    remaining -= best_pairs

print(f'Total files (set cover): {len(all_chosen)}')

# Step 2: find minimum subset of those files covering all PurgeVolumes.csv pairs
purge_pairs = set()
with open(r'C:\GitRepos\ScriptsAndTools\BambuScripts\libraries\PurgeVolumes.csv', encoding='utf-8-sig') as f:
    for row in csv.DictReader(f):
        frm, to = row['From'].strip(), row['To'].strip()
        if 'Silk' not in frm and 'Silk' not in to:
            purge_pairs.add((frm, to))

print(f'Production priority pairs to cover: {len(purge_pairs)}')

# Greedy set cover on just the purge pairs, using only files from all_chosen
prod_needed = set()
uncovered = set(purge_pairs)
while uncovered:
    best_idx, best_score = -1, -1
    for i, (combo, pairs) in enumerate(all_chosen):
        score = len(pairs & uncovered)
        if score > best_score:
            best_score, best_idx = score, i
    if best_score == 0:
        break
    prod_needed.add(best_idx)
    uncovered -= all_chosen[best_idx][1]

print(f'Files needed for Production: {len(prod_needed)}')
print(f'Files in The Rest: {len(all_chosen) - len(prod_needed)}')
print(f'Total: {len(all_chosen)}')

# Write files
def short(n):
    return n.replace('Esun ', '')

def make_3mf(group, dest_folder):
    hexes = [colors[n] for n in group]
    filename = '-'.join(short(n) for n in group) + '.3mf'
    out_path = os.path.join(dest_folder, filename)
    with zipfile.ZipFile(template, 'r') as zin:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'Metadata/project_settings.config':
                    cfg = json.loads(data.decode('utf-8'))
                    cfg['filament_colour'] = hexes
                    cfg['flush_multiplier'] = ['0']
                    cfg['flush_volumes_matrix'] = ['0'] * 16
                    data = json.dumps(cfg, indent=4).encode('utf-8')
                zout.writestr(item, data)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())

for i, (combo, _) in enumerate(all_chosen):
    folder = prod_folder if i in prod_needed else rest_folder
    make_3mf(combo, folder)

prod_count = len(prod_needed)
rest_count = len(all_chosen) - prod_count
print(f'Production: {prod_count}  The Rest: {rest_count}  Total: {len(all_chosen)}')
print('Done.')
