"""Shared helpers for the Report notebooks. Every notebook starts with:  from setup import *"""
import os, sys, json, re, glob, shutil, subprocess, time, hashlib, datetime, zipfile, collections
from pathlib import Path
NB = Path(__file__).resolve().parent
ROOT = NB.parent                      # the Report folder
PROJECT = ROOT.parent                 # the original project folder (used to gather the starting data and to compare)
os.chdir(ROOT)
# Read and write text as UTF-8 on every computer (Windows otherwise uses an older encoding and fails on characters such as arrows)
import builtins
if not getattr(builtins, '_utf8_open', False):
    _open = builtins.open
    def _utf8_open(file, mode='r', buffering=-1, encoding=None, *a, **k):
        if 'b' not in mode and encoding is None: encoding = 'utf-8'
        return _open(file, mode, buffering, encoding, *a, **k)
    builtins.open = _utf8_open; builtins._utf8_open = True
ENV = dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8')
OFFLINE = bool(os.environ.get('REPORT_OFFLINE'))   # set only for an offline test run; live internet steps are then skipped
_NEW = [('12', 'report'), ('16', 'my_checks'), ('15', 'checks'), ('14', 'notebooks'), ('13', 'brief'), ('11', 'manuscript'), ('10', 'figures'),
        ('09', 'tables'), ('08', 'taxonomy'), ('07', 'extraction'), ('06', 'tracker'), ('05', 'snowballing'), ('04', 'papers'), ('03', 'screening'),
        ('02', 'records'), ('01', 'search'), ('00', 'protocol')]
# the numbered folders of the original project (found by their number) and the names they have in the Report folder
RENAME = [(d[0].name, b) for n, b in _NEW for d in [sorted(x for x in PROJECT.glob(n + '_*') if x.is_dir())] if d]
def new_name(p):
    for a, b in RENAME: p = p.replace(a, b)
    return p

def J(p):
    with open(ROOT / p, encoding='utf-8') as f: return json.load(f)
def md5(p): return hashlib.md5(open(p, 'rb').read()).hexdigest()

CHECKS = []
def check(name, ok, detail=''):
    CHECKS.append({'check': name, 'pass': bool(ok), 'detail': str(detail)})
    print(('PASS  ' if ok else 'FAIL  ') + name + (('   [' + str(detail) + ']') if detail != '' else ''))
def expect(name, got, want): check(f'{name}: {want}', got == want, got)

def run(script, *args, timeout=7200, tail=15, folder='scripts', root_arg=True):
    """Run one script from Report/scripts with the Report folder as its argument and show the end of its output."""
    t0 = time.time()
    cmd = [sys.executable, str(ROOT / folder / script)] + ([str(ROOT)] if root_arg else []) + [str(a) for a in args]
    p = subprocess.run(cmd, capture_output=True, text=True, cwd=ROOT, timeout=timeout, encoding='utf-8', errors='replace', env=ENV)
    out = (p.stdout or '').replace(str(ROOT), '<Report>').rstrip().split('\n')
    print(f'--- {script}: exit code {p.returncode}, {time.time() - t0:.1f} s ---')
    print('\n'.join(l[:300] for l in out[-tail:]))
    fails = [l for l in out if l.startswith('FAIL')]      # the whole output is searched, not only the lines shown above
    if fails and len(out) > tail: print('FAIL lines in the output:'); print('\n'.join('  ' + l[:300] for l in fails[:10]))
    if p.returncode: print('ERROR:', (p.stderr or '')[-2500:])
    check(f'{script} ran without error', p.returncode == 0 and not fails, f'{len(fails)} FAIL lines, first: {fails[0][:150]}' if fails else '')
    return p

def save_checks(name):
    os.makedirs(ROOT / 'checks' / 'notebook_checks', exist_ok=True)
    json.dump(CHECKS, open(ROOT / 'checks' / 'notebook_checks' / f'{name}.json', 'w'), indent=1)
    n = sum(c['pass'] for c in CHECKS)
    print(f'{n} of {len(CHECKS)} checks pass in {name}')

# ---------------- comparison with the original project ----------------
def norm_text(s):
    s = new_name(s)
    return re.sub(r'\d{4}-\d\d-\d\d \d\d:\d\d', '<time>', s)
def xlsx_cells(p):
    import openpyxl
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    return {ws.title: [tuple(norm_text(v) if isinstance(v, str) else v for v in r) for r in ws.iter_rows(values_only=True)] for ws in wb}
def docx_text(p):
    return norm_text(re.sub(r'<[^>]+>', '', zipfile.ZipFile(p).read('word/document.xml').decode('utf-8', 'ignore')))
def json_norm(p):
    return json.loads(norm_text(open(p, encoding='utf-8').read()))
# files that are expected to differ from the copy of 13 November because the full text rule and the later checks changed the included set
CHANGED_BY_RULE = ['search/counts/prisma_final_v2.json', 'search/counts/recall_test.json', 'papers/retrieval_log/text_index.json',
                   'tracker/master_tracker/Master_Tracker_LATEST.xlsx',
                   'extraction/adoption_v2.json', 'extraction/analyses_and_surveys_v1.json', 'extraction/benchmark_roles_v1.json',
                   'extraction/benchmarks_coded_final_v1.json', 'extraction/tier2/tier2_all_merged.json', 'extraction/tier2/tier2_summary_counts.json',
                   'extraction/unique_benchmarks_counts_v1.json', 'extraction/unique_benchmarks_v1.json',
                   'taxonomy/final/taxonomy_tree_v1.json', 'taxonomy/final/taxonomy_v1.md',
                   'tables/manuscript_tables/T4_classification_scheme.md', 'tables/manuscript_tables/T4_classification_scheme.xlsx',
                   'tables/manuscript_tables/T5_main_benchmark_table.md', 'tables/manuscript_tables/T5_main_benchmark_table.xlsx',
                   'tables/manuscript_tables/T7_contamination_by_category.md', 'tables/manuscript_tables/T7_contamination_by_category.xlsx',
                   'tables/manuscript_tables/T_adoption_top30_v2.md', 'tables/manuscript_tables/T_adoption_top30_v2.xlsx',
                   'figures/lineage/swe_bench_family_lineage_data.json', 'figures/sources/figure_data_v1.json']
NOT_COMPARED = []     # files for which the original project folder holds no copy, so no comparison was made
def same_as_original(new_rel, old_rel=None, show=True):
    """Compare a file made here with the same file in the original project (folder names mapped)."""
    a = ROOT / new_rel
    if old_rel is None:
        old_rel = new_rel
        for x, y in RENAME: old_rel = re.sub(r'(^|/)' + re.escape(y) + r'(/|$)', r'\g<1>' + x + r'\2', old_rel, count=1) if old_rel.split('/')[0] == y else old_rel
    b = PROJECT / old_rel
    if not b.exists(): print('  (original not available, comparison not performed)', new_rel); NOT_COMPARED.append(new_rel); return None
    if not a.exists(): check(f'made {new_rel}', False, 'missing'); return False
    if md5(a) == md5(b): how, ok = 'identical bytes', True
    elif new_rel.endswith('.json'): ok = json_norm(a) == json_norm(b); how = 'identical content' if ok else 'content differs'
    elif new_rel.endswith('.xlsx'): ok = xlsx_cells(a) == xlsx_cells(b); how = 'identical cells' if ok else 'cells differ'
    elif new_rel.endswith('.docx'): ok = docx_text(a) == docx_text(b); how = 'identical text' if ok else 'text differs'
    elif new_rel.endswith(('.md', '.txt', '.csv')): ok = norm_text(open(a, encoding='utf-8').read()) == norm_text(open(b, encoding='utf-8').read()); how = 'identical text' if ok else 'text differs'
    elif new_rel.endswith(('.png', '.pdf')): ok = None; how = 'image drawn again (bytes depend on the computer; the data behind it is compared separately)'
    else: ok = False; how = 'differs'
    if ok is None: print(f'  INFO  {new_rel}: {how}'); return None
    if not ok and new_rel in CHANGED_BY_RULE:
        print(f'  INFO  {new_rel}: differs from the copy of 13 November, as expected after the full text rule ({how})'); return None
    check(f'same as the original: {new_rel}', ok, how)
    return ok
