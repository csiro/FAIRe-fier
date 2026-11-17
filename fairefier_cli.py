##################################################
# FAIRe-fier COMMAND-LINE INTERFACE
#
# Runs the FAIRe-fier validation without the Shiny app, for scripts, pipelines
# and workflow systems such as Galaxy. Reading, validation and writing of the
# workbooks are done by lib/validate_core.py, the same code the app uses
# (validate_module.py), so both give the same results. This script adds only
# its arguments, the choice of checklist, and two plain-text outputs.
#
# Usage:
#   python fairefier_cli.py WORKBOOK.xlsx [-o OUTDIR] [--checklist_version v1.0]
#                           [--checklist_dir DIR] [--strict]
#
# Outputs (the same files as the app's download buttons):
#   fairefier_metadata.xlsx    when a worksheet validates
#   fairefier_warn_error.xlsx  when there are warnings or errors
# plus, for scripts and pipelines:
#   fairefier_issues.tsv       a tab-separated table of every warning and error
#   fairefier_summary.json     pass or fail for each worksheet, counts and the app's message
#
# Contributed by Oliver Berry (CSIRO) for the Galaxy Australia eDNA Lab, under
# FAIRe-fier's BSD 3-Clause licence.
##################################################

import argparse
import json
import re
import sys
import warnings
from pathlib import Path

import pandas as pd

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))
import lib.validate_core as validate_core  # noqa: E402
from lib.validate_core import ReadError, DuplicateSampleError  # noqa: E402

term_name_col = validate_core.term_name_col
samp_name_col = validate_core.samp_name_col


# ---------------------------------------------------------------- plain-text outputs

def issues_table(res):
    """All warnings and errors in one long table, for scripts and pipelines."""
    parts = []
    for sheet, key_col in (('projectMetadata', term_name_col), ('sampleMetadata', samp_name_col)):
        for level in ('error', 'warning'):
            df = res[f"{'warn' if level == 'warning' else 'error'}_{sheet.lower()}"]
            if df.empty:
                continue
            t = pd.DataFrame({
                'worksheet': sheet,
                'level': level,
                'record': df[key_col].astype(str).str.strip(),
                'term': df['loc'].astype(str).str.strip(),
                'input': df['input'].astype('string').str.strip() if 'input' in df else None,
                'type': df['type'] if 'type' in df else '',
                'message': df['msg'].astype(str).str.strip(),
                'suggested': df['output'].astype('string').str.strip() if 'output' in df else '',
            })
            parts.append(t)
    cols = ['worksheet', 'level', 'record', 'term', 'input', 'type', 'message', 'suggested']
    return pd.concat(parts, ignore_index=True)[cols] if parts else pd.DataFrame(columns=cols)


# ---------------------------------------------------------------- checklist choice

def available_checklists(checklist_dir):
    """{version: path} for FAIRe_checklist_<version>.xlsx files, as in config_data."""
    found = {}
    for p in sorted(Path(checklist_dir).glob('FAIRe_checklist*.xlsx')):
        found[p.stem.rsplit('_', 1)[-1]] = p
    return found


def version_key(v):
    return [int(x) if x.isdigit() else x for x in re.findall(r'\d+|[A-Za-z]+', v)]


def choose_checklist(requested, checklist_dir, df_xlsx_projectmetadata):
    """Return (version, path, note). 'auto' reads checkls_ver from projectMetadata."""
    found = available_checklists(checklist_dir)
    if not found:
        raise ReadError(f"No FAIRe_checklist_*.xlsx files in {checklist_dir}.")
    latest = max(found, key=version_key)
    if requested != 'auto':
        if requested not in found:
            raise ReadError(f"Checklist version {requested} is not available. Available: {', '.join(sorted(found, key=version_key))}.")
        return requested, found[requested], f"Checklist {requested} chosen by the user."
    declared = None
    pm = df_xlsx_projectmetadata
    if term_name_col in pm.columns and 'project_level' in pm.columns:
        vals = pm.loc[pm[term_name_col] == 'checkls_ver', 'project_level']
        if len(vals) and not pd.isna(vals.iloc[0]) and str(vals.iloc[0]).strip():
            declared = str(vals.iloc[0]).strip()
    if declared:
        # Either a file name (FAIRe_checklist_v1.0.2.xlsx, as FAIRe-ator writes it) or a number (1.0)
        m = re.search(r'v(\d+(?:\.\d+)*)', declared) or re.fullmatch(r'(\d+(?:\.\d+)*)', declared)
        if m and f"v{m.group(1)}" in found:
            return f"v{m.group(1)}", found[f"v{m.group(1)}"], f"Checklist v{m.group(1)} taken from checkls_ver ({declared})."
        if m:   # a shorter version such as 1 or 1.0: the latest checklist it covers
            same = [k for k in found if k.startswith(f"v{m.group(1)}.")]
            if same:
                best = max(same, key=version_key)
                return best, found[best], f"checkls_ver is '{declared}'; used the latest matching checklist, {best}."
        return latest, found[latest], f"checkls_ver is '{declared}', which is not available; used the latest checklist, {latest}."
    return latest, found[latest], f"No checkls_ver in projectMetadata; used the latest checklist, {latest}."


# ---------------------------------------------------------------- main

def main(argv=None):
    parser = argparse.ArgumentParser(
        prog='fairefier',
        description='Validate a FAIRe eDNA metadata workbook (projectMetadata and sampleMetadata) against the FAIRe checklist.')
    parser.add_argument('workbook', help='Excel workbook with projectMetadata and sampleMetadata worksheets (or the _revise worksheets from a previous run)')
    parser.add_argument('-o', '--outdir', default='.', help='folder for the output files (default: current folder)')
    parser.add_argument('--checklist_version', default='auto',
                        help="checklist version, for example v1.0.2, or 'auto' to use checkls_ver from projectMetadata (default)")
    parser.add_argument('--checklist_dir', default=str(Path(__file__).resolve().parent),
                        help='folder holding FAIRe_checklist_<version>.xlsx files (default: the folder of this script)')
    parser.add_argument('--strict', action='store_true',
                        help='exit with status 1 if either worksheet fails validation')
    args = parser.parse_args(argv)

    # FAIRe-fier reports its warnings (e.g. "Closest option found") through Python's warnings
    # module. Some environments silence warnings (Galaxy job scripts set PYTHONWARNINGS=ignore),
    # which would drop them from the report. Restore Python's standard filters.
    warnings.resetwarnings()
    warnings.filterwarnings('default', category=DeprecationWarning, module='__main__', append=True)
    warnings.filterwarnings('ignore', category=DeprecationWarning, append=True)
    warnings.filterwarnings('ignore', category=PendingDeprecationWarning, append=True)
    warnings.filterwarnings('ignore', category=ImportWarning, append=True)
    warnings.filterwarnings('ignore', category=ResourceWarning, append=True)

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    try:
        df_pm, df_sm = validate_core.read_worksheets(args.workbook)
        version, checklist_path, note = choose_checklist(args.checklist_version, args.checklist_dir, df_pm)
    except DuplicateSampleError as e:
        sys.stderr.write(f"ERROR: {e}. Please ensure sample names are unique.\n")
        return 2
    except ReadError as e:
        sys.stderr.write(f"ERROR: {e}\n")
        return 2
    print(f"INFO: {note}")

    res = validate_core.run_validator(checklist_path, df_pm, df_sm)
    print(res['message'])

    written = []
    if not res['out_projectmetadata'].empty or not res['out_samplemetadata'].empty:
        validate_core.write_metadata(res['out_projectmetadata'], res['out_samplemetadata'], outdir / 'fairefier_metadata.xlsx')
        written.append('fairefier_metadata.xlsx')
    if any(not res[k].empty for k in ('warn_projectmetadata', 'error_projectmetadata', 'warn_samplemetadata', 'error_samplemetadata')):
        validate_core.write_warn_error(df_pm, res['warn_projectmetadata'], res['error_projectmetadata'],
                                       df_sm, res['warn_samplemetadata'], res['error_samplemetadata'],
                                       outdir / 'fairefier_warn_error.xlsx')
        written.append('fairefier_warn_error.xlsx')

    issues = issues_table(res)
    issues.to_csv(outdir / 'fairefier_issues.tsv', sep='\t', index=False)
    pm_ok = 'For projectMetadata: Success!' in res['message']
    sm_ok = 'For sampleMetadata: Success!' in res['message']
    summary = {
        'workbook': str(args.workbook),
        'checklist_version': version,
        'checklist_note': note,
        'valid': pm_ok and sm_ok,
        'projectMetadata': {'valid': pm_ok,
                            'errors': int(((issues.worksheet == 'projectMetadata') & (issues.level == 'error')).sum()),
                            'warnings': int(((issues.worksheet == 'projectMetadata') & (issues.level == 'warning')).sum())},
        'sampleMetadata': {'valid': sm_ok,
                           'errors': int(((issues.worksheet == 'sampleMetadata') & (issues.level == 'error')).sum()),
                           'warnings': int(((issues.worksheet == 'sampleMetadata') & (issues.level == 'warning')).sum())},
        'message': res['message'],
        'files': written + ['fairefier_issues.tsv'],
    }
    with open(outdir / 'fairefier_summary.json', 'w') as f:
        json.dump(summary, f, indent=2)
    print(f"INFO: wrote {', '.join(summary['files'])} and fairefier_summary.json")
    return 1 if (args.strict and not summary['valid']) else 0


if __name__ == '__main__':
    sys.exit(main())
