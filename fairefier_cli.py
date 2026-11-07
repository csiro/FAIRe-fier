##################################################
# FAIRe-fier COMMAND-LINE INTERFACE
#
# Runs the FAIRe-fier validation without the Shiny app, for scripts, pipelines
# and workflow systems such as Galaxy. The reading, validation, grouping and
# Excel-writing steps are taken unchanged from validate_module.py
# (validate_server: read_worksheets, run_validator, ddl_out, ddl_warn_error),
# with the Shiny reactive values replaced by function arguments. The validation
# rules themselves come from create_validator.py, which is imported as is.
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
import io
import json
import re
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import pandas.io.formats.excel
from pydantic import ValidationError

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))
import create_validator  # noqa: E402
import config_data  # noqa: E402

warnings.formatwarning = lambda msg, *args, **kwargs: f'{msg}\n'
pandas.io.formats.excel.ExcelFormatter.header_style = None

term_name_col = config_data.term_name_col
samp_name_col = config_data.samp_name_col
warn_color = '#DDAA33'
error_color = '#BB5566'

FAILED = 'Failed! Please download the error and/or warning files, review them, and correct any errors before resubmitting the data for validation.'


class ReadError(Exception):
    """The workbook could not be read as FAIRe metadata."""


# ---------------------------------------------------------------- reading

def read_worksheets(filepath_name):
    """Read projectMetadata and sampleMetadata worksheets (from validate_module.read_worksheets)."""
    try:
        xlsx = pd.ExcelFile(filepath_name)
        df_xlsx_projectmetadata = pd.read_excel(filepath_name, sheet_name='projectMetadata_revise' if 'projectMetadata_revise' in xlsx.sheet_names else 'projectMetadata')
        # Skip few columns of info if load from template
        if 'projectMetadata' in xlsx.sheet_names:
            df_xlsx_projectmetadata = df_xlsx_projectmetadata.iloc[:, 2:]
        skiprows = 0 if 'sampleMetadata_revise' in xlsx.sheet_names else 2
        df_xlsx_samplemetadata = pd.read_excel(filepath_name,
                                               sheet_name='sampleMetadata_revise' if 'sampleMetadata_revise' in xlsx.sheet_names else 'sampleMetadata',
                                               header=0, skiprows=skiprows)
    except Exception as e:
        raise ReadError(f"Worksheet named projectMetadata/projectMetadata_revise and/or sampleMetadata/sampleMetadata_revise not found or not readable ({e}).")
    # Duplicate detection in samp_name_col
    if samp_name_col not in df_xlsx_samplemetadata.columns:
        raise ReadError(f"No {samp_name_col} column in sampleMetadata (field names are expected on row 3).")
    samp_col = df_xlsx_samplemetadata[samp_name_col]
    dup_mask = samp_col.duplicated(keep=False)
    if dup_mask.any():
        dup_samp = samp_col[dup_mask]
        summary = ", ".join(
            f"'{samp}' ({count}rows: {[skiprows+2+dup_i for dup_i in dup_samp[dup_samp.isin([samp])].index]})"
            for samp, count in dup_samp.value_counts().items()
        )
        raise ReadError(f"Duplicate {samp_name_col} in sampleMetadata worksheet: {summary}. Please ensure sample names are unique.")
    return df_xlsx_projectmetadata, df_xlsx_samplemetadata


# ---------------------------------------------------------------- validation

def combine_locmsgall(df, group_colname, ni_ingroup):
    """Combine rows with same loc and message across all"""
    # Check if 'loc' value is a list
    if df['loc'].apply(lambda x: isinstance(x, list)).all():
        df['loc'] = df['loc'].str[0]  # First item in list is the term
    # Group by ['loc', 'msg']
    grouped = df.groupby(['loc', 'msg'])[group_colname].apply(lambda x: set(x)).reset_index()
    grouped[group_colname] = grouped[group_colname].apply(lambda x: 'ALL' if len(x) == ni_ingroup else None)
    df = df.merge(grouped, how='inner', on=['loc', 'msg'], suffixes=(None, '_all'))
    # Drop duplicates if 'ALL'
    df = df[~((df.duplicated(['loc', 'msg'])) & (df[f"{group_colname}_all"] == 'ALL'))]
    df[group_colname] = df.apply(lambda row: row[group_colname] if row[f"{group_colname}_all"] is None else 'ALL', axis=1)
    df = df.drop(columns=[f"{group_colname}_all"])
    df.reset_index(drop=True, inplace=True)
    df.insert(0, group_colname, df.pop(group_colname))
    return df


def combine_difftypemsg(df, group_colname):
    """Combine rows of different type for the same group_colname, loc, and input, and remove associated string errors if required"""
    # Replace None with a placeholder string to handle None grouping correctly
    df['input'] = df['input'].fillna('None')
    df['input'] = df['input'].astype('string')  # Convert values especially dict to str
    split_char = '; '
    df = df.groupby([group_colname, 'loc', 'input'], as_index=False).agg({
        'type': lambda x: split_char.join(x.unique()),
        'msg': lambda x: split_char.join(x.unique()),
    })

    def remove_lastrowstr(row):
        """Remove string error if in the last row of the grouped results"""
        list_type = row['type'].split(split_char)
        list_msg = row['msg'].split(split_char)
        # Check if associated str error is at the end and remove
        if len(list_type) > 1 and (list_type[-1] == 'string_type') and bool(re.search(r'\bstring\b', list_msg[-1])):
            list_type.pop()
            list_msg.pop()
            # Join back into string
            row['type'] = split_char.join(list_type)
            row['msg'] = split_char.join(list_msg)
        return row
    df = df.apply(remove_lastrowstr, axis=1)
    # Revert the placeholder 'None' back to actual None
    df['input'] = df['input'].replace('None', None)
    return df


def get_nonempty_term(input_dict):
    """Extract non-empty term_type in both project_level and in all other dictionaries for projectMetadata"""
    result = set()
    # Extract non-empty in project_level first
    result.update(k for k, v in input_dict['project_level'].items() if v)
    # Extract non-empty in all other columns
    other_dicts = {key: value for key, value in input_dict.items() if key != 'project_level'}
    if other_dicts:
        for subkey in next(iter(other_dicts.values())).keys():
            if all(subkey in dictionary and dictionary[subkey] for dictionary in other_dicts.values()):
                result.add(subkey)
    return result


def run_validator(checklist_path, df_xlsx_projectmetadata, df_xlsx_samplemetadata):
    """Run metadata validator (from validate_module.run_validator)"""
    # Load checklist
    df_dc = pd.read_excel(checklist_path, sheet_name='checklist', header=0)
    df_projectmetadata = df_dc[df_dc['data_type'].str.contains('projectMetadata')]
    df_samplemetadata = df_dc[df_dc['data_type'].str.contains('sampleMetadata')]
    df_samplemetadata = df_samplemetadata[df_samplemetadata[term_name_col] != samp_name_col]

    # projectMetadata
    validator_message = 'For projectMetadata: '
    df_xlsx_projectmetadata = df_xlsx_projectmetadata.fillna('').set_index(term_name_col).T
    dicts_xlsx_projectmetadata = df_xlsx_projectmetadata.to_dict('index')
    ni_termname = len(df_xlsx_projectmetadata.index.unique())
    df_out_projectmetadata = pd.DataFrame()
    df_error_projectmetadata = pd.DataFrame()
    df_warn_projectmetadata = pd.DataFrame()
    out_projectmetadata = []
    warn_projectmetadata, error_projectmetadata = [], []

    # Get filled mandatory in project_level and the rest in assay
    dict_meta = create_validator.create_dictterms(df_projectmetadata)
    list_term_mandatory = set(k for k, v in dict_meta.items() if v[1].is_required())
    list_term_tooptional = list_term_mandatory.intersection(get_nonempty_term(dicts_xlsx_projectmetadata))

    for study_assay, dict_xlsx_projectmetadata in dicts_xlsx_projectmetadata.items():
        try:
            # Catch all warnings
            with warnings.catch_warnings(record=True) as warning_list_projectmetadata:
                ednameta_projectmetadata = create_validator.validate_projectMetadata(df_projectmetadata, dict_meta, dict_xlsx_projectmetadata,
                                                                                     mandatory_to_optional=list_term_tooptional,
                                                                                     )
            # Output table
            df_out_projectmetadata = pd.DataFrame([ednameta_projectmetadata], index=[study_assay])
            out_projectmetadata.append(df_out_projectmetadata)
        except ValidationError as exc:
            error = exc.json(include_url=False, include_context=False)
            df_error_projectmetadata = pd.DataFrame(json.loads(error))
            df_error_projectmetadata.insert(0, term_name_col, study_assay)
            error_projectmetadata.append(df_error_projectmetadata)
        finally:
            if warning_list_projectmetadata:
                df_warn_projectmetadata = pd.read_csv(io.StringIO('\n'.join([wn.message.args[0] for wn in warning_list_projectmetadata])), sep="|", names=['loc', 'msg', 'input', 'output'])
                df_warn_projectmetadata.insert(0, term_name_col, study_assay)
                warn_projectmetadata.append(df_warn_projectmetadata)

    if out_projectmetadata and (not error_projectmetadata):
        df_out_projectmetadata = pd.concat(out_projectmetadata).T.rename_axis(term_name_col).reset_index()
        validator_message += 'Success!'
    if warn_projectmetadata:
        df_warn_projectmetadata = pd.concat(warn_projectmetadata)
        df_warn_projectmetadata = combine_locmsgall(df_warn_projectmetadata, term_name_col, ni_termname)
    if error_projectmetadata:
        df_out_projectmetadata = pd.DataFrame()  # Don't output file if error
        df_error_projectmetadata = pd.concat(error_projectmetadata)
        df_error_projectmetadata = combine_locmsgall(df_error_projectmetadata, term_name_col, ni_termname)
        df_error_projectmetadata = combine_difftypemsg(df_error_projectmetadata, term_name_col)
        validator_message += FAILED

    # sampleMetadata
    validator_message += '\nFor sampleMetadata: '
    df_xlsx_samplemetadata = df_xlsx_samplemetadata.fillna('').set_index(samp_name_col)
    dicts_xlsx_samplemetadata = df_xlsx_samplemetadata.to_dict('index')
    ni_sampname = len(df_xlsx_samplemetadata.index.unique())
    df_out_samplemetadata = pd.DataFrame()
    df_error_samplemetadata = pd.DataFrame()
    df_warn_samplemetadata = pd.DataFrame()
    out_samplemetadata = []
    warn_samplemetadata, error_samplemetadata = [], []
    # Check informationWithheld for latitude/longitude
    informationwithheld_latlon = False
    input_informationwithheld = dicts_xlsx_projectmetadata['project_level'].get('informationWithheld', None)
    if isinstance(input_informationwithheld, str):
        latlon_keywords = ['latitude', 'longitude', 'lat', 'long', 'gps', 'coordinate', 'location', 'site']
        if set(latlon_keywords).intersection(set(input_informationwithheld.lower().split())):
            informationwithheld_latlon = True
    for i, (samp_name, dict_xlsx_samplemetadata) in enumerate(dicts_xlsx_samplemetadata.items()):
        try:
            # Catch all warnings
            with warnings.catch_warnings(record=True) as warning_list_samplemetadata:
                if informationwithheld_latlon:
                    warnings.warn("decimalLatitude | Latitude and/or longitude is mentioned in informationWithheld. | Mandatory | Optional")
                    warnings.warn("decimalLongitude | Latitude and/or longitude is mentioned in informationWithheld. | Mandatory | Optional")
                dict_meta = create_validator.create_dictterms(df_samplemetadata)
                ednameta_samplemetadata = create_validator.validate_sampleMetadata(df_samplemetadata, dict_meta, dict_xlsx_samplemetadata,
                                                                                   inputs_pm=dicts_xlsx_projectmetadata,
                                                                                   informationwithheld_latlon=informationwithheld_latlon,
                                                                                   )
            # Output table
            df_out_samplemetadata = pd.DataFrame([ednameta_samplemetadata], index=[samp_name])
            out_samplemetadata.append(df_out_samplemetadata)
        except ValidationError as exc:
            # Errors
            error = exc.json(include_url=False, include_context=False)
            df_error_samplemetadata = pd.DataFrame(json.loads(error))
            df_error_samplemetadata.insert(0, samp_name_col, samp_name)
            error_samplemetadata.append(df_error_samplemetadata)
        finally:
            # Warnings
            if warning_list_samplemetadata:
                df_warn_samplemetadata = pd.read_csv(io.StringIO('\n'.join([wn.message.args[0] for wn in warning_list_samplemetadata])), sep="|", names=['loc', 'msg', 'input', 'output'])
                df_warn_samplemetadata.insert(0, samp_name_col, samp_name)
                warn_samplemetadata.append(df_warn_samplemetadata)

    if out_samplemetadata and (not error_samplemetadata):
        df_out_samplemetadata = pd.concat(out_samplemetadata).rename_axis(samp_name_col).reset_index()
        validator_message += 'Success!'
    if warn_samplemetadata:
        df_warn_samplemetadata = pd.concat(warn_samplemetadata)
        df_warn_samplemetadata = combine_locmsgall(df_warn_samplemetadata, samp_name_col, ni_sampname)
    if error_samplemetadata:
        df_out_samplemetadata = pd.DataFrame()  # Don't output file if error
        df_error_samplemetadata = pd.concat(error_samplemetadata)
        df_error_samplemetadata = combine_locmsgall(df_error_samplemetadata, samp_name_col, ni_sampname)
        df_error_samplemetadata = combine_difftypemsg(df_error_samplemetadata, samp_name_col)
        validator_message += FAILED

    return {
        'message': validator_message,
        'out_projectmetadata': df_out_projectmetadata,
        'warn_projectmetadata': df_warn_projectmetadata,
        'error_projectmetadata': df_error_projectmetadata,
        'out_samplemetadata': df_out_samplemetadata,
        'warn_samplemetadata': df_warn_samplemetadata,
        'error_samplemetadata': df_error_samplemetadata,
    }


# ---------------------------------------------------------------- writing

def write_metadata(res, path):
    """Validated metadata workbook (from validate_module.ddl_out)"""
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        res['out_projectmetadata'].to_excel(writer, sheet_name='projectMetadata', index=False)
        res['out_samplemetadata'].to_excel(writer, sheet_name='sampleMetadata', index=False)


def highlight_cells_bysampnameloc(df, df_err, group_colname, color, group_colname_asheader=False):
    """Highlight cells by group_colname and 'loc'"""
    style = pd.DataFrame('', index=df.index, columns=df.columns)
    for gcn, lc in df_err[[group_colname, 'loc']].values:
        gcn, lc = gcn.strip(), lc.strip()
        # Entries in group_colname are headers
        if group_colname_asheader:
            style.loc[(df[group_colname] == lc), gcn] = f"background-color: {color}"
        # Entries in group_colname are indices
        else:
            style.loc[(df[group_colname] == gcn), lc] = f"background-color: {color}"
    return style


def write_warn_error(res, df_xlsx_projectmetadata, df_xlsx_samplemetadata, path):
    """Warnings, errors and highlighted revise sheets (from validate_module.ddl_warn_error)"""
    df_warn_projectmetadata = res['warn_projectmetadata']
    df_error_projectmetadata = res['error_projectmetadata']
    df_warn_samplemetadata = res['warn_samplemetadata']
    df_error_samplemetadata = res['error_samplemetadata']

    # Color cells to be revised
    styler_projectmetadata = df_xlsx_projectmetadata.style
    styler_samplemetadata = df_xlsx_samplemetadata.style
    # Color column header if 'ALL', otherwise color cell
    if not df_warn_projectmetadata.empty:
        styler_projectmetadata.apply(lambda x: np.where(x.isin(df_warn_projectmetadata[df_warn_projectmetadata[term_name_col] == 'ALL']['loc'].str.strip().unique()), f"background-color: {warn_color}", ""), axis=None)\
            .apply(lambda df: highlight_cells_bysampnameloc(df, df_warn_projectmetadata[df_warn_projectmetadata[term_name_col] != 'ALL'], term_name_col, warn_color, group_colname_asheader=True), axis=None)
    if not df_error_projectmetadata.empty:
        styler_projectmetadata.apply(lambda x: np.where(x.isin(df_error_projectmetadata[df_error_projectmetadata[term_name_col] == 'ALL']['loc'].str.strip().unique()), f"background-color: {error_color}", ""), axis=None)\
            .apply(lambda df: highlight_cells_bysampnameloc(df, df_error_projectmetadata[df_error_projectmetadata[term_name_col] != 'ALL'], term_name_col, error_color, group_colname_asheader=True), axis=None)
    if not df_warn_samplemetadata.empty:
        styler_samplemetadata.apply_index(lambda x: np.where(x.isin(df_warn_samplemetadata[df_warn_samplemetadata[samp_name_col] == 'ALL']['loc'].str.strip().unique()), f"background-color: {warn_color}", ""), axis=1)\
            .apply(lambda df: highlight_cells_bysampnameloc(df, df_warn_samplemetadata[df_warn_samplemetadata[samp_name_col] != 'ALL'], samp_name_col, warn_color), axis=None)
    if not df_error_samplemetadata.empty:
        styler_samplemetadata.apply_index(lambda x: np.where(x.isin(df_error_samplemetadata[df_error_samplemetadata[samp_name_col] == 'ALL']['loc'].str.strip().unique()), f"background-color: {error_color}", ""), axis=1)\
            .apply(lambda df: highlight_cells_bysampnameloc(df, df_error_samplemetadata[df_error_samplemetadata[samp_name_col] != 'ALL'], samp_name_col, error_color), axis=None)
    # Write to Excel
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        (df_warn_projectmetadata if df_warn_projectmetadata.empty
            else df_warn_projectmetadata.style.map_index(lambda _: f"background-color: {warn_color}", axis='columns'))\
            .to_excel(writer, sheet_name='projectMetadata_warning', index=False)
        (df_error_projectmetadata if df_error_projectmetadata.empty
            else df_error_projectmetadata.style.map_index(lambda _: f"background-color: {error_color}", axis='columns'))\
            .to_excel(writer, sheet_name='projectMetadata_error', index=False)
        styler_projectmetadata.to_excel(writer, sheet_name='projectMetadata_revise', header=True, index=False)
        (df_warn_samplemetadata if df_warn_samplemetadata.empty
            else df_warn_samplemetadata.style.map_index(lambda _: f"background-color: {warn_color}", axis='columns'))\
            .to_excel(writer, sheet_name='sampleMetadata_warning', index=False)
        (df_error_samplemetadata if df_error_samplemetadata.empty
            else df_error_samplemetadata.style.map_index(lambda _: f"background-color: {error_color}", axis='columns'))\
            .to_excel(writer, sheet_name='sampleMetadata_error', index=False)
        styler_samplemetadata.to_excel(writer, sheet_name='sampleMetadata_revise', header=True, index=False)


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
        df_pm, df_sm = read_worksheets(args.workbook)
        version, checklist_path, note = choose_checklist(args.checklist_version, args.checklist_dir, df_pm)
    except ReadError as e:
        sys.stderr.write(f"ERROR: {e}\n")
        return 2
    print(f"INFO: {note}")

    res = run_validator(checklist_path, df_pm, df_sm)
    print(res['message'])

    written = []
    if not res['out_projectmetadata'].empty or not res['out_samplemetadata'].empty:
        write_metadata(res, outdir / 'fairefier_metadata.xlsx')
        written.append('fairefier_metadata.xlsx')
    if any(not res[k].empty for k in ('warn_projectmetadata', 'error_projectmetadata', 'warn_samplemetadata', 'error_samplemetadata')):
        write_warn_error(res, df_pm, df_sm, outdir / 'fairefier_warn_error.xlsx')
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
