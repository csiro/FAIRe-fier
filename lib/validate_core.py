##################################################
# FAIRe-fier VALIDATION CORE
# Author: Suk Yee Yong
#
# Reading the workbook, running the validator, grouping the warnings and errors,
# and writing the output workbooks. There is no Shiny here, so the Shiny app
# (validate_module.py) and the command line (fairefier_cli.py) share the same
# code and give the same results.
##################################################

from pathlib import Path
from pydantic import ValidationError

import io
import json
import numpy as np
import pandas as pd
import re
import warnings
warnings.formatwarning = lambda msg, *args, **kwargs: f'{msg}\n'

import pandas.io.formats.excel
pandas.io.formats.excel.ExcelFormatter.header_style = None

import lib.create_validator as create_validator
import config_data
term_name_col = config_data.term_name_col
samp_name_col = config_data.samp_name_col
warn_color = '#DDAA33'
error_color = '#BB5566'
failed_message = 'Failed! Please download the error and/or warning files, review them, and correct any errors before resubmitting the data for validation.'


class ReadError(Exception):
    """The workbook could not be read as FAIRe metadata"""


class WorksheetError(ReadError):
    """The projectMetadata or sampleMetadata worksheet is missing or cannot be read"""


class DuplicateSampleError(ReadError):
    """The same samp_name appears on more than one row of sampleMetadata"""


def read_worksheets(filepath_name):
    """Read Excel file from projectMetadata and sampleMetadata worksheets and returns DataFrames"""
    try:
        xlsx = pd.ExcelFile(filepath_name)
        df_xlsx_projectmetadata = pd.read_excel(filepath_name, sheet_name='projectMetadata_revise' if 'projectMetadata_revise' in xlsx.sheet_names else 'projectMetadata')
        # Skip few columns of info if load from template
        if 'projectMetadata' in xlsx.sheet_names:
            df_xlsx_projectmetadata = df_xlsx_projectmetadata.iloc[:,2:]
        skiprows = 0 if 'sampleMetadata_revise' in xlsx.sheet_names else 2
        df_xlsx_samplemetadata = pd.read_excel(filepath_name,
                                    sheet_name='sampleMetadata_revise' if 'sampleMetadata_revise' in xlsx.sheet_names else 'sampleMetadata',
                                    header=0, skiprows=skiprows)
    except Exception as e:
        raise WorksheetError(f"Worksheet named projectMetadata/projectMetadata_revise and/or sampleMetadata/sampleMetadata_revise not found or not readable ({e}).") from e
    if samp_name_col not in df_xlsx_samplemetadata.columns:
        raise WorksheetError(f"No {samp_name_col} column in sampleMetadata (field names are expected on row 3).")
    # Duplicate detection in samp_name_col
    samp_col = df_xlsx_samplemetadata[samp_name_col]
    dup_mask = samp_col.duplicated(keep=False)
    if dup_mask.any():
        dup_samp = samp_col[dup_mask]
        summary = ", ".join(
            f"'{samp}' ({count}rows: {[skiprows+2+dup_i for dup_i in dup_samp[dup_samp.isin([samp])].index]})"
            for samp, count in dup_samp.value_counts().items()
        )
        raise DuplicateSampleError(f"Duplicate {samp_name_col} in sampleMetadata worksheet: {summary}")
    # Successful file read
    return df_xlsx_projectmetadata, df_xlsx_samplemetadata


def combine_locmsgall(df, group_colname, ni_ingroup):
    """Combine rows with same loc and message across all"""
    # Check if 'loc' value is a list
    if df['loc'].apply(lambda x: isinstance(x, list)).all():
        df['loc'] = df['loc'].str[0] # First item in list is the term
    # Group by ['loc', 'msg']
    grouped = df.groupby(['loc', 'msg'])[group_colname].apply(lambda x: set(x)).reset_index()
    grouped[group_colname] = grouped[group_colname].apply(lambda x: 'ALL' if len(x) == ni_ingroup else None)
    df = df.merge(grouped, how='inner', on=['loc', 'msg'], suffixes=(None, '_all'))
    # Drop duplicates if 'ALL'
    df = df[~((df.duplicated(['loc', 'msg'])) & (df[f"{group_colname}_all"]=='ALL'))]
    df[group_colname] = df.apply(lambda row: row[group_colname] if row[f"{group_colname}_all"] is None else 'ALL', axis=1)
    df = df.drop(columns=[f"{group_colname}_all"])
    df.reset_index(drop=True, inplace=True)
    df.insert(0, group_colname, df.pop(group_colname))
    return df


def combine_difftypemsg(df, group_colname):
    """Combine rows of different type for the same group_colname, loc, and input, and remove associated string errors if required"""
    # Replace None with a placeholder string to handle None grouping correctly
    df['input'] = df['input'].fillna('None')
    df['input'] = df['input'].astype('string') # Convert values especially dict to str
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
        if len(list_type)>1 and (list_type[-1] == 'string_type') and bool(re.search(r'\bstring\b', list_msg[-1])):
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
    """Run metadata validator against the checklist file at checklist_path, and returns a dict of the validation message and DataFrames"""

    # Load checklist
    df_dc = pd.read_excel(Path(checklist_path), sheet_name='checklist', header=0)
    df_projectmetadata = df_dc[df_dc['data_type'].str.contains('projectMetadata')]
    df_samplemetadata = df_dc[df_dc['data_type'].str.contains('sampleMetadata')]
    df_samplemetadata = df_samplemetadata[df_samplemetadata[term_name_col]!=samp_name_col]

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
        df_out_projectmetadata = pd.DataFrame() # Don't output file if error
        df_error_projectmetadata = pd.concat(error_projectmetadata)
        df_error_projectmetadata = combine_locmsgall(df_error_projectmetadata, term_name_col, ni_termname)
        df_error_projectmetadata = combine_difftypemsg(df_error_projectmetadata, term_name_col)
        validator_message += failed_message

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
        df_out_samplemetadata = pd.DataFrame() # Don't output file if error
        df_error_samplemetadata = pd.concat(error_samplemetadata)
        df_error_samplemetadata = combine_locmsgall(df_error_samplemetadata, samp_name_col, ni_sampname)
        df_error_samplemetadata = combine_difftypemsg(df_error_samplemetadata, samp_name_col)
        validator_message += failed_message

    return {
        'message': validator_message,
        'in_projectmetadata': df_xlsx_projectmetadata.T,
        'out_projectmetadata': df_out_projectmetadata,
        'warn_projectmetadata': df_warn_projectmetadata,
        'error_projectmetadata': df_error_projectmetadata,
        'in_samplemetadata': df_xlsx_samplemetadata,
        'out_samplemetadata': df_out_samplemetadata,
        'warn_samplemetadata': df_warn_samplemetadata,
        'error_samplemetadata': df_error_samplemetadata,
    }


def write_metadata(df_out_projectmetadata, df_out_samplemetadata, target):
    """Write the validated metadata (fairefier_metadata.xlsx) to target, a file path or a buffer"""
    with pd.ExcelWriter(target, engine='openpyxl') as writer:
        df_out_projectmetadata.to_excel(writer, sheet_name='projectMetadata', index=False)
        df_out_samplemetadata.to_excel(writer, sheet_name='sampleMetadata', index=False)


def highlight_cells_bysampnameloc(df, df_err, group_colname, color, group_colname_asheader=False):
    """Highlight cells by group_colname and 'loc'"""
    style = pd.DataFrame('', index=df.index, columns=df.columns)
    for gcn, lc in df_err[[group_colname, 'loc']].values:
        gcn, lc = gcn.strip(), lc.strip()
        # Entries in group_colname are headers
        if group_colname_asheader:
            style.loc[(df[group_colname]==lc), gcn] = f"background-color: {color}"
        # Entries in group_colname are indices
        else:
            style.loc[(df[group_colname]==gcn), lc] = f"background-color: {color}"
    return style


def write_warn_error(df_xlsx_projectmetadata, df_warn_projectmetadata, df_error_projectmetadata,
                     df_xlsx_samplemetadata, df_warn_samplemetadata, df_error_samplemetadata, target):
    """Write the warnings, errors and highlighted _revise worksheets (fairefier_warn_error.xlsx) to target, a file path or a buffer"""
    # Color cells to be revised
    styler_projectmetadata = df_xlsx_projectmetadata.style
    styler_samplemetadata = df_xlsx_samplemetadata.style
    # Color column header if 'ALL', otherwise color cell
    if not df_warn_projectmetadata.empty:
        styler_projectmetadata.apply(lambda x: np.where(x.isin(df_warn_projectmetadata[df_warn_projectmetadata[term_name_col]=='ALL']['loc'].str.strip().unique()), f"background-color: {warn_color}", ""), axis=None)\
                            .apply(lambda df: highlight_cells_bysampnameloc(df, df_warn_projectmetadata[df_warn_projectmetadata[term_name_col]!='ALL'], term_name_col, warn_color, group_colname_asheader=True), axis=None)
    if not df_error_projectmetadata.empty:
        styler_projectmetadata.apply(lambda x: np.where(x.isin(df_error_projectmetadata[df_error_projectmetadata[term_name_col]=='ALL']['loc'].str.strip().unique()), f"background-color: {error_color}", ""), axis=None)\
                            .apply(lambda df: highlight_cells_bysampnameloc(df, df_error_projectmetadata[df_error_projectmetadata[term_name_col]!='ALL'], term_name_col, error_color, group_colname_asheader=True), axis=None)
    if not df_warn_samplemetadata.empty:
        styler_samplemetadata.apply_index(lambda x: np.where(x.isin(df_warn_samplemetadata[df_warn_samplemetadata[samp_name_col]=='ALL']['loc'].str.strip().unique()), f"background-color: {warn_color}", ""), axis=1)\
                            .apply(lambda df: highlight_cells_bysampnameloc(df, df_warn_samplemetadata[df_warn_samplemetadata[samp_name_col]!='ALL'], samp_name_col, warn_color), axis=None)
    if not df_error_samplemetadata.empty:
        styler_samplemetadata.apply_index(lambda x: np.where(x.isin(df_error_samplemetadata[df_error_samplemetadata[samp_name_col]=='ALL']['loc'].str.strip().unique()), f"background-color: {error_color}", ""), axis=1)\
                            .apply(lambda df: highlight_cells_bysampnameloc(df, df_error_samplemetadata[df_error_samplemetadata[samp_name_col]!='ALL'], samp_name_col, error_color), axis=None)
    # Write to Excel
    with pd.ExcelWriter(target, engine='openpyxl') as writer:
        (df_warn_projectmetadata if df_warn_projectmetadata.empty \
            else df_warn_projectmetadata.style.map_index(lambda _: f"background-color: {warn_color}", axis='columns'))\
            .to_excel(writer, sheet_name='projectMetadata_warning', index=False)
        (df_error_projectmetadata if df_error_projectmetadata.empty \
            else df_error_projectmetadata.style.map_index(lambda _: f"background-color: {error_color}", axis='columns'))\
            .to_excel(writer, sheet_name='projectMetadata_error', index=False)
        styler_projectmetadata.to_excel(writer, sheet_name='projectMetadata_revise', header=True, index=False)
        (df_warn_samplemetadata if df_warn_samplemetadata.empty \
            else df_warn_samplemetadata.style.map_index(lambda _: f"background-color: {warn_color}", axis='columns'))\
            .to_excel(writer, sheet_name='sampleMetadata_warning', index=False)
        (df_error_samplemetadata if df_error_samplemetadata.empty \
            else df_error_samplemetadata.style.map_index(lambda _: f"background-color: {error_color}", axis='columns'))\
            .to_excel(writer, sheet_name='sampleMetadata_error', index=False)
        styler_samplemetadata.to_excel(writer, sheet_name='sampleMetadata_revise', header=True, index=False)
