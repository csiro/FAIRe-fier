##################################################
# FAIRe-fier SHINY VALIDATE MODULE
# Author: Suk Yee Yong
##################################################

from pathlib import Path
from pydantic import ValidationError
from shiny import Inputs, Outputs, Session, reactive, render, ui

import io
import json
import numpy as np
import pandas as pd
import re
import warnings
warnings.formatwarning = lambda msg, *args, **kwargs: f'{msg}\n'

import pandas.io.formats.excel
pandas.io.formats.excel.ExcelFormatter.header_style = None

import create_validator as create_validator
import config_data
data_dir = config_data.data_dir
dict_checklistv = config_data.dict_checklistv
term_name_col = config_data.term_name_col
samp_name_col = config_data.samp_name_col
warn_color = '#DDAA33'
error_color = '#BB5566'


def validate_server(input: Inputs, output: Outputs, session: Session):
    @reactive.calc
    def read_worksheets():
        """Read Excel file from projectMetadata and sampleMetadata worksheets and returns DataFrames"""
        filepath = input.file_upload()
        filepath_name = filepath[0]['datapath']
        if filepath_name.endswith('.xlsx') or filepath_name.endswith('.xls'):
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
                # Duplicate detection in samp_name_col
                samp_col = df_xlsx_samplemetadata[samp_name_col]
                dup_mask = samp_col.duplicated(keep=False)
                if dup_mask.any():
                    dup_samp = samp_col[dup_mask]
                    summary = ", ".join(
                        f"'{samp}' ({count}rows: {[skiprows+2+dup_i for dup_i in dup_samp[dup_samp.isin([samp])].index]})"
                        for samp, count in dup_samp.value_counts().items()
                    )
                    return None, f"Duplicate {samp_name_col} in sampleMetadata worksheet: {summary}"
                # Successful file read
                return df_xlsx_projectmetadata, df_xlsx_samplemetadata
            except Exception as e:
                return None, str(e)
        else:
            return None, "Unsupported file type"
    
    @reactive.effect
    @reactive.event(input.file_upload)
    def handle_file_upload():
        result, error_message = read_worksheets()
        if result is None:
            if error_message == "Unsupported file type":
                ui.notification_show("Unsupported file type. Please upload single Excel file (.xlsx/.xls) with 'projectMetadata' and 'sampleMetadata' worksheets.", type='error', duration=15)
            elif f"Duplicate {samp_name_col}" in error_message:
                ui.notification_show(f"{error_message}. Please ensure sample names are unique.", type='error', duration=60*3)
            else:
                ui.notification_show("Worksheet named projectMetadata/projectMetadata_revise and/or sampleMetadata/sampleMetadata_revise not found! Please include those Excel worksheets.", type='error', duration=15)
        else:
            # Handle successful file read here
            pass
    
    @reactive.calc
    def run_validator():
        """Run metadata validator"""
        
        filename_version = input.dropdown_checklistv.get()
        df_xlsx_projectmetadata = rdf_input_projectmetadata.get()
        df_xlsx_samplemetadata = rdf_input_samplemetadata.get()
        
        # Load checklist
        df_dc = pd.read_excel(Path(data_dir, dict_checklistv[filename_version]), sheet_name='checklist', header=0)
        df_projectmetadata = df_dc[df_dc['data_type'].str.contains('projectMetadata')]
        df_samplemetadata = df_dc[df_dc['data_type'].str.contains('sampleMetadata')]
        df_samplemetadata = df_samplemetadata[df_samplemetadata[term_name_col]!=samp_name_col]
        
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
            validator_message += 'Failed! Please download the error and/or warning files, review them, and correct any errors before resubmitting the data for validation.'
        
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
            validator_message += 'Failed! Please download the error and/or warning files, review them, and correct any errors before resubmitting the data for validation.'
        
        return [validator_message,
                df_xlsx_projectmetadata.T, df_out_projectmetadata, df_warn_projectmetadata, df_error_projectmetadata,
                df_xlsx_samplemetadata, df_out_samplemetadata, df_warn_samplemetadata, df_error_samplemetadata,
        ]
    
    # Store outputs from read_worksheets()
    rfile_upload = reactive.value(False)
    rdf_input_projectmetadata = reactive.value(pd.DataFrame())
    rdf_input_samplemetadata = reactive.value(pd.DataFrame())
    # Store outputs from run_validator()
    rtextbox_display = reactive.value('')
    rdf_in_projectmetadata = reactive.value(pd.DataFrame())
    rdf_out_projectmetadata = reactive.value(pd.DataFrame())
    rdf_warn_projectmetadata = reactive.value(pd.DataFrame())
    rdf_error_projectmetadata = reactive.value(pd.DataFrame())
    rdf_in_samplemetadata = reactive.value(pd.DataFrame())
    rdf_out_samplemetadata = reactive.value(pd.DataFrame())
    rdf_warn_samplemetadata = reactive.value(pd.DataFrame())
    rdf_error_samplemetadata = reactive.value(pd.DataFrame())
    
    @render.ui
    @reactive.event(rfile_upload)
    def ui_file_upload():
        return ui.input_file('file_upload', label='Upload Excel file', multiple=False, accept=['.xlsx'], width='60%')
    
    @render.data_frame
    @reactive.event(rdf_input_projectmetadata)
    def dataframe_input_projectmetadata():
        return rdf_input_projectmetadata.get()
    
    @render.data_frame
    @reactive.event(rdf_input_samplemetadata)
    def dataframe_input_samplemetadata():
        return rdf_input_samplemetadata.get()
    
    @render.text
    @reactive.event(rtextbox_display)
    def textbox_display():
        return rtextbox_display.get()
    
    @render.data_frame
    @reactive.event(rdf_out_projectmetadata)
    def dataframe_out_projectmetadata():
        return rdf_out_projectmetadata.get()
    
    @render.data_frame
    @reactive.event(rdf_warn_projectmetadata)
    def dataframe_warn_projectmetadata():
       return rdf_warn_projectmetadata.get()
    
    @render.data_frame
    @reactive.event(rdf_error_projectmetadata)
    def dataframe_error_projectmetadata():
        return rdf_error_projectmetadata.get()
    
    @render.data_frame
    @reactive.event(rdf_out_samplemetadata)
    def dataframe_out_samplemetadata():
        return rdf_out_samplemetadata.get()
    
    @render.data_frame
    @reactive.event(rdf_warn_samplemetadata)
    def dataframe_warn_samplemetadata():
        return rdf_warn_samplemetadata.get()
    
    @render.data_frame
    @reactive.event(rdf_error_samplemetadata)
    def dataframe_error_samplemetadata():
        return rdf_error_samplemetadata.get()
    
    @reactive.effect
    @reactive.event(input.button_submit)
    def on_button_submit():
        """Event trigger Submit button"""
        filepath = input.file_upload()
        if not filepath:
            ui.notification_show('No file submitted! Please drop or upload file.', type='error')
        else:
            out_read_worksheets = read_worksheets()
            rdf_input_projectmetadata.set(out_read_worksheets[0])
            rdf_input_samplemetadata.set(out_read_worksheets[1])
            out_run_validator = run_validator()
            rtextbox_display.set(out_run_validator[0])
            rdf_in_projectmetadata.set(out_run_validator[1])
            rdf_out_projectmetadata.set(out_run_validator[2])
            rdf_warn_projectmetadata.set(out_run_validator[3])
            rdf_error_projectmetadata.set(out_run_validator[4])
            rdf_in_samplemetadata.set(out_run_validator[5])
            rdf_out_samplemetadata.set(out_run_validator[6])
            rdf_warn_samplemetadata.set(out_run_validator[7])
            rdf_error_samplemetadata.set(out_run_validator[8])
    
    @reactive.effect
    @reactive.event(input.button_clear)
    def on_button_clear():
        rfile_upload.set(not rfile_upload())
        rdf_input_projectmetadata.set(None)
        rdf_input_samplemetadata.set(None)
        rtextbox_display.set(None)
        rdf_in_projectmetadata.set(None)
        rdf_out_projectmetadata.set(None)
        rdf_warn_projectmetadata.set(None)
        rdf_error_projectmetadata.set(None)
        rdf_in_samplemetadata.set(None)
        rdf_out_samplemetadata.set(None)
        rdf_warn_samplemetadata.set(None)
        rdf_error_samplemetadata.set(None)
    
    @render.ui
    @reactive.event(rdf_out_projectmetadata)
    @reactive.event(rdf_out_samplemetadata)
    def ui_ddl_out():
        if not rdf_out_projectmetadata.get().empty or not rdf_out_samplemetadata.get().empty:
            return ui.download_button('ddl_out', 'Metadata', class_='btn-success')
    
    @render.download(filename="fairefier_metadata.xlsx")
    def ddl_out():
        df_out_projectmetadata = rdf_out_projectmetadata.get()
        df_out_samplemetadata = rdf_out_samplemetadata.get()
        # savefilename_out = 'fairefier_metadata.xlsx'
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_out_projectmetadata.to_excel(writer, sheet_name='projectMetadata', index=False)
            df_out_samplemetadata.to_excel(writer, sheet_name='sampleMetadata', index=False)
        buffer.seek(0)
        return buffer
    
    @render.ui
    @reactive.event(rdf_warn_projectmetadata)
    @reactive.event(rdf_error_projectmetadata)
    @reactive.event(rdf_warn_samplemetadata)
    @reactive.event(rdf_error_samplemetadata)
    def ui_ddl_warn_error():
        if not rdf_warn_projectmetadata.get().empty or not rdf_error_projectmetadata.get().empty \
            or not rdf_warn_samplemetadata.get().empty or not rdf_error_samplemetadata.get().empty:
            return ui.download_button('ddl_warn_error', 'Warnings+Errors', class_='btn-info')
    
    @render.download(filename="fairefier_warn_error.xlsx")
    def ddl_warn_error():
        df_xlsx_projectmetadata = rdf_input_projectmetadata.get()
        df_warn_projectmetadata = rdf_warn_projectmetadata.get()
        df_error_projectmetadata = rdf_error_projectmetadata.get()
        df_xlsx_samplemetadata = rdf_input_samplemetadata.get()
        df_warn_samplemetadata = rdf_warn_samplemetadata.get()
        df_error_samplemetadata = rdf_error_samplemetadata.get()
        # savefilename_warn_error = 'fairefier_warn_error.xlsx'
        
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
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
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
        buffer.seek(0)
        return buffer
