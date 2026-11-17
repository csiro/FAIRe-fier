##################################################
# FAIRe-fier SHINY VALIDATE MODULE
# Author: Suk Yee Yong
#
# The Shiny parts of the app. Reading, validation and writing of the outputs
# are in lib/validate_core.py, shared with the command line (fairefier_cli.py).
##################################################

from pathlib import Path
from shiny import Inputs, Outputs, Session, reactive, render, ui

import io
import pandas as pd

import lib.validate_core as validate_core
import config_data
data_dir = config_data.data_dir
dict_checklistv = config_data.dict_checklistv
samp_name_col = config_data.samp_name_col


def validate_server(input: Inputs, output: Outputs, session: Session):
    @reactive.calc
    def read_worksheets():
        """Read Excel file from projectMetadata and sampleMetadata worksheets and returns DataFrames"""
        filepath = input.file_upload()
        filepath_name = filepath[0]['datapath']
        if filepath_name.endswith('.xlsx') or filepath_name.endswith('.xls'):
            try:
                return validate_core.read_worksheets(filepath_name)
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
        return validate_core.run_validator(Path(data_dir, dict_checklistv[filename_version]),
                                           rdf_input_projectmetadata.get(), rdf_input_samplemetadata.get())
    
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
            rtextbox_display.set(out_run_validator['message'])
            rdf_in_projectmetadata.set(out_run_validator['in_projectmetadata'])
            rdf_out_projectmetadata.set(out_run_validator['out_projectmetadata'])
            rdf_warn_projectmetadata.set(out_run_validator['warn_projectmetadata'])
            rdf_error_projectmetadata.set(out_run_validator['error_projectmetadata'])
            rdf_in_samplemetadata.set(out_run_validator['in_samplemetadata'])
            rdf_out_samplemetadata.set(out_run_validator['out_samplemetadata'])
            rdf_warn_samplemetadata.set(out_run_validator['warn_samplemetadata'])
            rdf_error_samplemetadata.set(out_run_validator['error_samplemetadata'])
    
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
        buffer = io.BytesIO()
        validate_core.write_metadata(rdf_out_projectmetadata.get(), rdf_out_samplemetadata.get(), buffer)
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
        buffer = io.BytesIO()
        validate_core.write_warn_error(rdf_input_projectmetadata.get(), rdf_warn_projectmetadata.get(), rdf_error_projectmetadata.get(),
                                       rdf_input_samplemetadata.get(), rdf_warn_samplemetadata.get(), rdf_error_samplemetadata.get(),
                                       buffer)
        buffer.seek(0)
        return buffer
