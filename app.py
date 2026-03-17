##################################################
# FAIRe-fier: FAIR eDNA METADATA VERIFIER SHINY APP
# Author: Suk Yee Yong
##################################################

from pathlib import Path
from shiny import App, ui

from validate_module import validate_server
import config_data
dict_checklistv = config_data.dict_checklistv


app_ui = ui.page_fluid(
    ui.markdown("# FAIRe-fier: FAIR eDNA metadata verifier"),
    
    # Info
    ui.accordion(
        ui.accordion_panel('Quick start', ui.markdown("""\
            0. Download and edit template.
            1. Select checklist version of the template.
            2. Drop or upload a single Excel file containing `projectMetadata` and `sampleMetadata` worksheets of the eDNA study results.
            3. Click `Submit` to validate the table. The Validation Message displays the run status.
            4. The output files from the validator will be displayed under `Download Outputs`. Upon successful validation, the `Metadata` button will appear, allowing the download of the `fairefier_metadata.xlsx` file. If there are any warnings or errors, the `Warnings+Errors` button will appear, enabling the download of the `fairefier_warn_error.xlsx` file. Additionally, corrections needed to the input file are highlighted in red for errors and yellow for warnings in the `_revise` worksheet.
            5. If validation fails, review and correct any errors before resubmitting. Either the corrected input file or the corrected `_revise` worksheet can be used for resubmission.""")),
        id='info', open=False, multiple=True
    ),
    ui.tags.br(),
    
    # Upload and show data
    ui.layout_columns(
        ui.output_ui('ui_file_upload'),
        ui.input_select('dropdown_checklistv', choices=list(dict_checklistv.keys()), selected=list(dict_checklistv.keys())[0], label='Select checklist version', multiple=False),
        col_widths=(8,4),
    ),
    ui.layout_columns(
        ui.input_action_button('button_clear', 'Clear'),
        ui.input_task_button('button_submit', 'Submit'),
    ),
    ui.accordion(
        ui.accordion_panel('Show uploaded worksheets', ui.navset_card_tab(
            ui.nav_panel('projectMetadata', ui.output_data_frame('dataframe_input_projectmetadata')),
            ui.nav_panel('sampleMetadata', ui.output_data_frame('dataframe_input_samplemetadata')),
        )),
        id='show_worksheet', open=False, multiple=True
    ),
    ui.tags.br(),
    
    # Validation message
    ui.card(ui.card_header('Validation Message'), ui.output_text_verbatim('textbox_display', placeholder=True)),
    
    # projectMetadata
    ui.markdown('## Worksheet: projectMetadata'),
    ui.navset_card_tab(
        ui.nav_panel('Output', ui.output_data_frame('dataframe_out_projectmetadata')),
        ui.nav_panel('Warning', ui.output_data_frame('dataframe_warn_projectmetadata')),
        ui.nav_panel('Error', ui.output_data_frame('dataframe_error_projectmetadata')),
    ),
    
    # sampleMetadata
    ui.markdown('## Worksheet: sampleMetadata'),
    ui.navset_card_tab(
        ui.nav_panel('Output', ui.output_data_frame('dataframe_out_samplemetadata')),
        ui.nav_panel('Warning', ui.output_data_frame('dataframe_warn_samplemetadata')),
        ui.nav_panel('Error', ui.output_data_frame('dataframe_error_samplemetadata')),
    ),
    
    # Save button
    ui.markdown('## Download Outputs'),
    ui.output_ui('ui_ddl_out'),
    ui.output_ui('ui_ddl_warn_error'),
    
    # Footer
    ui.card_footer(ui.HTML("""<br><br>
    <hr style='margin: 1em' />
    <p style='text-align:center;font-size:smaller'>
    Developed by Suk Yee Yong 2024
    </p>""")),
)


app = App(app_ui, validate_server)
