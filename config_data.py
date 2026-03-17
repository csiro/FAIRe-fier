##################################################
# FAIRe-fier CONFIGS
# Author: Suk Yee Yong
##################################################

from pathlib import Path
import glob

data_dir = Path.cwd()
list_checklistfile = glob.glob1(data_dir, 'FAIRe_checklist*.xlsx')
dict_checklistv = {Path(v).stem.rsplit('_', 1)[-1]: v for v in list_checklistfile}

term_name_col = 'term_name'
samp_name_col = 'samp_name'
