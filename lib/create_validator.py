##################################################
# FAIRe-fier CREATE VALIDATORS
# Author: Suk Yee Yong
##################################################

from datetime import datetime
from pydantic import BaseModel, ConfigDict, create_model, Field, field_validator, model_validator, ValidationError, ValidationInfo
from pydantic.fields import FieldInfo
from typing import Literal, Tuple, Union

import difflib
import pandas as pd
import re
import sys
sys.path.append('..')
import warnings
warnings.formatwarning = lambda msg, *args, **kwargs: f'{msg}\n'

import config_data
term_name_col = config_data.term_name_col


def warn_exterm(cls, data):
    """Check for extra terms not defined in the model and raise warning"""
    extra_terms = set(data.keys()) - set(cls.model_fields.keys())
    if extra_terms:
        for exf in extra_terms:
            warnings.warn(f"{exf} | Extra term found that is not defined in the model. Please review to confirm whether it should be included. |  | ")
    return data


def warn_assayname(cls, data):
    """Check all assay_name listed in detected_notDetected_<assay_name> else raise warning"""
    missing_assayname = set(k for k in cls.model_fields.keys() if 'detected_notDetected' in k) - set(k for k in data.keys() if 'detected_notDetected' in k)
    if missing_assayname:
        for ma in missing_assayname:
            warnings.warn(f"{ma} | Term is not defined in the model with the <assay_name> listed. Please review to confirm whether it should be included. |  | ")
    return data


def convert_emptystr2none(cls, v):
    """Convert empty string to None"""
    if (isinstance(v, str) and (v.strip()=='')) or (v is None):
        return None
    return v


def convert_str2numeric(cls, v, info: ValidationInfo):
    """Convert string to numeric for term_type='integer','numeric'"""
    if not v: return v # Empty string or None
    if (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
    if v and isinstance(v, str):
        v = v.replace(' ', '') # Group number without space
        # Contains scientific notation
        if (not v.isalpha() and not v.isdigit()) and any(sn in v for sn in ['x10', 'e']):
            v = re.sub(r'x10', 'e', v)
            v = re.sub(r'(\d)\^', r'\1e', v) # For case with '^', e.g. 10^8
            try:
                return float(v)
            except:
                raise ValueError(f"Require scientific notation to be in format '3x10+8' or '3e8'.")
        # Contains only letters
        assert not v.isalpha(), "Require value to be numerics."
        # Contains symbol
        if not v.isalnum():
            # For value range '-'
            # range_val = re.split(r'(?<=\d)-', v)
            # assert '-' in v and all(val.lstrip('-+').replace('.', '', 1).isdigit() for val in range_val) and len(range_val) == 2, "Require value to be numerics or range MIN-MAX."
            # return '-'.join(range_val)
            match = re.match(r'^(-?\d*\.?\d+)-(-?\d*\.?\d+)$', v)
            error_msg = "Require value to be numerics or range MIN-MAX."
            assert match is not None, error_msg
            start, end = match.groups()
            assert start.lstrip('-+').replace('.', '', 1).isdigit() and end.lstrip('-+').replace('.', '', 1).isdigit(), error_msg
            return f"{start}-{end}"
        # Contains both letters and numbers
        if v.isalnum() and not v.isalpha() and not v.isdigit():
            #v_num = re.findall('[-+]?(?:\d*\.*\d+)', v)[0]
            v_num = re.findall('[-+]?(?:[0-9]*[0-9]+)', v)[0]
            warnings.warn(f"{info.field_name} | Value contains both letters and numbers. Extracting only the numerics. | {v} | {v_num}")
            return v_num
        # Contains number
        else:
            return v
    else:
        return v


def convert_float2int(cls, v):
    """Convert float to int for term_type='boolean'"""
    if isinstance(v, float):
        return int(v)
    if not v: return v # Empty string or None
    elif (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
    assert not isinstance(v, str), "Require value to be boolean."
    return v


def extract_cvocaboptions(df):
    """Extract controlled_vocabulary_options for term_type='controlled vocabulary'"""
    def extract_cvocabopts(cls, v, info: ValidationInfo):
        if not v: return v # Empty string or None
        if (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
        cvocab_options = df[df[term_name_col]==info.field_name]['controlled_vocabulary_options'].item()
        cvocab_optionsi = [cv.strip() for cv in cvocab_options.split('|')]
        selected_options = []
        for vali in v.split('|'):
            vali = vali.strip()
            # other
            if vali.startswith('other'):
                assert vali.partition('other:')[-1].strip(), "other is selected, but detail is not given or not in the right format. Please provide description in the format `other:TEXT`."
                selected_options.append(vali)
            # Exact match
            elif set([vali]).intersection(set(cvocab_optionsi)):
                selected_options.append(vali)
            # Partial match
            else:
                matches = difflib.get_close_matches(vali, cvocab_optionsi, cutoff=0.7)
                assert len(matches) != 0, "Option not found. Please select from the provided options."
                assert len(matches) == 1, f"Multiple matches found {matches}. Please specify the right option."
                warnings.warn(f"{info.field_name} | Closest option found. | {vali} | {matches[0]}")
                selected_options.append(matches[0])
        return ' | '.join(selected_options)
    return extract_cvocabopts


def check_numericcvocab(df):
    """Check numeric or controlled vocabulary"""
    def check_numcvocab(cls, v, info: ValidationInfo):
        if isinstance(v, (float, int)):
            return v
        else:
            try:
                return convert_str2numeric(cls, v, info)
            except:
                return extract_cvocaboptions(df)(cls, v, info)
    return check_numcvocab


def check_decimallatlon(cls, v, info: ValidationInfo):
    """Check number of decimal places for term_name_col='decimalLatitude','decimalLongitude'"""
    if not v: return v # Empty string or None
    if (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
    if str(v)[::-1].find('.') < 5:
        warnings.warn(f"{info.field_name} | Recommended value to have at least 5 decimal places (0.00001). | {v} | ")
    return v


def check_date(cls, v):
    """Check date format"""
    err_msg = "Require date to be in yyyy-mm-dd."
    if not v: return v # Empty string or None
    if isinstance(v, datetime):
        try:
            v = v.strftime('%Y-%m-%d')
        except:
            raise ValueError(err_msg)
    else:
        if (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
        assert datetime.strptime(str(v).strip(), '%Y-%m-%d'), err_msg
    return v


def check_datetime(cls, v, info: ValidationInfo):
    """Check datetime formats"""
    if not v: return v # Empty string or None
    dtformats = [
        '%Y-%m-%dT%H:%M:%SZ', '%Y-%m-%dT%H:%M:%S%z', '%Y-%m-%dT%H:%M:%S',
        '%Y-%m-%dT%H:%MZ', '%Y-%m-%dT%H:%M%z', '%Y-%m-%dT%H:%M',
        '%Y-%m-%d', '%Y-%m', '%Y'
    ]
    v = str(v)
    for dtfmt in dtformats:
        try:
            if '/' in v:
                start, end = v.split('/')
                dt1 = datetime.strptime(start.strip(), dtfmt)
                dt2 = datetime.strptime(end.strip(), dtfmt)
                if dt1 <= dt2:
                    return '/'.join((dt1.strftime(dtfmt), dt2.strftime(dtfmt)))
                else:
                    dt1_dt2 = '/'.join((dt2.strftime(dtfmt), dt1.strftime(dtfmt)))
                    warnings.warn(f"{info.field_name} | Datetime range START>END. Reordered to START<=END. | {v} | {dt1_dt2}")
                    return dt1_dt2
            else:
                if (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
                dt = datetime.strptime(v, dtfmt)
                return dt.strftime(dtfmt)
        except:
            continue
    dtformats_list = dtformats.copy()
    dtformats_list[1] = f"{dtformats_list[1]}/{dtformats_list[1]}"
    raise ValueError(f"Require datetime to have one of these formats {dtformats_list}. Please ensure cells are formatted as `Category: Text` in Excel.")
    return v


def check_termlabelid(cls, v):
    """Check term label and ID format"""
    if not v: return v # Empty string or None
    if (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
    # Search pattern ' [' in middle of text and end with ']'
    pattern = r'^(.+)\s*\[([^\]]+)\]$'
    match = re.match(pattern, v.strip())
    assert bool(match), "Require value to be in format termLabel [termID]."
    return f"{match.group(1).strip()} [{match.group(2)}]"


def check_duration(cls, v):
    """Check ISO 8601 duration format in PnYnMnWnDTnHnMnS"""
    if not v: return v # Empty string or None
    if (v and isinstance(v, str)) and (v.lower().startswith('missing') or v.lower().startswith('not applicable')): return v.lower() # Accept NA
    v = str(v)
    pattern_pt = r'^P(?=\d)(?:(\d+Y)?(\d+M)?(\d+W)?(\d+D)?)(T(?=\d)(\d+H)?(\d+M)?(\d+S)?)?$' # Period and optional time: PnYnMnWnDTnHnMnS
    pattern_t = r'^T(?=\d)(\d+H)?(\d+M)?(\d+S)?$' # Time: TnHnMnS
    pattern = re.compile(r'|'.join((pattern_pt, pattern_t)), re.IGNORECASE)
    match = pattern.match(v)
    assert bool(match), "Require duration to be in format PnYnMnWnDTnHnMnS."
    return v.upper()


def check_assayname(cls, v):
    """Check assay_name in sampleMetadata is a subset of those in projectMetadata"""
    sm_assayname = [val.strip() for val in v.split('|')]
    if len(sm_assayname) > 1:
        pm_assayname = [k.rsplit('_',1)[-1] for k in cls.model_fields.keys() if 'detected_notDetected_' in k]
        diff_assayname = set(sm_assayname) - set(pm_assayname)
        assert not diff_assayname, f"{list(diff_assayname)} is specified but not listed in projectMetadata assay_name."
        return ' | '.join(sm_assayname)
    return v


def create_dictterms(df: pd.DataFrame) -> dict[Tuple[type, Field]]:
    """Create dict of {<term_name_col>: (<type>, Field(...))} for dynamic model creation"""
    dict_terms = {}
    for dt in df.to_records(index=False):
        # Term type option
        if dt['term_type'].lower() == 'integer':
            termtype_opt = (int, str)
        elif dt['term_type'].lower() == 'numeric':
            termtype_opt = (float, int, str)
        elif dt['term_type'].lower() == 'boolean':
            termtype_opt = (Literal[0, 1, '0', '1'], str)
        elif dt['term_type'].lower() == 'numeric or controlled vocabulary':
            termtype_opt = (str, float, int)
        else:
            termtype_opt = (str, float, int, datetime)
        termtype_union = Union[termtype_opt] if len(termtype_opt) > 1 else termtype_opt[0]
        # Requirement level option
        if dt['requirement_level'].lower() == 'mandatory':
            dict_terms[dt[term_name_col]] = (termtype_union, Field(description=dt['description']))
        else:
            dict_terms[dt[term_name_col]] = (Union[None, termtype_union], Field(None, description=dt['description']))
    return dict_terms


def validate_projectMetadata(df: pd.DataFrame, dict_terms: dict[Tuple[type, Field]], inputs: dict, mandatory_to_optional: list=None) -> BaseModel:
    """Validate projectMetadata model depending on the inputs"""
    validators = {
        'warn_exterm': model_validator(mode='before')(warn_exterm),
        'convert_emptystr2none': field_validator('*', mode='before')(convert_emptystr2none),
        'convert_str2numeric': field_validator(*df.loc[df['term_type'].isin(['integer', 'numeric'])][term_name_col].values, mode='before')(convert_str2numeric),
        'extract_cvocaboptions': field_validator(*df.loc[df['term_type']=='controlled vocabulary'][term_name_col].values, mode='before')(extract_cvocaboptions(df)),
        'check_datetime': field_validator('mod_date', mode='before')(check_datetime),
        'check_numericcvocab': field_validator('min_reads_cutoff', mode='before')(check_numericcvocab(df)),
    }
    cls = create_model('eDNAMetaTable', **dict_terms, __validators__=validators, __config__=ConfigDict(extra='allow'))
    # cls.model_fields
    
    # Change mandatory to optional
    if mandatory_to_optional is not None:
        for term in mandatory_to_optional:
            cls.model_fields[term].default = None
            cls.model_fields[term].annotation = cls.model_fields[term].annotation | None
    
    # Change mandatory to optional from requirement_level_condition
    if str(inputs.get('pcr_0_1', None)) != '1':
        for term in ['targetTaxonomicAssay', 'pcr_primer_forward', 'pcr_primer_reverse']:
            if cls.model_fields.get(term, None):
                cls.model_fields[term].default = None
                cls.model_fields[term].annotation = cls.model_fields[term].annotation | None
    if not ((inputs.get('assay_type', None) == 'targeted') and (str(inputs.get('pcr_0_1', None)) == '1')):
        for term in ['amp_vis_method', 'detection_criteria', 'pcr_assay_lod', 'pcr_assay_lod_unit', 'pcr_assay_loq', 'pcr_assay_loq_unit']:
            if cls.model_fields.get(term, None):
                cls.model_fields[term].default = None
                cls.model_fields[term].annotation = cls.model_fields[term].annotation | None
    if not ((inputs.get('assay_type', None) == 'targeted') and (inputs.get('amp_vis_method', None) == 'qPCR')):
        for term in ['thresholdQuantificationCycle']:
            if cls.model_fields.get(term, None):
                cls.model_fields[term].default = None
                cls.model_fields[term].annotation = cls.model_fields[term].annotation | None
    cls.model_rebuild(force=True)
    return cls.model_validate(inputs).model_dump()


def validate_sampleMetadata(df: pd.DataFrame, dict_terms: dict[Tuple[type, Field]], inputs: dict, inputs_pm: dict, informationwithheld_latlon: bool=False) -> BaseModel:
    """Validate sampleMetadata model depending on the inputs"""
    # List terms for term_type=='boolean'
    termtype_bool = df.loc[df['term_type']=='Boolean'][term_name_col].tolist()
    # Add detected_notDetected_<assay_name> term_name for multiple assays from assay_name listed in projectMetadata if assay_type=='targeted'
    list_assay_name = [an.strip() for an in inputs_pm['project_level'].get('assay_name', '').split('|')]
    if (inputs_pm['project_level'].get('assay_type') == 'targeted') and len(list_assay_name) > 1:
        dict_dnd = {f"detected_notDetected_{an}": tuple(dict_terms['detected_notDetected']) for an in list_assay_name}
        # Insert new dict at specific position
        dict_terms = list(dict_terms.items())
        idx = [i for i, (k, _) in enumerate(dict_terms) if k == 'detected_notDetected'][0]
        dict_terms[idx + 1:idx + 1] = list(dict_dnd.items())
        dict_terms = dict(dict_terms)
        dict_terms.pop('detected_notDetected', None)
        # Modify term_type=='boolean' list with assay_name
        termtype_bool.extend([f"detected_notDetected_{an}" for an in list_assay_name])
        termtype_bool.remove('detected_notDetected')
    validators = ({
        'warn_assayname': model_validator(mode='before')(warn_assayname),
        'check_assayname': field_validator('assay_name', mode='before')(check_assayname),
    } if (inputs_pm['project_level'].get('assay_type') == 'targeted') else {}) \
    | {
        'warn_exterm': model_validator(mode='before')(warn_exterm),
        'convert_emptystr2none': field_validator('*', mode='before')(convert_emptystr2none),
        'convert_str2numeric': field_validator(*df.loc[df['term_type'].isin(['integer', 'numeric'])][term_name_col].values, mode='before')(convert_str2numeric),
        'convert_float2int': field_validator(*termtype_bool, mode='before')(convert_float2int),
        'extract_cvocaboptions': field_validator(*df.loc[df['term_type']=='controlled vocabulary'][term_name_col].values, mode='before')(extract_cvocaboptions(df)),
        'check_decimallatlon': field_validator('decimalLatitude', 'decimalLongitude', mode='before')(check_decimallatlon),
        'check_datetime': field_validator('eventDate', 'date_ext', mode='before')(check_datetime),
        'check_termlabelid': field_validator(*df[term_name_col].loc[df[term_name_col].str.startswith('env_')].values, 'host_life_stage', mode='before')(check_termlabelid),
        'check_numericcvocab': field_validator('samp_store_temp', 'prepped_samp_store_temp', mode='before')(check_numericcvocab(df)),
        'check_duration': field_validator(*df.loc[df['fixed_format']=='PnYnMnWnDTnHnMnS'][term_name_col].values, mode='before')(check_duration),
    }
    cls = create_model('eDNAMetaTable', **dict_terms, __validators__=validators, __config__=ConfigDict(extra='allow'))
    # cls.model_fields
    
    # Change mandatory to optional from requirement_level_condition
    if str(inputs.get('samp_category', None)) != 'negative control':
        cls.model_fields['neg_cont_type'].default = None
        cls.model_fields['neg_cont_type'].annotation = cls.model_fields['neg_cont_type'].annotation | None
    if str(inputs.get('samp_category', None)) != 'positive control':
        cls.model_fields['pos_cont_type'].default = None
        cls.model_fields['pos_cont_type'].annotation = cls.model_fields['pos_cont_type'].annotation | None
    if informationwithheld_latlon or (str(inputs.get('samp_category', None)) in ['negative control', 'positive control', 'PCR standard']):
        for term in ['decimalLatitude', 'decimalLongitude']:
            if cls.model_fields.get(term, None):
                cls.model_fields[term].default = None
                cls.model_fields[term].annotation = Union[None, float, int, str]
    if str(inputs.get('samp_category', None)) in ['negative control', 'positive control', 'PCR standard']:
        for term in ['env_broad_scale', 'env_local_scale', 'env_medium']:
            if cls.model_fields.get(term, None):
                cls.model_fields[term].default = None
                cls.model_fields[term].annotation = cls.model_fields[term].annotation | None
    if inputs_pm['project_level'].get('assay_type') != 'targeted':
        for term in [k for k in cls.model_fields.keys() if 'detected_notDetected' in k]:
            cls.model_fields[term].default = None
            cls.model_fields[term].annotation = cls.model_fields[term].annotation | None
    cls.model_fields['eventDate'].annotation = Union[str, int]
    cls.model_rebuild(force=True)
    return cls.model_validate(inputs).model_dump()
