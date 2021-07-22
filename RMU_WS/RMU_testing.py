import pickle
import time
from tkinter.constants import ACTIVE, DISABLED
from urllib.parse import DefragResult
from requests.api import options
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from PIL import Image, ImageDraw
from datetime import datetime,timedelta
from openpyxl import Workbook,load_workbook
import glob
import os
import random
import tkinter as tk
from datetime import datetime
import requests
import threading
import json

# Variables Decelaration
chrome_path = r'D:\Process Improvement Project\python_programming\chromedriver.exe'
base_url = '125.22.84.50:7860'
url_ioms = 'http://{}/ioms'.format(base_url)
url_ioms_bd = 'http://{}/IOMS/DOMS/breakdownhome'.format(base_url)
ioms_username = '41018328'
ioms_password = 'india@123'
dr = ''
breakdown_id = ''
ids = ['divisionname', 'gridname', 'feedername', 'subtraname','relaytrip','reason','load',
            'Expectedhours', 'Expectedminutes', 'infoname', 'infoempno', 'infodesig', 'infocontact']
psdid = 'P31052100004'

def init_driver():
    global dr
    dr = webdriver.Chrome(executable_path=chrome_path)
    dr.maximize_window()
    login_ioms()

def login_ioms():
    dr.get(url_ioms)
    dr.find_element_by_name('UserID').send_keys(ioms_username)
    dr.find_element_by_name('Password').send_keys(ioms_password)
    dr.find_element_by_xpath('//*[@id="login"]/form/table/tbody/tr[3]/td[3]/button').click()

def generate_bd():
    da = {'load': '20', 'Expectedhours': '2', 'Expectedminutes': '0', 'infoname': 'Ashwani', 'infoempno': '41018328','infodesig': 'MGR', 
    'infocontact': '7489447555', 'relaytrip': '1', 'reason': '2'}
    dr.get(url_ioms_bd)
    check_title()
    p_id = ''
    for id in ids:
        elm = dr.find_element_by_id(id)
        typ = elm.get_attribute('outerHTML').split(' ')[0].split('<')[1]
        if p_id == '':
            p_id = id
        if  typ == 'select':
            s = Select(elm)
            opts = s.options
            if id not in da.keys():
                    da[id] = random.randint(1,len(opts)-1)
            while True:
                opts = s.options
                if len(opts) > 1 and not bool(s.is_multiple):
                    break
                elif len(opts) > 0 and bool(s.is_multiple):
                    break
                else:
                    reset_select(p_id,da)
            if s.is_multiple:
                print(id)
                s.deselect_all()
                time.sleep(2)
                opts[da[id]].click()
            else:
                s.select_by_index(da[id])
            imgloader()
            p_id = id
        elif typ == 'input':
            if id == 'load':
                elm.send_keys('\b')
            elm.send_keys(da[id])

def check_title():
    while True:
        if '500' not in dr.title:
            break
        else:
            dr.refresh()

def reset_select(p_id,da):
    s_el =Select(dr.find_element_by_id(p_id))
    s_el.select_by_index('0')
    imgloader()
    s_el.select_by_value(da[p_id])
    imgloader()

def imgloader():
    imgloader = dr.find_element_by_id('imageLoader')
    while True:
        if 'none' in imgloader.get_attribute('style'):
            break

def read_bdid():
    global breakdown_id
    txt_ = dr.find_element_by_xpath('//*[@id="fademsgboxsmall"]/div/ul/li').text
    breakdown_id = txt_.split('=')[1].split(' ')[0]
    # lbl_bdid.config(text=breakdown_id)

def restore_bd():
    url_restore = 'http://10.125.64.87/IOMS/DOMS/Updatefault/{}'.format(breakdown_id)
    dr.get(url_restore)
    check_title()
    dr.find_element_by_id('Backfeed').click()
    time.sleep(2)
    Select(dr.find_element_by_id('backfeedsource')).select_by_index(1)
    dr.find_element_by_name('remarks1').send_keys('TEST RUN')
    dr.find_element_by_id('chkexib').click()
    dr.find_element_by_xpath('//*[@id="backfeed_box"]/table[3]/tbody/tr/td/button[1]').click()

def force_majeure():
    dr.find_element_by_id('ForceMajeure').click()
    Select(dr.find_element_by_id('element')).select_by_index(3)
    Select(dr.find_element_by_id('subelement')).select_by_index(3)
    dr.find_element_by_id('saveforcemajeure').click()

def close_bd():
    url_restore = 'http://10.125.64.87/IOMS/DOMS/Updatefault/{}'.format(breakdown_id)
    dr.get(url_restore)
    check_title()
    dr.find_element_by_id('Close').click()
    Select(dr.find_element_by_id('closingtype')).select_by_index(1)
    remarks = ['actualfinding','remarks1','fltlocation']
    for remark in remarks:
        dr.find_element_by_id(remark).send_keys('TEST RUN')
    affected_elements = ['DISTRIBUTION TRANSFORMER','RMU']
    Select(dr.find_element_by_id('closingelement')).select_by_value(affected_elements[1])
    time.sleep(1)
    Select(dr.find_element_by_id('subclosingproblem')).select_by_value('MECHANISM STRUCK/FAULTY')
    time.sleep(1)
    subdivisions_ids =['subdivisionid','clssubstation','clsdt']
    for sub_id in subdivisions_ids:
        sel_ele = Select(dr.find_element_by_id(sub_id))
        index = random.randint(1,len(sel_ele.options)-1)
        sel_ele.select_by_index(index)
        time.sleep(1)
    dr.find_element_by_xpath('//*[@id="close_box"]/table/tbody/tr[13]/td/button[1]').click()

def fill_FIR():
    select_elements = [
        "Configuration","SUBDIVISION","Month","YearOfManufacturing","RMUType","SwitchingPosition","TypeOfFault","SF6","sdo_Circle"
    ]
    for select_element in select_elements:
        element = Select(dr.find_element_by_id(select_element))
        selected_option = element.all_selected_options[0]
        if selected_option.get_attribute('title') == 'Select':
            nos = len(element.options)-1
            index = random.randint(1,nos)
            element.select_by_index(index)
    input_elements = {"SrNo":'235689',
    "DateOfInspection":(datetime.now()-timedelta(1)).strftime('%d-%m-%Y %H:%M:%S'),
    "sdo_CostCenter":'25634',
    "underwarranty_rmks":"SDO REMARK - TEST",
    "Reason_of_failure" : "Reason of Failure - TEST",
    "Name_Of_Incharge":"Tester",
    "PR_No":'41012563',
    "EmailId_of_Incharge":"test@relianceada.com",
    "Mobile_no":'7458963210'}
    for key,val in input_elements.items():
        script = '$("#{}").val("{}")'.format(key,val)
        dr.execute_script(script)

def upload_file_in_FIR():
    file_upload = dr.find_element_by_id('fileToUpload_report')
    file_list = glob.glob('files\*.*')
    file_paths = []
    for file in file_list:
        file_paths.append(os.path.abspath(file))
    n = random.randint(0,len(file_list)-1)
    file_upload.send_keys(file_list[n])
    
def paste_psd_id():
    global psdid
    dr.find_element_by_id('psd_id').send_keys(psdid)

def fill_non_repairable():
    input_elements_ok = [
    "CompletenessOfAssembly",
    "OperationalTest",
    "InterlockFunctionalTest",
    "PowerFrequency",
    "ClsdTestPE_Rmks",
    "clsdTestPh_PhRemarks",
    "CB_ACROSS_ISOLATOR_Remarks",
    "CB_ISOLATOR_Remarks",
    "TEST_ISOLATOR_2_Remarks",
    "TEST_ISOLATOR_1_Remarks",
    "SF6GASPRESSURE",
    "VERIFICATION_WIRING_CIRCUIT",
    "RELAY_FUNCTION_TESTING",
    "CB_TRIP_TEST_RELAY",
    "CT_RATIO_TEST",
    "VOLTAGE_INDICATION_TEST",
    "BUSHING_CONDITION",
    "FPI_EFI_TEST",
    "INSP_REP_REMARKS"
    ]
    input_elements_radio = [
    "PH_E_28KVAC_12KVDC",
    "PH_E_28KVAC_12KVDC",
    "PH_PH_E_28KVAC_12KVDC",
    "PH_PH_E_28KVAC_12KVDC",
    "accross_isolator32_12",
    "accross_isolator32_12",    
    "between_isolator32_12",
    "between_isolator32_12",
    ]
    input_elements_megger = [
    "ISOLATOR_1_TO_CB_R",
    "ISOLATOR_2_TO_CB_R",
    "ISOLATOR_1_TO_ISOLATOR_2_R",
    "ISOLATOR_1_TO_EARTH_R",
    "ISOLATOR_2_TO_EARTH_R",
    "CB_TO_EARTH_R",
    "ISOLATOR_TO_EXT_BUSHING_R",
    "ISOLATOR_1_TO_CB_Y",
    "ISOLATOR_2_TO_CB_Y",
    "ISOLATOR_1_TO_ISOLATOR_2_Y",
    "ISOLATOR_1_TO_EARTH_Y",
    "ISOLATOR_2_TO_EARTH_Y",
    "CB_TO_EARTH_Y",
    "ISOLATOR_TO_EXT_BUSHING_Y",
    "ISOLATOR_1_TO_CB_B",
    "ISOLATOR_2_TO_CB_B",
    "ISOLATOR_1_TO_ISOLATOR_2_B",
    "ISOLATOR_1_TO_EARTH_B",
    "ISOLATOR_2_TO_EARTH_B",
    "CB_TO_EARTH_B",
    "ISOLATOR_TO_EXT_BUSHING_B",
    ]
    input_elements_workshop ={
    "NAME" : 'Test Engineer',
    "JOBCARD_DESIGNATION": 'ENGR',
    "JOBCARD_NO" : random.randint(100,999),
    "MFG": 'BSES WS',
    }
    for inp in input_elements_ok:
        script = script = '$("#{}").val("TESTING")'.format(inp)
        dr.execute_script(script)
    for inp in input_elements_megger:
        val = random.randint(1000,9999)
        script = script = '$("#{}").val("{}")'.format(inp,val)
        dr.execute_script(script)
    for key,val in input_elements_workshop.items():
        script = '$("#{}").val("{}")'.format(key,val)
        dr.execute_script(script)

def non_repairable_sdo():
    input_elements_sdo ={
    "SITE_ENGG_SDO_NAME" : 'Test Engineer',
    "SITE_ENGG_SDO_DESIGNATION" : 'MGR',
    "SITE_ENGG_SDO_PR_NO" : '41012563',
    "SITE_ENGG_SDO":'SDO',
    }
    for key,val in input_elements_sdo.items():
        script = '$("#{}").val("{}")'.format(key,val)
        dr.execute_script(script)

def fill_justification():
    element_list = {
        "loc_ht_panel":'xyz location',
        "Dt_rating" : '400kVA',
        "panel_age" : '10 year',
        'panel_type' : 'Indoor',
        'panel_make' : 'ABB',
        'proposed_config' : '3-way',
        'sdo_boq_remarks' : 'Testing-Remark'
    }
    for key,val in element_list.items():
        script = '$("#{}").val("{}")'.format(key,val)
        dr.execute_script(script)

def hoto_sdo():
    select_elements = [
        "ACCESSORIES_SDO_SF6GASMANOMETER1",
        "ACCESSORIES_REMARKS_SDO_SF6GASMANOMETER1",
        "ACCESSORIES_STORE_SF6GASMANOMETER1",
        "ACCESSORIES_REMARKS_STORE_SF6GASMANOMETER1",
        "ACCESSORIES_SDO_RELAY2",
        "ACCESSORIES_REMARKS_SDO_RELAY2",
        "ACCESSORIES_STORE_RELAY2",
        "ACCESSORIES_REMARKS_STORE_RELAY2",
        "ACCESSORIES_SDO_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_REMARKS_SDO_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_STORE_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_REMARKS_STORE_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_SDO_EFPIORFPI4",
        "ACCESSORIES_REMARKS_SDO_EFPIORFPI4",
        "ACCESSORIES_STORE_EFPIORFPI4",
        "ACCESSORIES_REMARKS_STORE_EFPIORFPI4",
        "ACCESSORIES_SDO_CT$WIREOFEFPIORFPI5",
        "ACCESSORIES_REMARKS_SDO_CT$WIREOFEFPIORFPI5","ACCESSORIES_STORE_CT$WIREOFEFPIORFPI5","ACCESSORIES_REMARKS_STORE_CT$WIREOFEFPIORFPI5","ACCESSORIES_SDO_CTSOFCB-1$RATING6","ACCESSORIES_REMARKS_SDO_CTSOFCB-1$RATING6","ACCESSORIES_STORE_CTSOFCB-1$RATING6","ACCESSORIES_REMARKS_STORE_CTSOFCB-1$RATING6","ACCESSORIES_SDO_CTSOFCB-2$RATING7","ACCESSORIES_REMARKS_SDO_CTSOFCB-2$RATING7","ACCESSORIES_STORE_CTSOFCB-2$RATING7","ACCESSORIES_REMARKS_STORE_CTSOFCB-2$RATING7","ACCESSORIES_SDO_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_REMARKS_SDO_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_STORE_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_REMARKS_STORE_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_SDO_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_REMARKS_SDO_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_STORE_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_REMARKS_STORE_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_SDO_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_REMARKS_SDO_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_STORE_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_REMARKS_STORE_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_SDO_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_REMARKS_SDO_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_STORE_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_REMARKS_STORE_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_SDO_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_REMARKS_SDO_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_STORE_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_REMARKS_STORE_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_SDO_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_REMARKS_SDO_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_STORE_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_REMARKS_STORE_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_SDO_INTERNALLOCKCONDITION14","ACCESSORIES_REMARKS_SDO_INTERNALLOCKCONDITION14","ACCESSORIES_STORE_INTERNALLOCKCONDITION14","ACCESSORIES_REMARKS_STORE_INTERNALLOCKCONDITION14","ACCESSORIES_SDO_CB-RMECHANISMCONDITION15","ACCESSORIES_REMARKS_SDO_CB-RMECHANISMCONDITION15","ACCESSORIES_STORE_CB-RMECHANISMCONDITION15","ACCESSORIES_REMARKS_STORE_CB-RMECHANISMCONDITION15","ACCESSORIES_SDO_CB-LMECHANISMCONDITION16","ACCESSORIES_REMARKS_SDO_CB-LMECHANISMCONDITION16","ACCESSORIES_STORE_CB-LMECHANISMCONDITION16","ACCESSORIES_REMARKS_STORE_CB-LMECHANISMCONDITION16","ACCESSORIES_SDO_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_REMARKS_SDO_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_STORE_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_REMARKS_STORE_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_SDO_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_REMARKS_SDO_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_STORE_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_REMARKS_STORE_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_SDO_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_REMARKS_SDO_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_STORE_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_REMARKS_STORE_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_SDO_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_REMARKS_SDO_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_STORE_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_REMARKS_STORE_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_SDO_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_REMARKS_SDO_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_STORE_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_REMARKS_STORE_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_SDO_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_REMARKS_SDO_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_STORE_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_REMARKS_STORE_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_SDO_MARSHALBOXCOVER23","ACCESSORIES_REMARKS_SDO_MARSHALBOXCOVER23","ACCESSORIES_STORE_MARSHALBOXCOVER23","ACCESSORIES_REMARKS_STORE_MARSHALBOXCOVER23","ACCESSORIES_SDO_CLEAT$CLAMPCHANNEL24","ACCESSORIES_REMARKS_SDO_CLEAT$CLAMPCHANNEL24","ACCESSORIES_STORE_CLEAT$CLAMPCHANNEL24","ACCESSORIES_REMARKS_STORE_CLEAT$CLAMPCHANNEL24","ACCESSORIES_SDO_TERMINATIONCOVEROFODRMU25","ACCESSORIES_REMARKS_SDO_TERMINATIONCOVEROFODRMU25","ACCESSORIES_STORE_TERMINATIONCOVEROFODRMU25","ACCESSORIES_REMARKS_STORE_TERMINATIONCOVEROFODRMU25","ACCESSORIES_SDO_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_REMARKS_SDO_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_STORE_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_REMARKS_STORE_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_SDO_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_REMARKS_SDO_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_STORE_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_REMARKS_STORE_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_SDO_SIDECOVEROFRMU28","ACCESSORIES_REMARKS_SDO_SIDECOVEROFRMU28","ACCESSORIES_STORE_SIDECOVEROFRMU28","ACCESSORIES_REMARKS_STORE_SIDECOVEROFRMU28","ACCESSORIES_SDO_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_REMARKS_SDO_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_STORE_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_REMARKS_STORE_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_SDO_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_REMARKS_SDO_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_STORE_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_REMARKS_STORE_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_SDO_EARTHSTRIPOFRMU31","ACCESSORIES_REMARKS_SDO_EARTHSTRIPOFRMU31","ACCESSORIES_STORE_EARTHSTRIPOFRMU31","ACCESSORIES_REMARKS_STORE_EARTHSTRIPOFRMU31","ACCESSORIES_SDO_COLLAROFBUSHING32","ACCESSORIES_REMARKS_SDO_COLLAROFBUSHING32","ACCESSORIES_STORE_COLLAROFBUSHING32","ACCESSORIES_REMARKS_STORE_COLLAROFBUSHING32","ACCESSORIES_SDO_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_REMARKS_SDO_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_STORE_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_REMARKS_STORE_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_SDO_BOOTOFCABLETERMINATION34","ACCESSORIES_REMARKS_SDO_BOOTOFCABLETERMINATION34","ACCESSORIES_STORE_BOOTOFCABLETERMINATION34","ACCESSORIES_REMARKS_STORE_BOOTOFCABLETERMINATION34"
    ]
    for select_element in select_elements:
        element = Select(dr.find_element_by_id(select_element))
        if 'SDO' in select_element:
            nos = len(element.options)-1
            index = random.randint(1,nos)
            element.select_by_index(index)
    dr.find_element_by_id('sdo_rmks').send_keys('TESTING - REMARKS')

def hoto_workshop():
    select_elements = [
        "ACCESSORIES_SDO_SF6GASMANOMETER1",
        "ACCESSORIES_REMARKS_SDO_SF6GASMANOMETER1",
        "ACCESSORIES_STORE_SF6GASMANOMETER1",
        "ACCESSORIES_REMARKS_STORE_SF6GASMANOMETER1",
        "ACCESSORIES_SDO_RELAY2",
        "ACCESSORIES_REMARKS_SDO_RELAY2",
        "ACCESSORIES_STORE_RELAY2",
        "ACCESSORIES_REMARKS_STORE_RELAY2",
        "ACCESSORIES_SDO_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_REMARKS_SDO_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_STORE_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_REMARKS_STORE_VPIVOLTAGEPRESENCEINDICATOR3",
        "ACCESSORIES_SDO_EFPIORFPI4",
        "ACCESSORIES_REMARKS_SDO_EFPIORFPI4",
        "ACCESSORIES_STORE_EFPIORFPI4",
        "ACCESSORIES_REMARKS_STORE_EFPIORFPI4",
        "ACCESSORIES_SDO_CT$WIREOFEFPIORFPI5",
        "ACCESSORIES_REMARKS_SDO_CT$WIREOFEFPIORFPI5","ACCESSORIES_STORE_CT$WIREOFEFPIORFPI5","ACCESSORIES_REMARKS_STORE_CT$WIREOFEFPIORFPI5","ACCESSORIES_SDO_CTSOFCB-1$RATING6","ACCESSORIES_REMARKS_SDO_CTSOFCB-1$RATING6","ACCESSORIES_STORE_CTSOFCB-1$RATING6","ACCESSORIES_REMARKS_STORE_CTSOFCB-1$RATING6","ACCESSORIES_SDO_CTSOFCB-2$RATING7","ACCESSORIES_REMARKS_SDO_CTSOFCB-2$RATING7","ACCESSORIES_STORE_CTSOFCB-2$RATING7","ACCESSORIES_REMARKS_STORE_CTSOFCB-2$RATING7","ACCESSORIES_SDO_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_REMARKS_SDO_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_STORE_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_REMARKS_STORE_CLEAT$CLAMPFOR3WAY-3nos4WAY-4nos8","ACCESSORIES_SDO_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_REMARKS_SDO_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_STORE_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_REMARKS_STORE_BASEPLATEORCABLEENTRYPLATEINIDRMU9","ACCESSORIES_SDO_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_REMARKS_SDO_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_STORE_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_REMARKS_STORE_HV_BUSHINGPORCELAINEPOXYINSULATORS10","ACCESSORIES_SDO_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_REMARKS_SDO_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_STORE_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_REMARKS_STORE_HV_BUSHINGMETALPARTSBRASS11","ACCESSORIES_SDO_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_REMARKS_SDO_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_STORE_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_REMARKS_STORE_EXTENSIBLE_BUSHINGPORCELAINEPOXYINSULATORS12","ACCESSORIES_SDO_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_REMARKS_SDO_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_STORE_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_REMARKS_STORE_EXTENSIBLE_BUSHINGMETALPARTSBRASS13","ACCESSORIES_SDO_INTERNALLOCKCONDITION14","ACCESSORIES_REMARKS_SDO_INTERNALLOCKCONDITION14","ACCESSORIES_STORE_INTERNALLOCKCONDITION14","ACCESSORIES_REMARKS_STORE_INTERNALLOCKCONDITION14","ACCESSORIES_SDO_CB-RMECHANISMCONDITION15","ACCESSORIES_REMARKS_SDO_CB-RMECHANISMCONDITION15","ACCESSORIES_STORE_CB-RMECHANISMCONDITION15","ACCESSORIES_REMARKS_STORE_CB-RMECHANISMCONDITION15","ACCESSORIES_SDO_CB-LMECHANISMCONDITION16","ACCESSORIES_REMARKS_SDO_CB-LMECHANISMCONDITION16","ACCESSORIES_STORE_CB-LMECHANISMCONDITION16","ACCESSORIES_REMARKS_STORE_CB-LMECHANISMCONDITION16","ACCESSORIES_SDO_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_REMARKS_SDO_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_STORE_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_REMARKS_STORE_ISOLATOR-LMECHANISMCONDITION17","ACCESSORIES_SDO_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_REMARKS_SDO_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_STORE_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_REMARKS_STORE_ISOLATOR-LMECHANISMCONDITION18","ACCESSORIES_SDO_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_REMARKS_SDO_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_STORE_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_REMARKS_STORE_EXTENSIBLEBUSHINGTERMINATIONCOVER19","ACCESSORIES_SDO_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_REMARKS_SDO_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_STORE_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_REMARKS_STORE_LOCKINGARRANGEMENTOFDOOR20","ACCESSORIES_SDO_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_REMARKS_SDO_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_STORE_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_REMARKS_STORE_NUTBOLTOFTERMINATIONCOVER21","ACCESSORIES_SDO_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_REMARKS_SDO_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_STORE_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_REMARKS_STORE_MARSHALBOXORRELAYWIRINGCONNECTOR22","ACCESSORIES_SDO_MARSHALBOXCOVER23","ACCESSORIES_REMARKS_SDO_MARSHALBOXCOVER23","ACCESSORIES_STORE_MARSHALBOXCOVER23","ACCESSORIES_REMARKS_STORE_MARSHALBOXCOVER23","ACCESSORIES_SDO_CLEAT$CLAMPCHANNEL24","ACCESSORIES_REMARKS_SDO_CLEAT$CLAMPCHANNEL24","ACCESSORIES_STORE_CLEAT$CLAMPCHANNEL24","ACCESSORIES_REMARKS_STORE_CLEAT$CLAMPCHANNEL24","ACCESSORIES_SDO_TERMINATIONCOVEROFODRMU25","ACCESSORIES_REMARKS_SDO_TERMINATIONCOVEROFODRMU25","ACCESSORIES_STORE_TERMINATIONCOVEROFODRMU25","ACCESSORIES_REMARKS_STORE_TERMINATIONCOVEROFODRMU25","ACCESSORIES_SDO_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_REMARKS_SDO_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_STORE_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_REMARKS_STORE_FRONT_COVEROFIDRMUTOP26","ACCESSORIES_SDO_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_REMARKS_SDO_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_STORE_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_REMARKS_STORE_FRONT_COVEROFIDRMUBOTTOM27","ACCESSORIES_SDO_SIDECOVEROFRMU28","ACCESSORIES_REMARKS_SDO_SIDECOVEROFRMU28","ACCESSORIES_STORE_SIDECOVEROFRMU28","ACCESSORIES_REMARKS_STORE_SIDECOVEROFRMU28","ACCESSORIES_SDO_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_REMARKS_SDO_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_STORE_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_REMARKS_STORE_TESTINGPOINTCOVEROFRMU29","ACCESSORIES_SDO_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_REMARKS_SDO_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_STORE_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_REMARKS_STORE_TESTPOINTCOPPERSHORTSTRIPINODRMU30","ACCESSORIES_SDO_EARTHSTRIPOFRMU31","ACCESSORIES_REMARKS_SDO_EARTHSTRIPOFRMU31","ACCESSORIES_STORE_EARTHSTRIPOFRMU31","ACCESSORIES_REMARKS_STORE_EARTHSTRIPOFRMU31","ACCESSORIES_SDO_COLLAROFBUSHING32","ACCESSORIES_REMARKS_SDO_COLLAROFBUSHING32","ACCESSORIES_STORE_COLLAROFBUSHING32","ACCESSORIES_REMARKS_STORE_COLLAROFBUSHING32","ACCESSORIES_SDO_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_REMARKS_SDO_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_STORE_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_REMARKS_STORE_GENERALCONDITIONOFRMUBODY33","ACCESSORIES_SDO_BOOTOFCABLETERMINATION34","ACCESSORIES_REMARKS_SDO_BOOTOFCABLETERMINATION34","ACCESSORIES_STORE_BOOTOFCABLETERMINATION34","ACCESSORIES_REMARKS_STORE_BOOTOFCABLETERMINATION34"
    ]
    for select_element in select_elements:
        element = Select(dr.find_element_by_id(select_element))
        if 'STORE' in select_element:
            nos = len(element.options)-1
            index = random.randint(1,nos)
            element.select_by_index(index)
    dr.find_element_by_id('Workshop_rmks').send_keys('TESTING - REMARKS')

def boq_1():
    elms = ['proposal_no','proposal_Des']
    dr.find_element_by_id(elms[0]).send_keys('900000564123')
    dr.find_element_by_id(elms[1]).send_keys('Replacement of RMU at XYZ Location')
    
def finance_store():
    dr.find_element_by_id('PM11OrderNo').send_keys('900000564123')

def hoto_sdo_1():
    dr.find_element_by_id('CurrentRating').send_keys('200 A')
    dr.find_element_by_id('MakeOfRelay').send_keys('ABB')
    dr.find_element_by_id('Actual_Equipid').send_keys('DL-11SWGXYZ3000')

# 23-06-2021 10:50:55
# pathscreenshot = r'D:\Process Improvement Project\python_programming\RMU_WS\Screenshot\{}'
# dr.get_screenshot_as_file(pathscreenshot.format('screen1.png'))
# for n in [5,6,8,9]:
#     breakdown_id='B2306210000{}'.format(n)
#     try:
#         close_bd()
#     except:
#         pass
# select_elements = dr.find_elements_by_tag_name('Select')
# for select in select_elements:
#     print('"{}"'.format(select.get_attribute('id')),end=',')

# fade msgbox id = fademsgboxsmall
# bd_col_xpath = '//*[@id="gridContent"]/table/tbody/tr[{}]/td[1]/a'
# bd_ids = []
# col_count = len(dr.find_elements_by_xpath('//*[@id="gridContent"]/table/tbody/tr'))
# for n in range(1,col_count+1):
#     txt = dr.find_element_by_xpath(bd_col_xpath.format(n)).text
#     if txt[0] == 'B':
#         bd_ids.append(txt)
#         print(txt)

# url = 'http://10.125.64.87/IOMS/DOMS/Updatefault/{}'
# data = []
# for bd in bd_ids:
#     dr.get(url.format(bd))
#     d = {}
#     for inp in inp_id+sel_id:
#         try:
#             ele = dr.find_element_by_id(inp)
#             val = ele.get_attribute('value')
#             d[inp] = val
#         except:
#             d[inp] = ''
#             pass
#         data.append(d)

# 500 - Internal server error.

# 'divisionname': 'DWK', 'gridname': 'DLDLHIUMNR6001^66 kV G - 3 P.P.K ( BINDAPUR ) GRID','feedername': 'BIND_11KV_307616', 
    # 'subtraname': 'DLDLHIDWKS1138', 'dtname': 'DL-1LDTRJE63053927', 'voltage': '0',
import clipboard
def copy_text():
    clipboard.copy('Remarks - Test')

topmost = True

def windowontop():
    global topmost
    if topmost:
        window.attributes('-topmost',1)
        topmost = False
    else:
        window.attributes('-topmost',0)
        topmost = True


window = tk.Tk()
window.title("RMU assistant")
window.geometry('1250x35')

btn_open_chrome = tk.Button(window,text='Open Browser',bd=5,command=init_driver,width=12)
# btn_open_chrome.pack(padx=5,pady=5)

btn_login_ioms = tk.Button(window,text='Login IOMS',bd=5,command=login_ioms,width=12)
# btn_login_ioms.pack(padx=5,pady=5)

# btn_reg_bd = tk.Button(window,text='Register BD',bd=5,command=generate_bd,width=40)
# btn_reg_bd.pack(padx=5,pady=5)

# btn_read_bdid = tk.Button(window,text='Read BD_ID',bd=5,command=read_bdid,width=40)
# btn_read_bdid.pack(padx=5,pady=5)

# lbl_bdid = tk.Label(window,text='Bxxxxxxxxxx')
# lbl_bdid.pack(padx=5,pady=1)

# btn_restore_bdid = tk.Button(window,text='Restore BD',bd=5,command=restore_bd,width=40)
# btn_restore_bdid.pack(padx=5,pady=5)

# btn_close_bdid = tk.Button(window,text='Close BD',bd=5,command=close_bd,width=40)
# btn_close_bdid.pack(padx=5,pady=5)

btn_Fill_FIR = tk.Button(window,text='Fill FIR',bd=5,command=fill_FIR,width=10)
# btn_Fill_FIR.pack(padx=5,pady=5)

# btn_file_upload_fir = tk.Button(window,text='File Upload FIR',bd=5,command=upload_file_in_FIR,width=40)
# btn_file_upload_fir.pack(padx=5,pady=5)

btn_paste_PSD_ID = tk.Button(window,text='Paste PSD-ID',bd=5,command=paste_psd_id,width=12)
# btn_paste_PSD_ID.pack(padx=5,pady=5)

btn_non_repairable = tk.Button(window,text='Non-Repairable WS',bd=5,command=fill_non_repairable,width=15)
# btn_non_repairable.pack(padx=5,pady=5)

btn_non_repairable_sdo = tk.Button(window,text='Non-Repairable SDO',bd=5,command=non_repairable_sdo,width=15)
# btn_non_repairable_sdo.pack(padx=5,pady=5)
btn_boq_1 = tk.Button(window,text='BOQ-1',bd=5,command=boq_1,width=10)

btn_fin_store = tk.Button(window,text='Finance/Store',bd=5,command=finance_store,width=15)

btn_fill_justification = tk.Button(window,text='Justification',bd=5,command=fill_justification,width=12)
# btn_fill_justification.pack(padx=5,pady=5)

btn_hoto_sdo_1 = tk.Button(window,text='HOTO_1',bd=5,command=hoto_sdo_1,width=10)

btn_hoto_sdo = tk.Button(window,text='HOTO_2',bd=5,command=hoto_sdo,width=10)
# btn_hoto_sdo.pack(padx=5,pady=5)

btn_hoto_workshop = tk.Button(window,text='HOTO WS',bd=5,command=hoto_workshop,width=10)
# btn_hoto_workshop.pack(padx=5,pady=5)

btn_copy_text = tk.Button(window,text='<>',bd=5,command=copy_text,width=5)
btn_top = tk.Button(window,text='top',bd=5,command=windowontop,width=5)
# btn_list = []
# btn_list.append(btn_Fill_FIR)
# btn_list.append(btn_paste_PSD_ID)
# btn_list.append(btn_non_repairable),btn_non_repairable_sdo,btn_fill_justification,btn_hoto_sdo,btn_hoto_workshop)
btn_list = [btn_open_chrome,btn_login_ioms,btn_Fill_FIR,btn_paste_PSD_ID,btn_non_repairable,
btn_non_repairable_sdo,btn_boq_1,btn_fill_justification,btn_fin_store,btn_hoto_sdo_1,btn_hoto_sdo,
btn_hoto_workshop,btn_copy_text,btn_top]



# for btn in btn_list:
#     btn.pack(padx=5,pady=5)

for n,btn in enumerate(btn_list):
    btn.grid(row=1,column=n)

# window.overrideredirect(1)
window.attributes('-topmost',1)

window.mainloop()