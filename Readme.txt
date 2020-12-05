##########################################
UL_Read_existing_config.py
##########################################


To run: python UL_Read_existing_config.py > op.txt

Hardcoded Inputs: (change for monthly runs)
  policy_file = "./PM_INTERNET_OUT.txt"
  splunk_sheet = "./Script inputs/Splunk devices and site mappping/DeviceInventorySplunk10092020.xlsx"
  data_file_from_nw = "./Script inputs/UL-test-wan_priority-6jul-1.txt"
  coral_task_json = "./Script inputs/CMS_Coral_report/20Jul2020/803f6ddf-28e1-4bef-b502-f0957113462d/"
  np_folder = "./Script inputs/NP_download/RawInventory_Export_2020-Sep-10_08-05-05/"

  Smartsheet for master tracker from WANm migration - Downloaded automatically

Outputs:

  ./Script output/UL_Read/UL_devices<timestamp>.xlsx - consolidated output for WAN and SDWAN assessment
  ./Script output/UL_Read/Route-summary<timstamp>/ - folder for site wise route comparison


##########################################
UL_CR_Report_read.py
##########################################

To run: python UL_CR_Report_read.py > op.txt

Hardcoded Inputs: (change for monthly runs)
  filename = "./Script input/C&R/07aug/Audit_Summary_408.xlsm"	
  CR_mapping = "./Script input/C&R/CR to bucket mapping.xlsx"

Outputs:
  ./Script output/CR_summary<timestamp>.xlsx
