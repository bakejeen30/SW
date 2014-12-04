'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "Note - Forwarding Address"
start_time = timer

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'*************************************************************Dialogue
BeginDialog forwarding_address_received, 0, 0, 276, 225, "Forwarding Address "
  Text 5, 10, 140, 10, "Forwarding Address Received from USPS"
  Text 5, 35, 50, 10, "Case Number:"
  EditBox 55, 30, 90, 15, case_number
  Text 5, 55, 50, 10, "Date Received:"
  EditBox 65, 50, 120, 15, date_received
  Text 5, 75, 60, 10, "Case Note Entry:"
  EditBox 70, 70, 155, 15, case_note
  ButtonGroup ButtonPressed
    PushButton 235, 70, 35, 15, "ADDR", stat_addr
  CheckBox 5, 110, 10, 20, "", forwarded_check
  Text 20, 115, 110, 15, "Forwarded mail to new address"
  CheckBox 5, 130, 10, 10, "", change_report_check
  Text 20, 130, 185, 10, "Change Report form mailed to new forwarding address"
  CheckBox 5, 145, 10, 10, "", TIKL_check
  Text 20, 145, 75, 10, "TIKL for 10 Day Return"
  CheckBox 5, 155, 10, 15, "", address_standardized_check
  Text 20, 160, 70, 10, "Address Standarized"
  Text 5, 185, 70, 10, "Sign your case note:"
  EditBox 80, 180, 130, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 210, 50, 15
    CancelButton 225, 210, 50, 15
EndDialog
'THE SCRIPT
'*********************************************************Connect to BlueZone
EMConnect ""

'***********************************************************8***Does the dialog
Dialog forwarding_address_received
If buttonpressed = 0 then STOPSCRIPT


'*****************************************************Check for MAXIS
'EMSendKey "<enter>"
'EMWaitReady 0, 0
'EMReadScreen MAXIS_check, 5, 1, 39
'If MAXIS_check <> "MAXIS" then
'	MsgBox "You are not in MAXIS"
'	StopScript
'End If
MAXIS_check_function

'Find MAXIS case number if there is one
call MAXIS_case_number_finder (case_number)

'Finds Verifs Postponed
call navigate_to_screen ("STAT", "ADDR")
EMReadScreen verifs_postponed, 2, 9, 74


'***************************This DO Loop gets back to the SELF screen
DO
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"


'******************Navigates to CASE/NOTE for the case number entered
EMWriteScreen "CASE", 16, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "NOTE", 21, 70


'****************************************Transmits to the next screen
EMSendKey "<enter>"
EMWaitReady 0, 0


'*****************************************PF9 to open a new case note
EMSendKey "<PF9>"
EMWaitReady 0, 0


'***************************************************Writes to the Case Note
call write_new_line_in_case_note (">>>Forwarding Address<<<")
call write_editbox_in_case_note ("Date Received", date_received, 6)
call write_editbox_in_case_note ("Note", case_note, 6)
If forwarded_check = 1 then call write_new_line_in_case_note ("* Forwarded mail to new address")
If change_report_check = 1 then write_new_line_in_case_note ("* Change Report form mailed to forwarding address") 
IF TIKL_check = 1 then call write_new_line_in_case_note ("* TIKLed for 10 day return")
IF address_standardized_check = 1 then call write_new_line_in_case_note ("* Address Standardized")
call write_new_line_in_case_note ("---")
call write_new_line_in_case_note (worker_signature)Enter file contents here
