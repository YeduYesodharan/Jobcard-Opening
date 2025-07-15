*** Settings ***
Library   RPA.Tables
Library   RPA.Excel.Files
Resource  Resources/dms_jobcard_extraction.robot
Resource  Resources/dms_process_start.robot
Variables  Variables/variables.py

*** Keywords ***
Read Summary File
    [Arguments]    ${summary_file}
    TRY
        Open Workbook    ${summary_file}
        ${excel_details}=    Read Worksheet As Table    header=True
        Close Workbook
        Filter Table By Column    ${excel_details}    DMS Execution Status    !=    Success
        Filter Table By Column    ${excel_details}    Exception Reason    !=    Already Entered In ERP

        ${row_count}    Get Length    ${excel_details}
        IF    ${row_count} == 0
            Log    There is no jobcard details found in DMS report to proceed with ERP
            Fail    There is no jobcard details found in DMS report to proceed with ERP
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Fail    ${error_message} 
    END
    [Return]  ${excel_details}

Read Failure Jobcard Extraction In Consolidated Report
    [Arguments]    ${consolidated_report_path}
    TRY
        Open Workbook    ${consolidated_report_path}
        ${failure_jobcard_rows}=    Read Worksheet As Table    header=True
        Close Workbook
        Filter Table By Column    ${failure_jobcard_rows}    DMS Execution Status    ==    Fail

        ${row_count}    Get Length    ${failure_jobcard_rows}
        
        IF  ${row_count} == 0
            Return From Keyword    False    ${failure_jobcard_rows}
        ELSE
            Return From Keyword    True    ${failure_jobcard_rows}
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Fail    ${error_message} 
    END

Read Consolidated Report
    [Arguments]    ${Input_Sheet_Path}
    TRY
        Open Workbook    ${Input_Sheet_Path}
        ${excel_details}=    Read Worksheet As Table    header=True
        Close Workbook
        Filter Table By Column    ${excel_details}    DMS Execution Status    ==    ${DMS_Status1}
        Filter Table By Column    ${excel_details}    Recall Status    ==    ${status}
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Fail    ${error_message} 
    END
    [Return]  ${excel_details}

    

    

    

    

    