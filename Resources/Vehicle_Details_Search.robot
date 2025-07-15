*** Settings ***
Library             RPA.Desktop
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             OperatingSystem
Library             String
Library             Collections
Library             RPA.Windows
Library             Libraries/ERP_methods.py   
Library             RPA.RobotLogListener
Variables           Variables/variables.py 
Resource            JobCard_Tab_Data_entry.robot
Resource            New_Chassis_Creation.robot
Resource            Customer_Creation&Update_reg.robot
Resource            Erp_Recall_Marking.robot
Resource            Wrappers.robot

*** Variables ***
${Search_get_Vehicle_details}    3|1|1|1|1|1|1|1|5|1
${Reg_No}    1|2|1|1
${Chassis_No}    1|2|1|2
${Mobile_no}    1|2|1|4
${Search_inputbox}   1|2|1|6|2    #itchamp
${bodyshop_edit_customer}    3|1|1|1|1|1|1|1|6|1
${go_button}         1|2|1|6|3    #itchamp
${no_record_msg_content}    1|1|3
${no_rec_ok_button}    1|1|1
${no_rec_header}    1|1|2
${no_rec_close_button}    1|1|2|1
${ok_to_tabs}    1|3|2
${chassis_creation_btn}    1|2|1|7
${create_chssis_chckbox}    1|1|1|1|1|1
${alert_box}    1|1
${search_window_close}    1|3|1
# ${log_folder}     ${CURDIR}${/}..\\Log
# ${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
${log_folder}     C:\\JobcardOpeningIntegrated\\Screenshot
${Body_Branch_path}    3|1|1|1|1|1|1|1|4|2|1|1    
${Body_location_path}    3|1|1|1|1|1|1|1|5|2|1|1
${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx
*** Keywords ***

Vehicle Details Search with Registration Number
    [Arguments]    ${Registration No.}   ${Job Card No.}    ${Input_Sheet_Path}
    TRY        
        
        Click Action    path:${Search_inputbox}
        Type Text Action    ${Registration No.}
        # Sleep    ${Min_time}
        Click Action    path:${go_button}
        # Sleep    ${Min_time}
        Capture Screenshot
        ${alert_box_visible_1}   Get Text Action Maximum Retry    path:${alert_box}
        
        Capture Screenshot
        RETURN    ${alert_box_visible_1}
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Vehicle Search with Registration Number    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    ${search_window_close}
        # Exit For Loop
    END 

Vehicle Details Search with Chassis Number  
    [Arguments]   ${Chassis No.}    ${Job Card No.}    ${Input_Sheet_Path}
    TRY
        # Sleep    ${Min_time} 
        Click Action    path:${Search_inputbox}
        Type Text Action    ${Chassis No.}
        # Sleep    ${Min_time}
        Click Action    path:${go_button}
        Sleep    ${Min_time}
        Capture Screenshot
        ${alert_box_visible_2}   Get Text Action Maximum Retry   path:${alert_box}
        Capture Screenshot
        RETURN    ${alert_box_visible_2}
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Vehicle Search with Chassis Number.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    ${search_window_close}
        # Exit For Loop
    END 

Vehicle Details Search with Mobile Number  
    [Arguments]    ${Phone & Mobile No.}    ${Job Card No.}    ${Input_Sheet_Path}
    TRY
        # Sleep    ${Min_time}
        Click Action    path:${Search_inputbox}  
        Type Text Action    ${Phone & Mobile No.}
        # Sleep    ${Min_time}
        Click Action    path:${go_button}
        # Sleep    ${Min_time} 
        Capture Screenshot
        ${alert_box_visible_3}   Get Text Action Maximum Retry    path:${alert_box}
        Capture Screenshot
        RETURN    ${alert_box_visible_3}
    EXCEPT  AS   ${error_message}
       
       Log    ${error_message}
       Capture Screenshot
       update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Vehicle Search with Mobile Number.    ${exception_reason_column_name}
       update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
       Click Action    ${search_window_close}
    #    Exit For Loop
    END 

Bodyshop Branch and Location path
    [Arguments]    ${Job Card No.}    ${Input_Sheet_Path}    ${Branch}
    TRY
        Capture Screenshot
        Click Action Maximum Retry    path:${Body_Branch_path}
        ${branch_code}    ERP_methods.Get Erp Branch Location Code    ${branch_mapping}    ${Branch}
        IF    '${branch_code}' != 'branch code not available'
       
            RPA.Desktop.Press Keys    ctrl  a
            Type Text Action    ${branch_code}
            Press Keys Action    enter
            Click Action   path:${Body_Location_path}
            Click Action    path:${bodyshop_edit_customer}
            RETURN  ${True}
        ELSE
            RETURN  ${False}
        END
        

    EXCEPT  AS   ${error_message}
       
       Log    ${error_message}
       Capture Screenshot
       update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while adding the BodyShop Jobcard Branch and Location path    ${exception_reason_column_name}      
       update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
       Press Keys Action    esc  
       Press Keys Action    enter
    #    Exit For Loop
    END 
Service Jobcard Branch and Location path  
    [Arguments]    ${Job Card No.}    ${Input_Sheet_Path}    ${Branch}
    TRY
        Capture Screenshot
        Click Action Maximum Retry   path:${Branch_path}
        ${branch_code}    ERP_methods.Get Erp Branch Location Code    ${branch_mapping}    ${Branch}
        IF    '${branch_code}' != 'branch code not available'
   
            RPA.Desktop.Press Keys    ctrl  a
            Type Text Action    ${branch_code}
            Press Keys Action    enter
            Click Action   path:${Location_path}
            Click Action    path:${Search_get_Vehicle_details}
            RETURN  ${True}
        ELSE
            RETURN  ${False}
        END    
    EXCEPT  AS   ${error_message}
       
       Log    ${error_message}  
       Capture Screenshot
       update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while adding the Service Jobcard Branch and Location path    ${exception_reason_column_name}  
       update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
       Press Keys Action    esc  
       Press Keys Action    enter
    #    Exit For Loop
    END 


