*** Settings ***
Library   SikuliLibrary  mode=OLD
Library   RPA.Desktop
Library   RPA.Tables
Library   RPA.Excel.Files
library   RPA.Windows
Library   Dialogs
Library   Collections
Library   OperatingSystem
Library    String
Variables  Variables/variables.py
Library    Libraries/business_operations.py
Resource   Resources/dms_login.robot
Resource   Resources/Main_Flow.robot 
Library    Libraries/mail_send.py


*** Variables ***
${today_radio_button_image}          ${imagerootfolder}\\today_radio_button.png
# ${today_radio_button_image}          ${imagerootfolder}\\current_month_radio.png
${export_btn_image}                  ${imagerootfolder}\\export_arrow.png
${xlsx_extn_image}                   ${imagerootfolder}\\xlsx_img.png
${erp_close_btn}                    ${imagerootfolder}\\erp_close_button.png
${yes_btn}                          ${imagerootfolder}\\yes_btn.png
${3_dot_img}                        ${imagerootfolder}\\3_dot.png
${3_dot_select}                      1|1|1|2|1|1|1|1|1|5
${submit_btn}                        1|1|2|2
${downloaded_status_msg_path}        1|1|1|1|1|1
${popup_close_button}                1|1|1|1|2|4
${log_folder}    ${CURDIR}${/}..\\Screenshot



*** Keywords ***
ERP Initial Report Download
    [Arguments]    ${path}    ${region_mapping_sheet}    ${location}
    TRY
        ${login_status}   ERP Login

        IF    ${login_status} == ${True}
            
            Regular Service Jobcard Report Menu Select        
            ${erp_report_regular_full_name}    Jobcard Report Download    ${regular}
            
            Bodyshop Service Jobcard Report Menu Select
            ${erp_report_bodyshop_full_name}    Jobcard Report Download    ${bodyshop}
            #${erp_report_bodyshop_full_name}  Set Variable   E:\\JobcardOpeningIntegrated\\ProcessRelatedFolders\\2025-04-02\\Downloads\\ERP_Report_Bodyshop_Service20250402105247.xlsx

            ${downloads_dir}    Prepare Erp Report Save Location Path
            Log    ${location} 
            Log    ${region_mapping_sheet}
            ${erp_combined_report_path}    Combine ERP Regular And Bodyshop Report        ${erp_report_regular_full_name}    ${erp_report_bodyshop_full_name}    ${downloads_dir}    ${region_mapping_sheet}    ${location}

            # ------   Script changes for single JC run with DMS and ERP. --------
            Erp Regular And Bodyshop Page Navigation
        ELSE
            Fail    Unable to login to the ERP
        END              
    EXCEPT   AS   ${message}
        Close Erp
        # Send Email Output     ${curr_date}    ${report_name}    ${message}    ${EMPTY}    
        Fail   ${message}
    END
    [Return]    ${erp_combined_report_path} 

Erp Regular And Bodyshop Page Navigation
    TRY
        ${rsjnav_status}    Service Jobcard Navigation  
        IF    ${rsjnav_status} == ${True}
                ${bsjnav_status}   BodyShop Jobcard Navigation
                IF    ${bsjnav_status} == ${True}
                    Log    navigation successfull
                END
        END
    EXCEPT   AS   ${message}  
        Fail   ${message}
    END
    

Close Erp
   Select Window
   SikuliLibrary.Click    ${erp_close_btn}
   SikuliLibrary.Click    ${yes_btn}
   Run Keyword And Ignore Error    SikuliLibrary.Click    ${erp_close_btn}
   Sleep    ${normal_sleep}
   Run Keyword And Ignore Error    SikuliLibrary.Click    ${erp_close_btn}
   Run    taskkill /f /im WLauncher.exe


Common Menu Select Upto Service
    TRY        
        Click Action   name:${erp_reports_menu}
        Click Action   name:${erp_reports_menu} > ${menu_AutoDms}
        Click Action   name:${erp_reports_menu} > ${menu_service}
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Fail   ${error_message}
    END 


#change yy
Regular Service Jobcard Report Menu Select
    TRY

        Select Window
        Click Action   name:${erp_reports_menu}
        Click Action   name:${erp_reports_menu} > ${menu_AutoDms}
        Click Action   name:${erp_reports_menu} > ${menu_service}     
        Click Action   name:${erp_reports_menu} > ${menu_regular_service}
        Click Action   name:${erp_reports_menu} > ${erp_jobcard_menu}
        Click Action   name:${erp_reports_menu} > ${erp_service_jobcard_menu}
        Click Action   name:${erp_reports_menu} > ${erp_standard_report_menu} 

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${3_dot_img}     ${time_out}
        # ${3_dot_img_exists}=    SikuliLibrary.Exists    ${3_dot_img}   
        # IF    ${3_dot_img_exists}==True

    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message} 
    END

Jobcard Menu Select For Regular Service TestDB
    TRY
        Press Keys Action    TAB
        Sleep    ${normal_sleep}
        Press Keys Action    TAB
        Press Keys Action    ENTER
    EXCEPT    AS   ${message}
        Capture Screenshot 
        Fail    ${message}
    END

Jobcard Menu Select For Regular Service Prod
    TRY
        Press Keys Action    TAB
        Sleep    ${normal_sleep}
        Press Keys Action    ENTER
    EXCEPT    AS   ${message}
        Capture Screenshot 
        Fail    ${message}
    END
    
#change yy  
Bodyshop Service Jobcard Report Menu Select
    TRY
        Select Window
        Click Action   name:${erp_reports_menu}
        Click Action   name:${erp_reports_menu} > ${menu_AutoDms}
        Click Action   name:${erp_reports_menu} > ${menu_service}     
        Click Action   name:${erp_reports_menu} > ${menu_bodyshop_service}

        Press Keys Action    TAB
        Sleep    ${normal_sleep}
        Press Keys Action    ENTER
        Sleep    ${normal_sleep}
        Press Keys Action    ENTER
        Sleep    ${normal_sleep}
        Press Keys Action    ENTER

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${3_dot_img}     ${time_out}
              
    EXCEPT  AS   ${error_message}  
        Capture Screenshot         
        Fail    ${error_message}
    END


Standard Report Menu Select
    TRY
        Press Keys Action    TAB
        Press Keys Action    ENTER
    EXCEPT    AS    ${message}
        Capture Screenshot 
        Fail   ${message} 
    END
    

Jobcard Report Download
    [Arguments]    ${service_type}
    TRY
        # Select Window
        Sleep    ${normal_sleep}
        click Action    path:"${3_dot_select}"

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${today_radio_button_image}     ${time_out}
        ${today_radio_button_image_exists}=    SikuliLibrary.Exists    ${today_radio_button_image}   
        IF    ${today_radio_button_image_exists}==True
            SikuliLibrary.Click    ${today_radio_button_image}
            Sleep    ${normal_sleep}
            Click Action    path:"${submit_btn}"
            #---- Calling a keyword used to download ERP report ------
            ${erp_report_full_name}    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    Export As Excel Selection    ${service_type}
            Return From Keyword    ${erp_report_full_name}
        ELSE
            Log    DMS report download failed due to Today's option is not available.
            Capture Screenshot
            Fail   DMS report download failed due to Today's option is not available. 
        END
    EXCEPT    AS    ${message}
        Capture Screenshot 
        Send Email Output     ${curr_date}    ${report_name}    DMS report download failed due to Today's option is not available.    ${EMPTY}    
        Fail   ${message} 
    END
    
    

Export As Excel Selection
    [Arguments]    ${service_type}
    TRY
        # Select Window
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${export_btn_image}     ${time_out}
        ${export_btn_image_exists}=    SikuliLibrary.Exists    ${export_btn_image}   
        IF    ${export_btn_image_exists}==True
            # SikuliLibrary.Click    ${export_btn_image}
            RPA.Desktop.Press Keys    alt    x
            Sleep    ${normal_sleep}
            SikuliLibrary.Click    ${xlsx_extn_image}  
            #---- Generating the ERP report path -----    
            ${downloads_dir}    Prepare Erp Report Save Location Path 

            #--- Generating the ERP report file name ------
            ${downloaded_erp_report_name}    Prepare File Name Of Downloaded Erp Report  ${service_type}  

            ${erp_report_full_name}    Set Variable    ${downloads_dir}\\${downloaded_erp_report_name}.xlsx
            Sleep    ${Min_time}
            # Sleep    3s
            Type Text    ${downloads_dir}\\${downloaded_erp_report_name}     enter=True
            
            Check For Erp Report Downloaded Successfully

            Return From Keyword    ${erp_report_full_name}
        ELSE
            Fail    Unable to download Regular Jobcard Report
        END
    EXCEPT    AS    ${message}
        Capture Screenshot
        Fail    ${message}
    END
    

Prepare Erp Report Save Location Path
    TRY
        ${project_root}     business_operations.get_process_root_directory
        ${downloads_dir}    Set Variable    ${project_root}\\${process_related_folder_name}\\${curr_date}\\${downloads_dir_name}
        Log    ${downloads_dir}
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Fail    ${error_message} 
    END    
    [Return]    ${downloads_dir}


Select Window
    Sleep    ${Min_sleep}
    ${windows}=  List Windows
    FOR  ${window}  IN  @{windows}
        ${title}=   utility.Get Title Starting With    ${window}    Wings ERP 23E - Web Client
        log  ${title}
        Exit For Loop If    '${title}'!='None'
    END
    Control Window   name:"${title}"
        
Check For Erp Report Downloaded Successfully
    TRY
        Select Window
        Capture Screenshot
        ${downloaded_status_msg}    Get Text Action    path:"${downloaded_status_msg_path}"
        Capture Screenshot
        Log    ${downloaded_status_msg}
        Log    ${report_downloaded_status_message}

        ${contains}=    Evaluate   """${report_downloaded_status_message}""" in r"""${downloaded_status_msg}"""
        
        IF    "${contains}" == "True"
            Log    ${downloaded_status_msg}
            # TODO File exist
            # IF    not ${var_in_py_expr2}
            #     Fail
            # END
            Sleep    1s
            Run Keyword And Ignore Error    Click Action    path:"${popup_close_button}"
            Sleep    0.5s
            Run Keyword And Ignore Error    RPA.Desktop.Press Keys    enter
            Sleep    0.5s
            RPA.Desktop.Press Keys    alt    c

        ELSE
            Log    ERP Report Download Failed
            Fail    ERP Report Download Failed
        END
    EXCEPT    AS   ${message}
        Capture Screenshot 
        Fail    ${message}
    END

