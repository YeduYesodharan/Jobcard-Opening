*** Settings ***
Library   SikuliLibrary  mode=OLD
Library   RPA.Browser.Selenium
Library   RPA.Desktop
Library   RPA.Tables
Library   RPA.Excel.Files
library   RPA.Windows
Library   Dialogs
Library    Collections
Library    String
Variables  Variables/variables.py
Library    Libraries/business_operations.py
Resource   Resources/dms_jobcard_extraction.robot
Library    Libraries/mail_send.py
Resource   Resources/Wrappers.robot
Library    Libraries/utility.py
Library    Libraries/ERP_methods.py


*** Variables ***
${imagerootfolder}                       ${CURDIR}${/}..\\Locators
${service_menu_image}                    ${imagerootfolder}\\service_menu.png
${reports_image}                         ${imagerootfolder}\\reports.png
${jobcard_status_report_image}           ${imagerootfolder}\\jobcard_status_report_menu.png
${jobcard_rpt_ratio_image}               ${imagerootfolder}\\jobcard_status_rpt_ratio.png
${jobcard_status_rpt_papersize_title}    ${imagerootfolder}\\jobcard_status_rpt_papersize_title.png
${132_column_select}                     ${imagerootfolder}\\132_column_select.png
${132_column_select_v3}                  ${imagerootfolder}\\132_column_select_v3.png
${132_column_detailed_option}            ${imagerootfolder}\\132_column_detailed_option.png 
${132_column_detailed_option_v3}         ${imagerootfolder}\\132_column_detailed_option_v3.png                            
${unbilled_select_option}                ${imagerootfolder}\\unbilled_select_option.png


${save_report}                           ${imagerootfolder}\\save_report.png
${dms_downloads_status}                  ${imagerootfolder}\\dms_downloads_status.png
${close_csv_popup}                       ${imagerootfolder}\\csv_download_close.png 

${80_column_select}                     ${imagerootfolder}\\80_column_select.png
${unbilled_select}                      ${imagerootfolder}\\unbilled_select.png
${dealer_details_checkbox}              ${imagerootfolder}\\dealer_details_checkbox.png
${dealer_details_checkbox_checked}      ${imagerootfolder}\\dealer_details_checkbox_checked.png
${location_details_checkbox}            ${imagerootfolder}\\location_details_checkbox.png
${location_details_checkbox_checked}    ${imagerootfolder}\\location_details_checkbox_checked.png
${report_formate}                       ${imagerootfolder}\\report_formate.png
${report_formate_select}                ${imagerootfolder}\\report_formate_select.png
${dms_report_save_link}                 ${imagerootfolder}\\dms_report_save_link.png

${excel_select}                         ${imagerootfolder}\\excel_select.png
${submit_report_btn}                    ${imagerootfolder}\\submit_report_btn.png
${status_rpt_save_btn}                  ${imagerootfolder}\\status_rpt_save_btn.png
# ${log_folder}                           ${CURDIR}${/}..\\Log
${log_folder}     C:\\JobcardOpeningIntegrated\\Screenshot
${dealer_select_error_popup}            ${imagerootfolder}\\dealer_select_error_popup.png
${dealer_error_popup_ok_btn}            ${imagerootfolder}\\dealer_error_popup_ok_btn.png
${select_location_error_popup}          ${imagerootfolder}\\select_location_error_popup.png
${service_btn_v3}                       ${imagerootfolder}\\service_btn_v3.png
${report_sizedown_v3}                   ${imagerootfolder}\\report_sizedown_v3.png
${jcstatus_down}                        ${imagerootfolder}\\jcstatus_down.png

#Keyword used to generate Jobcard Carry Over Ratio Report from DMS to

*** Keywords ***
Dms Jobcard Status Report Menu Select
    TRY
        #---------------------------Sikuli change in V3-------------------------#
        # Set Global Variable    ${service_menu_image}
        Set Global Variable    ${service_btn_v3}

        # SikuliLibrary.Wait Until Screen Contain    ${service_menu_image}     ${time_out}
        # SikuliLibrary.Click    ${service_menu_image}
        #---------------------------Sikuli change in V3-------------------------#
        
        SikuliLibrary.Wait Until Screen Contain    ${service_btn_v3}     ${time_out}
        SikuliLibrary.Click    ${service_btn_v3}

        SikuliLibrary.Wait Until Screen Contain    ${reports_image}     ${time_out}
        SikuliLibrary.Click    ${reports_image}

        SikuliLibrary.Wait Until Screen Contain    ${jobcard_status_report_image}     ${time_out}
        SikuliLibrary.Click    ${jobcard_status_report_image}

        SikuliLibrary.Wait Until Screen Contain    ${jobcard_rpt_ratio_image}     ${time_out}
        SikuliLibrary.Click    ${jobcard_rpt_ratio_image}

        Capture Screenshot
        
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message} 
    END
132 Column Select
    # SikuliLibrary.Wait Until Screen Contain    ${132_column_select}     ${time_out}
    SikuliLibrary.Wait Until Screen Contain    ${132_column_select_v3}     ${time_out}
    # SikuliLibrary.Click    ${132_column_select}
    SikuliLibrary.Click    ${132_column_select_v3}
    RPA.Desktop.Press Keys    1
    RPA.Desktop.Press Keys    enter

Unbilled Column Select
    SikuliLibrary.Wait Until Screen Contain    ${unbilled_select}     ${time_out}
    SikuliLibrary.Click    ${unbilled_select}
    RPA.Desktop.Press Keys    U
    RPA.Desktop.Press Keys    enter

Dealer Detail Check 
    SikuliLibrary.Wait Until Screen Contain    ${dealer_details_checkbox}     ${time_out}  
    Sleep    1s          
    SikuliLibrary.Click    ${dealer_details_checkbox}   
    Sleep    1s

Location Details Check 
    SikuliLibrary.Wait Until Screen Contain    ${location_details_checkbox}     ${time_out}  
    Sleep    1s          
    SikuliLibrary.Click    ${location_details_checkbox}
    Sleep    1s

Report Format Select
    SikuliLibrary.Wait Until Screen Contain    ${report_formate}     ${time_out}
    SikuliLibrary.Click    ${report_formate}
    Sleep    1s
    # RPA.Desktop.Press Keys    E
    RPA.Desktop.Press Keys    Down
    Sleep    1s
    RPA.Desktop.Press Keys    Down
    Sleep    1s
    RPA.Desktop.Press Keys    enter
    Sleep    1s
    
Generate Jobcard Status Report
    TRY

        SikuliLibrary.Wait Until Screen Contain    ${jobcard_status_rpt_papersize_title}     ${time_out}
        # Sleep    8
        
        ${132_column_detailed_option_exists}=    Set Variable    False
        ${attempt1}=    Set Variable    0
        WHILE    '${132_column_detailed_option_exists}' == 'False' and ${attempt1} < 5
            132 Column Select
            Capture Screenshot
            # ${132_column_detailed_option_exists}=    SikuliLibrary.Exists    ${132_column_detailed_option}
            ${132_column_detailed_option_exists}=    SikuliLibrary.Exists    ${132_column_detailed_option_v3}
            ${attempt1}=    Evaluate    ${attempt1} + 1
        END


        ${unbilled_select_option_exists}=    Set Variable    False
        ${attempt1}=    Set Variable    0
        WHILE    '${unbilled_select_option_exists}' == 'False' and ${attempt1} < 5
            Unbilled Column Select
            Capture Screenshot
            ${unbilled_select_option_exists}=    SikuliLibrary.Exists    ${unbilled_select_option}
            ${attempt1}=    Evaluate    ${attempt1} + 1
        END
    
        
        ${dealer_details_checkbox_checked_exists}=    Set Variable    False
        ${attempt1}=    Set Variable    0
        WHILE    '${dealer_details_checkbox_checked_exists}' == 'False' and ${attempt1} < 5
            Dealer Detail Check
            Capture Screenshot
            ${dealer_details_checkbox_checked_exists}=    SikuliLibrary.Exists    ${dealer_details_checkbox_checked}
            ${attempt1}=    Evaluate    ${attempt1} + 1
        END


        ${location_details_checkbox_checked_exists}=    Set Variable    False
        ${attempt1}=    Set Variable    0
        WHILE    '${location_details_checkbox_checked_exists}' == 'False' and ${attempt1} < 5
            Location Details Check
            Capture Screenshot
            ${location_details_checkbox_checked_exists}=    SikuliLibrary.Exists    ${location_details_checkbox_checked}
            ${attempt1}=    Evaluate    ${attempt1} + 1
        END

        # Report Format Select
        ${report_formate_select_exists}=    Set Variable    False
        ${attempt1}=    Set Variable    0
        WHILE    '${report_formate_select_exists}' == 'False' and ${attempt1} < 5
            Report Format Select
            ${report_formate_select_exists}=    SikuliLibrary.Exists    ${report_formate_select}
            ${attempt1}=    Evaluate    ${attempt1} + 1
        END
        Capture Screenshot
        # ------------------------------------------------

        ${all_conditions_passed}=    Set Variable    True

        IF    '${132_column_detailed_option_exists}' != 'True'
            Log    132 Column selection failed
            ${all_conditions_passed}=    Set Variable    False
        END

        IF    '${unbilled_select_option_exists}' != 'True'
            Log    Unbilled Column selection failed
            ${all_conditions_passed}=    Set Variable    False
        END

        IF    '${dealer_details_checkbox_checked_exists}' != 'True'
            Log    Dealer Details checkbox selection failed
            ${all_conditions_passed}=    Set Variable    False
        END

        IF    '${location_details_checkbox_checked_exists}' != 'True'
            Log    Location Details checkbox selection failed
            ${all_conditions_passed}=    Set Variable    False
        END

        IF    '${report_formate_select_exists}' != 'True'
            Log    Report Format selection failed
            ${all_conditions_passed}=    Set Variable    False
        END

        IF    '${all_conditions_passed}' == 'True'
            Capture Screenshot
            SikuliLibrary.Wait Until Screen Contain    ${submit_report_btn}     ${time_out}
            SikuliLibrary.Click    ${submit_report_btn}

            # --- New change added for "DMS report add handle for the windows that appear in the report download page for the dealer or branch selection failure or any other field entry failures."
            ${dealer_select_error_popup_exists}=    SikuliLibrary.Exists    ${dealer_select_error_popup}    ${avg_time}
            IF    ${dealer_select_error_popup_exists}==$True
                SikuliLibrary.Click    ${dealer_error_popup_ok_btn}
                Sleep    0.5s
                Dealer Detail Check
                ${dealer_details_checkbox_checked_exists}=    SikuliLibrary.Exists    ${dealer_details_checkbox_checked}    ${avg_time}
                IF    ${dealer_details_checkbox_checked_exists}==$True
                    Log    Successfully selected the Dealer.
                    SikuliLibrary.Click    ${submit_report_btn}
                ELSE
                   Fail    Unable to select the Dealer.
                END
            END

            ${select_location_error_popup_exists}=    SikuliLibrary.Exists    ${select_location_error_popup}    ${avg_time}
            IF    ${select_location_error_popup_exists}==$True
                SikuliLibrary.Click    ${dealer_error_popup_ok_btn}
                Sleep    0.5s
                Location Details Check
                ${location_details_checkbox_checked_exists}=    SikuliLibrary.Exists    ${location_details_checkbox_checked}    ${avg_time}
                IF    ${location_details_checkbox_checked_exists}==$True
                    Log    Successfully selected the Location.
                    SikuliLibrary.Click    ${submit_report_btn}
                ELSE
                   Fail    Unable to select the Location.
                END
            END   
            # ---------------------------------------------------------------------------------------------
        ELSE
            Fail    One or more report download setup steps failed after 5 attempts
        END

        # ------------------------------------------------
        Log    ${max_time_out}
        #-------------------------------------commented for auto save csv in DMS3 Sikuli--------------------------------------------#
        # Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${dms_report_save_link}     ${max_time_out} 
        # ${dms_report_save_link_exists}=    SikuliLibrary.Exists    ${dms_report_save_link}
        # IF    ${dms_report_save_link_exists}==False
        #     Fail    Failed To Download DMS Report.
        # ELSE
        #    SikuliLibrary.Click    ${dms_report_save_link}
        # END 
        #-------------------------------------commented for auto save csv in DMS3 Sikuli--------------------------------------------#
        #---- USed to check if the dms report downloaded successfully -----
        ${dms_report_exist_status}    Check For The Dms Report Present In The Download Folder
        IF    ${dms_report_exist_status} == False

            # RPA.Browser.Selenium.Close All Browsers
            ERP_methods.Close Browser Processes
            Fail    Failed To Download DMS Report.
        ELSE

            # RPA.Browser.Selenium.Close All Browsers
            ERP_methods.Close Browser Processes
            #-------------------------------------commented for auto save csv in DMS3 Sikuli--------------------------------------------#
            # ${dms_downloaded_status_message}=    Set Variable    ${EMPTY}
            # ${attempt1}=    Set Variable    0

            # WHILE    $dms_downloaded_status_message == "" and ${attempt1} < 3
            #     Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${dms_downloads_status}     ${time_out}
            #     ${dms_downloads_status_exists}=    SikuliLibrary.Exists    ${dms_downloads_status}

            #     IF    ${dms_downloads_status_exists} == False
            #         Fail    Failed To Download DMS Report. Status message not found.
            #     ELSE
            #         ${dms_downloaded_status_message}=    SikuliLibrary.Get Text    ${dms_downloads_status}
            #         Log    Current Status: ${dms_downloaded_status_message}
            #     END
            #     ${attempt1}=    Evaluate    ${attempt1} + 1
            # END
            # Log    ${dms_downloaded_status_message}
            # ${contains}=    Evaluate    "${dms_report_download_status_message}" in """${dms_downloaded_status_message}"""

            # IF    "${contains}" == "True"
            #     Log    DMS report downloaded successfully.

            #     Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${close_csv_popup}     ${time_out} 
            #     ${close_csv_popup_exists}=    SikuliLibrary.Exists    ${close_csv_popup}

            #     IF    ${close_csv_popup_exists} == False
            #         Fail    Failed To Close CSV Popup After Download.
            #     ELSE
            #         SikuliLibrary.Click    ${close_csv_popup}
            #     END 
            # ELSE
            #     Log    Failed To Download DMS Report.
            #     Fail    Failed To Download DMS Report.
            # END
            #-------------------------------------commented for auto save csv in DMS3 Sikuli--------------------------------------------#
        END

    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    

Downloaded Dms Vs Erp Report Validation
    [Arguments]    ${curr_date}    ${erp_combined_report_path}
    TRY
        #----- Preparing Results Downloads folder path ---------
        ${results_download_folder}    Prepare Results Folder Path    ${curr_date}  

        #----- Move the DMS report from system downloads to the path "process related folder/ current_date folder/ Downloads" -------
        ${consolidated_report_path}    ${carry_over_ratio_report_path}        move_dms_report_to_currentdate_folder    ${curr_date}
        
        #----- Check if the DMS report is empty or not -----------
        ${carry_over_ratio_report_exist}    utility.Check File Exist And Empty    ${carry_over_ratio_report_path}

        IF   "${carry_over_ratio_report_exist}" == "${False}"
            
            move_report_to_destination_folder    ${carry_over_ratio_report_path}    ${results_download_folder}

            Log    ${carry_over_ratio_report_path}
            Send Email Output     ${curr_date}    ${report_name}    There is no Jobcard details found in the DMS report to proceed.    ${carry_over_ratio_report_path}
           
            Fail    There is no Jobcard details found in the DMS report to proceed.
        ELSE
            ${failure_report_path}    Prepare Failure Report Path    ${curr_date} 

            validate_erp_report_vs_dms_report    ${carry_over_ratio_report_path}    ${erp_combined_report_path}    ${consolidated_report_path}    ${failure_report_path}
            
            update_final_jobcard_details_report    ${consolidated_report_path}    ${carry_over_ratio_report_path}
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Close All Applications
        Capture Screenshot 
        Fail    ${error_message} 
    END
    [Return]    ${consolidated_report_path}    ${carry_over_ratio_report_path}


# *** Tasks ***
# Demo  
#     Generate Jobcard Status Report
