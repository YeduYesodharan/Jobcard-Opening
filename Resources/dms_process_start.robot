*** Settings ***
Documentation       Template robot main suite.
Library    OperatingSystem
Resource   Resources/dms_login.robot
Resource   Resources/dms_generate_jobcard_report.robot
Resource   Resources/dms_jobcard_extraction.robot
Resource   Resources/utility.robot
Library    Libraries/business_operations.py
Resource   Resources/Wrappers.robot
Resource   Resources/erp_initial_report_download.robot
Variables  Variables/variables.py
Library    Libraries/ERP_methods.py


*** Variables ***
# ${log_folder}                            ${CURDIR}${/}..\\Log
# ${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
${log_folder}     C:\\JobcardOpeningIntegrated\\Screenshot
${imagerootfolder}                       ${CURDIR}${/}..\\Locators
${exit_btn_logout_image}                 ${imagerootfolder}\\exit_btn_logout.png
${exeed_maximum_idle_time_popup_image}   ${imagerootfolder}\\exeed_maximum_idle_time_popup.png
${sessionout_ok_btn_image}               ${imagerootfolder}\\sessionout_ok_btn.png

                                            


*** Keywords ***
DMS Data Extraction Process
    [Arguments]    ${curr_date}    ${erp_combined_report_path}    ${date_timestamp}
    TRY
        #-----Keyword used to login to the DMS----
        #-----Return Value: login_status: True or False -----
        # ${login_status}    Run Keyword And Return Status     Login To DMS
        ${login_status}    Login To DMS
        IF    ${login_status}==True
            
            #-----Keyword used select Dms Jobcard Status Report Menu for download----
            Wait Until Keyword Succeeds    ${retry}    ${average_sleep}    Dms Jobcard Status Report Menu Select
            
            #-----Keyword/Navigation used to generate Jobcard Carry Over Ratio Report from DMS----
            #-----Return Value: report_download_status: True or False -----
            ${report_download_status}   Run Keyword And Return Status    Wait Until Keyword Succeeds    ${retry}    ${average_sleep}    Generate Jobcard Status Report
            
            IF    ${report_download_status}==True 
                
                #-----Keyword used to validate Jobcard Carry Over Ratio Report from DMS Vs ERP report----
                #-----Return Value: consolidated_report_path, carry_over_ratio_report_path-----
                #-----Argument: curr_date, erp_combined_report_path-----
                ${consolidated_report_path}    ${carry_over_ratio_report_path}   Downloaded Dms Vs Erp Report Validation    ${curr_date}    ${erp_combined_report_path}

                Log    ${consolidated_report_path}
                Log    ${carry_over_ratio_report_path}

                #-----Keyword used to extract DMS jabcard details----
                #-----Return Value: extraction_status True or False-----
                #-----Argument: consolidated_report_path, carry_over_ratio_report_path, failure_jobcard_rows-----
                ${extraction_status}   Run Keyword And Return Status    Jobcard Extraction Process    ${consolidated_report_path}    ${carry_over_ratio_report_path}    ${EMPTY}    ${date_timestamp}

                Log    ${extraction_status}

                IF    ${extraction_status}==True
                    
                    #-----Keyword used to extract the failure jobcard extraction details from consolidated_report----
                    #-----Return Value: status True or False, failure_jobcard_rows-----
                    #-----Argument: consolidated_report_path-----
                    ${status}    ${failure_jobcard_rows}    Read Failure Jobcard Extraction In Consolidated Report    ${consolidated_report_path}
                    

                    IF    ${status}==True
                        
                        #-----Keyword used to extract the jobcard details from DMS----
                        #-----Return Value: extraction_status True or False, failure_jobcard_rows-----
                        #-----Argument: consolidated_report_path, carry_over_ratio_report_path, failure_jobcard_rows-----
                        ${extraction_status}   Run Keyword And Return Status    Jobcard Extraction Process    ${consolidated_report_path}    ${carry_over_ratio_report_path}    ${failure_jobcard_rows}    ${date_timestamp}

                        IF    ${extraction_status}==True
                            #-----Keyword used to Prepare Failure Report path----
                            ${failure_report_path}    Prepare Failure Report Path    ${curr_date}
                            #-----Keyword used to Creating Failure Report----
                            Create Failure Report    ${failure_report_path}
                            #-----Keyword used to Updating Failure Report----
                            Update Failure Report    ${consolidated_report_path}    ${failure_report_path}
                        END
                    END
                END
                #-----Keyword used to Move Consolidated Report To Processed Folder after single execution----
                #-----Argument: consolidated_report_path-----
                ${consolid_report_path_in_processed_erp_dir}    Move Consolidated Report To Processed Folder    ${consolidated_report_path}
                Close Erp 

                # Run Keyword And Ignore Error    Close DMS
                Return From Keyword    ${consolid_report_path_in_processed_erp_dir}    ${carry_over_ratio_report_path}
            ELSE
                # Login To DMS  
                # Generate Jobcard Status Report
                Capture Screenshot
                Fail    Unable to download Jobcard Carry Over Ratio Report.
            END
        ELSE
            ERP_methods.Close Browser Processes
            Run Keyword And Ignore Error    Close Erp
            Fail    Unable to login to the DMS
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        # mail send code
        Capture Screenshot
        Fail    ${error_message} 
    END

 Copy Consolidated Report To Results Folder
    [Arguments]    ${consolid_report_path_in_processed_erp_dir}    ${curr_date}
    TRY
        ${project_root}     business_operations.get_process_root_directory
        ${results_folder}    Set Variable     ${project_root}\\${results_dir}\\${curr_date}
        
        ${consolidated_report_path_in_results_dir}    business_operations.copy_report_to_destination_folder    ${consolid_report_path_in_processed_erp_dir}    ${results_folder}
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message} 
    END
    [Return]    ${consolidated_report_path_in_results_dir}

Copy Dms Report To Results Downloads Folder    
    [Arguments]    ${carry_over_ratio_report_path}    ${curr_date}
    TRY
        ${results_download_folder}    Prepare Results Folder Path    ${curr_date} 

        move_report_to_destination_folder    ${carry_over_ratio_report_path}    ${results_download_folder}
        
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END


Close DMS
    Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${exit_btn_logout_image}     ${time_out}
    ${exit_btn_logout_image_exists}=    SikuliLibrary.Exists    ${exit_btn_logout_image}   
    IF    ${exit_btn_logout_image_exists} == True
        SikuliLibrary.Click    ${exit_btn_logout_image} 

        ${exeed_maximum_idle_time_popup_image_exists}=    SikuliLibrary.Exists    ${exeed_maximum_idle_time_popup_image}   
        IF    ${exeed_maximum_idle_time_popup_image_exists} == True
            SikuliLibrary.Click    ${exeed_maximum_idle_time_popup_image} 
            Sleep    ${normal_sleep}
            SikuliLibrary.Click    ${sessionout_ok_btn_image}            
        END  
    END
    ${exit_btn_logout_image_exists}=    SikuliLibrary.Exists    ${exit_btn_logout_image}   
    IF    ${exit_btn_logout_image_exists} == True
        SikuliLibrary.Click    ${exit_btn_logout_image} 
        Sleep    ${normal_sleep}
    END







    
  