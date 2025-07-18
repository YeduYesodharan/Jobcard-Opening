*** Settings ***
Library   SikuliLibrary  mode=OLD
Library   RPA.Desktop
Library   RPA.Tables
Library   RPA.Excel.Files
Library   OperatingSystem
library   RPA.Windows
Library   Dialogs
# Library   JSONLibrary
Library    Collections
Library    String
Variables  Variables/variables.py
Library    Libraries/business_operations.py
Library    Libraries/Post_request.py
Resource   Resources/dms_generate_jobcard_report.robot
Resource   Resources/utility.robot
Resource   Resources/dms_login.robot
Resource   Resources/dms_recall_check.robot


*** Variables ***
${url}                                   http://rpa.popularmaruti.com/store-jobcards-creation-data
# ${region_mapping_sheet}                  ${CURDIR}${/}..\\Mapping\\Location Mapping DMS ERP.xlsx
${region_mapping_sheet}    C:\\JobcardOpeningIntegrated\\Mapping\\Location Mapping DMS ERP.xlsx
${imagerootfolder}                       ${CURDIR}${/}..\\Locators
${log_folder}     ${CURDIR}${/}..\\Screenshot
${dms_captured_folder}                   ${CURDIR}${/}..\\DMS_captured
${transactions_menu_image}               ${imagerootfolder}\\transactions_menu.png
${jobcard_opening_menu_image}            ${imagerootfolder}\\jobcard_opening_menu.png
${opening_menu_image}                    ${imagerootfolder}\\opening_menu.png
${jobcard_number_txtbox_image}           ${imagerootfolder}\\jobcard_number_textbox.png
${jobcard_no_label_image}                ${imagerootfolder}\\jobcard_no_label.png

# Images used to extract vehicle details
${registration_number_label_image}           ${imagerootfolder}\\registration_number_label.png
${omr_label_image}                           ${imagerootfolder}\\omr_label.png
${service_type_label_image}                  ${imagerootfolder}\\service_type_label.png
${subservice_type_label_image}               ${imagerootfolder}\\sub_service_type_label.png
${service_adviser_label_image}               ${imagerootfolder}\\service_adviser_label.png
${vehicle_id_label_image}                    ${imagerootfolder}\\vehicle_id_label.png
${vehicle_model_label_image}                 ${imagerootfolder}\\vehicle_model_label.png
${vehicle_varient_label_image}               ${imagerootfolder}\\vehicle_varient_label.png
${color_label_image}                         ${imagerootfolder}\\color_label.png
${vehicle_sale_date_label_image}             ${imagerootfolder}\\vehicle_sale_date_label.png
${extended_warranty_label_image}             ${imagerootfolder}\\extended_warranty_label.png
${mcp_label_image}                           ${imagerootfolder}\\mcp_label.png
${technical_campaign_label_image}            ${imagerootfolder}\\technical_campaign_label.png
${promised_vehicle_datetime_label_image}     ${imagerootfolder}\\promised_vehicle_datetime_label.png

# Images used to extract customer details
${customer_id_label_image}                   ${imagerootfolder}\\customer_id_label.png
${customer_name_label_image}                 ${imagerootfolder}\\customer_name_label.png
${address_label_image}                       ${imagerootfolder}\\address_label.png
${city_label_image}                          ${imagerootfolder}\\city_label.png
${state_label_image}                         ${imagerootfolder}\\state_label.png
${phone_label_image}                         ${imagerootfolder}\\phone_label.png
${mobile_label_image}                        ${imagerootfolder}\\mobile_label.png

${exit_button_image}                        ${imagerootfolder}\\exit_button.png
${job_card_details_title_image}             ${imagerootfolder}\\job_card_details_title.png
${demand_code_button_image}                 ${imagerootfolder}\\demand_code_button.png
${drt_details_title_image}                  ${imagerootfolder}\\drt_details_title.png
${demand_code_error_popup_image}            ${imagerootfolder}\\demand_code_error_popup.png
${popup_ok_btn_image}                       ${imagerootfolder}\\popup_ok_btn.png
${service_menu_button_image}                ${imagerootfolder}\\service_menu_button.png
${back_btn_image}                           ${imagerootfolder}\\back_btn.png
${customer_demands_title_image}             ${imagerootfolder}\\customer_demands_title.png
${demand_code_heading_image}                ${imagerootfolder}\\demand_code_heading.png
${sl_no_field_image}                        ${imagerootfolder}\\sl_no_field_image.png
${demand_code_click_image}                  ${imagerootfolder}\\demand_code_click.png

${msr_app_download_popup_image}                   ${imagerootfolder}\\msr_app_download_popup.png
${msr_app_download_popup_okbtn_image}             ${imagerootfolder}\\msr_app_download_popup_okbtn.png
${service_followup_popup_image}                   ${imagerootfolder}\\service_followup_popup.png
${service_followup_okbtn_image}                   ${imagerootfolder}\\service_followup_okbtn.png
${exit_yes_button_image}                          ${imagerootfolder}\\exit_yes_button.png
${extended_warrenty_popup_image}                  ${imagerootfolder}\\extended_warrenty_popup.png
${extended_warrenty_popup_okbtn_image}            ${imagerootfolder}\\extended_warrenty_popup_okbtn.png
${ok_new_popup_image}                             ${imagerootfolder}\\ok_new_popup.png
${cvms_popup_ok_btn_image}                        ${imagerootfolder}\\cvms_popup_ok_btn.png
${valid_demand_code_enter_popup_image}            ${imagerootfolder}\\valid_demand_code_enter_popup.png
${demand_code_ok_btn_image}                       ${imagerootfolder}\\demand_code_ok_btn.png
${autorized_popup_image}                          ${imagerootfolder}\\autorized_popup.png
${auth_ok_btn_image}                              ${imagerootfolder}\\auth_ok_btn.png
${dms_full_image}                                 ${imagerootfolder}\\dms_full_image.png
${jobcard_details_title_input_jobcard_image}      ${imagerootfolder}\\jobcard_details_title_input_jobcard.png
${session_out_popup}                              ${imagerootfolder}\\session_out_popup.png
${session_out_popup1}                             ${imagerootfolder}\\session_out_popup1.png
${new_ccp_purchase_popup}                         ${imagerootfolder}\\new_ccp_purchase_popup.png
${new_ccp_purchase_ok_btn}                        ${imagerootfolder}\\new_ccp_purchase_ok_btn.png
${main_DMS_exit_btn}                              ${imagerootfolder}\\main_DMS_exit_btn.png
${demandcode_DMS3_new_popup}                      ${imagerootfolder}\\demandcode_DMS3_new_popup.png
${demandcode_DMS3_new_popup_ok}                   ${imagerootfolder}\\demandcode_DMS3_new_popup_ok.png
${DMS3_ccp_popup_ok}                              ${imagerootfolder}\\DMS3_ccp_popup_ok.png
${DMS3_ccp_popup}                                 ${imagerootfolder}\\DMS3_ccp_popup.png
${demand_code_error_popup_ok_dms3}                ${imagerootfolder}\\demand_code_error_popup_ok_dms3.png 
${DMS3_session_expire_popup}                      ${imagerootfolder}\\DMS3_session_expire_popup.png  
${DMS3_session_expire_popup_ok}                   ${imagerootfolder}\\DMS3_session_expire_popup_ok.png

${dms_execution_status_column_name}        DMS Execution Status
${exception_reason_column_name}        Exception Reason



*** Keywords ***
Jobcard Extraction Process
    [Arguments]    ${consolidated_report_path}    ${carry_over_ratio_report_path}    ${failure_jobcard_rows}    ${date_timestamp}
    TRY        
        
        IF    "${failure_jobcard_rows}" == "${EMPTY}"
            ${summary_data}    Read Summary File    ${consolidated_report_path}   
        ELSE
            ${summary_data}    Set Variable    ${failure_jobcard_rows}
        END    
        ${count}=    Set Variable    0

        FOR    ${index_val}    ${row}    IN ENUMERATE    @{summary_data}
            TRY

                ${jc_no}   Set Variable     ${row}[Job Card No]

                Update Execution Status As Null    ${consolidated_report_path}    ${jc_no}    ''    ${dms_execution_status_column_name} 
                Update Execution Status As Null    ${consolidated_report_path}    ${jc_no}    ''    ${exception_reason_column_name}  
                Update Execution Status As Null    ${consolidated_report_path}    ${jc_no}    ''    ${execution_status_column_name}

                # ----- Keyword used to switch to the DMS application------
                # Bring Window To Front    ${dms_title}

                #----------------------DMS3 Change-----------------------------#  
                # ${window_switched}    Run Keyword And Return Status    Bring Window To Front    ${dms_title}
                # IF    ${window_switched} == ${True}
                #     #---------------------------------------Session Expired DMS3 Popup handle--------------------------------------#
                #     Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${DMS3_session_expire_popup}    ${session_time_out}
                #     ${DMS3_session_expire_popup_exists}    SikuliLibrary.Exists    ${DMS3_session_expire_popup}
                #     IF    ${DMS3_session_expire_popup_exists} == ${True}
                #         SikuliLibrary.Click    ${DMS3_session_expire_popup_ok}
                #     END
                    #---------------------------------------Session Expired DMS3 Popup handle--------------------------------------#

                    #---------------------------------Added to exit Main DMS3 page-------------------------------------------------#
                    # Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${main_DMS_exit_btn}    ${time_out}
                    # ${main_DMS_exit_btn_exists}    SikuliLibrary.Exists    ${main_DMS_exit_btn}
                    # IF    ${main_DMS_exit_btn_exists} == ${True}
                    #     SikuliLibrary.Click    ${main_DMS_exit_btn}
                    # END
                    #--------------------------------Added to exit Main DMS3 page--------------------------------------------------#
                # END
                #----------------------DMS3 Change-----------------------------#  

                IF    ${index_val} > 0
                    Login To DMS
                END

                Capture Screenshot
                
                ${jobcard_opening_present_status}=   Run Keyword And Return Status     Select Jobcard Opening Menu
                Capture Screenshot
            
                IF    '${jobcard_opening_present_status}' == 'False'
                    Log To Console    ${jc_no} 
                    
                    ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired   
                    IF    ${session_exp_status} != ${True}            
                        Retry Login And Menu Select Module
                        Capture Screenshot
                    ELSE   
                        ${jobcard_opening_present_status}=   Run Keyword And Return Status     Select Jobcard Opening Menu
                        IF    '${jobcard_opening_present_status}' == 'False'
                           
                            ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired 
                            IF    ${session_exp_status} != ${True}            
                                Retry Login And Menu Select Module
                                Capture Screenshot
                            END
                            
                        END
                    END
                END

                ${run_continue}    Set Variable    ${False}
                Capture Screenshot
                ${jobcard_input_status}    Input Jobcard Number    ${jc_no}
                
                IF    ${jobcard_input_status} == ${True}
                    
                    ${extraction_status}    ${extracted_data_dict}    ${empty_keys}    Extract Fields    ${jc_no}

                    Log    ${extraction_status}
                    Log    ${extracted_data_dict}
                    
                    
                    IF    ${extraction_status} == ${False}
                        
                        ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired
                        
                        IF    ${session_exp_status} != ${True} 
                            Log    ${jc_no}                
                            Retry Login And Menu Select Module
                            ${jobcard_input_status}    Input Jobcard Number    ${jc_no}
                            
                            IF    ${jobcard_input_status} == ${True}
                                
                                ${extraction_status}    ${extracted_data_dict}    ${empty_keys}    Extract Fields    ${jc_no}
                                
                                IF    ${extraction_status} == ${False}
                                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Jobcard Extraction Failed.    ${exception_reason_column_name}         
                                    Capture Screenshot
                                    Exit Page
                                ELSE
                                    ${run_continue}    Set Variable    ${True}
                                END

                            ELSE
                                update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                                update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Job card number entry failed due to invisible field.    ${exception_reason_column_name}         
                                Capture Screenshot
                                Exit Page
                            END
                        ELSE
                            Exit Page
                            Select Jobcard Opening Menu
                            ${jobcard_input_status}     Input Jobcard Number    ${jc_no}
                            
                            IF    ${jobcard_input_status} == ${True}
                                
                                ${extraction_status}    ${extracted_data_dict}    ${empty_keys}    Extract Fields    ${jc_no}
                                
                                IF    ${extraction_status} == ${False}
                                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Jobcard Extraction Failed.    ${exception_reason_column_name}         
                                    Capture Screenshot
                                    Exit Page
                                ELSE
                                    ${run_continue}    Set Variable    ${True}
                                END
                            ELSE
                                update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                                update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Job card number entry failed due to invisible field.    ${exception_reason_column_name}         
                                Capture Screenshot
                                Exit Page
                            END
                        END
                    ELSE
                        ${run_continue}    Set Variable    ${True}
                    END
                    
                    IF    ${run_continue} == ${True}
                    
                        Update Extracted Data To Summary Excel    ${extracted_data_dict}    ${consolidated_report_path}    ${jc_no}    ${carry_over_ratio_report_path}     ${empty_keys}

                        ${demand_code_str}    Extract Demand Codes

                        IF    "${demand_code_str}" == "${False}"
                            
                            ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired
                            
                            IF    ${session_exp_status} != ${True}
                                Log    ${jc_no}  
                                Retry Login And Menu Select Module
                                ${jobcard_input_status}    Input Jobcard Number    ${jc_no} 

                                IF    ${jobcard_input_status} == ${True}

                                    ${demand_code_str}    Extract Demand Codes 
                                    ${status}    update_demand_codes_in_output_excel    ${consolidated_report_path}    ${jc_no}    ${demand_code_str}
                                ELSE
                                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Job card number entry failed due to invisible field.   ${exception_reason_column_name}         
                                    Capture Screenshot
                                    Exit Page
                                END
                            ELSE
                                ${demand_code_str}    Extract Demand Codes
                                ${status}    update_demand_codes_in_output_excel    ${consolidated_report_path}    ${jc_no}    ${demand_code_str}
                            
                            END            
                        
                        ELSE
                            ${status}    update_demand_codes_in_output_excel    ${consolidated_report_path}    ${jc_no}    ${demand_code_str}
                        END

                        IF    ${status} == ${True}
                            # update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    ${status}    ${dms_execution_status_column_name}
                            
                            ${recall_check_status}    ${extracted_data}    ${recall_status}    Check Recall Exist    ${jc_no}
                            
                            # IF    ${recall_check_status} == ${False}
                                # Retry Login And Menu Select Module
                                # Input Jobcard Number    ${jc_no}
                                
                                # ${recall_check_status}    ${extracted_data}    ${recall_status}    Check Recall Exist    ${jc_no}
                                
                            # IF    ${recall_check_status} == ${False}
                                # update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name}
                                # update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Recall Checking Failed Due to Some Reason. Retried, But Something Went Wrong.    ${exception_reason_column_name}
                                # update_recall_in_output_excel    ${consolidated_report_path}      ${jc_no}      ${extracted_data}    ${recall_status}
                            # ELSE
                            update_recall_in_output_excel    ${consolidated_report_path}      ${jc_no}      ${extracted_data}    ${recall_status}

                            IF    "${empty_keys}" == "${EMPTY}"
                                update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Success    ${dms_execution_status_column_name}
                                update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    ${EMPTY}    ${exception_reason_column_name}
                            END
                                    
                    
                            # END
                            # ELSE
                            #     update_recall_in_output_excel    ${consolidated_report_path}      ${jc_no}      ${extracted_data}    ${recall_status}
                            # END
                        END           

                        Exit Page
                    ELSE
                        update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                        update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Job card number entry failed due to invisible field.    ${exception_reason_column_name}         
                        Capture Screenshot
                        Exit Page
                    END
                
                ELSE
                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Job card number entry failed due to invisible field.    ${exception_reason_column_name}         
                    Capture Screenshot
                    Exit Page
                END

                # ----- Need to check if consolidated report contains dms execution status is success and recall status is yes rows ------
                ${continue_erp_status_with_recall}    check_jobcard_status    ${consolidated_report_path}    ${jc_no}    Yes

                IF   ${continue_erp_status_with_recall} == $True
                    #-----Keyword used to copy consolidated_report from processed directory to another central machine E:sftp/test directory for Recall process ----
                    #-----Argument: current_date and consolid_report_path_in_Inprogress_directory, config_sheet, jc number -----
                    consolidated_report_copy_to_central_machine    ${curr_date}    ${consolidated_report_path}    ${sheet_path}    ${jc_no} 
                END


                # ----- Need to check if consolidated report contains dms execution status is success and recall status is no rows ------
                ${continue_erp_status}    check_jobcard_status    ${consolidated_report_path}    ${jc_no}    No
                IF   ${continue_erp_status} == $True
                    
                    # ----- Keyword used to switch to the ERP application------     
                    # -----------------------------------DMS3 Change------------------------------#                 
                    Bring Window To Front    ${erp_title}
                    # Window Navigation    ${erp_title}    
                    # -----------------------------------DMS3 Change------------------------------#      
                    Capture Screenshot
                    
                    #------ Keyword used to enter jobcard details to ERP------
                    ${consolidated_report_path}    ERP_methods.Extract And Correct Registration    ${consolidated_report_path}
                    
                    ${jobcard_data_entry_status}    ${erp_message}   Run Keyword And Ignore Error    Jobcard Data Entry    ${consolidated_report_path}     ${jc_no}   
                    
                    Capture Screenshot

                    # Python function used to insert single row status to DB through API
                    # bot_run_status_save_to_db    ${url}    ${consolidated_report_path}   ${region_mapping_sheet}    ${jc_no}    ${date_timestamp}
                    # -----------------------------------DMS3 Change------------------------------#   
                    ${db_save_status}    bot_run_status_save_to_db    ${url}    ${consolidated_report_path}   ${region_mapping_sheet}    ${jc_no}    ${date_timestamp}            
                    IF    """${db_save_status}""" == """success"""
                        Continue For Loop
                    END
                    # -----------------------------------DMS3 Change------------------------------#               
                END
                
            EXCEPT    AS    ${error_message}
                Log    ${error_message}
                Capture Screenshot
                IF    """${error_message}""" == """Window with partial title 'Wings ERP 23E - Web Client - PRO' not found or could not be activated."""
                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${execution_status_column_name} 
                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    ${error_message}    ${exception_reason_column_name}

                ELSE IF    """${error_message}""" == """Window with partial title 'Dealer Management System' not found or could not be activated."""
                    ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired
                        
                    IF    ${session_exp_status} != ${True}
                        Login To DMS
                    END                    
                ELSE
                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    Fail    ${dms_execution_status_column_name} 
                    update_execution_status_in_summary_report    ${consolidated_report_path}    ${jc_no}    ${error_message}    ${exception_reason_column_name}         
                    ${exit_not_found}   Run Keyword And Return Status    Exit Page
                    IF    ${exit_not_found} == False
                        ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired
                        
                        IF    ${session_exp_status} != ${True}
                            Login To DMS
                        END
                    END
                END
                # Python function used to insert single row status to DB through API
                ${db_save_status}    bot_run_status_save_to_db    ${url}    ${consolidated_report_path}   ${region_mapping_sheet}    ${jc_no}    ${date_timestamp}
                # -----------------------------------DMS3 Change------------------------------#               
                IF    """${db_save_status}""" == """success"""
                    Continue For Loop
                END
                # Continue For Loop
                # -----------------------------------DMS3 Change------------------------------#  
                # ${count}=    Evaluate    ${count} + 1
                # Exit For Loop If    ${count} == 2
                             
            END
            # ${count}=    Evaluate    ${count} + 1
            # Exit For Loop If    ${count} == 2
        END
    EXCEPT    AS    ${error_message}
        Log    ${error_message}
        Capture Screenshot
        Exit Page
        Send Email Output     ${curr_date}    ${report_name}    ${error_message}    ${consolidated_report_path}       
    END

Exit Page 
    TRY
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${exit_button_image}    ${time_out}
        ${exit_button_image_exist}    SikuliLibrary.Exists    ${exit_button_image}
        IF    ${exit_button_image_exist} == ${False}
            
            ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired
                        
                IF    ${session_exp_status} != ${True}
                    Login To DMS
                ELSE
                    Fail    ${EMPTY}
                END    
        ELSE
            SikuliLibrary.Click    ${exit_button_image}  

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${exit_yes_button_image}    ${time_out}
            ${exit_yes_button_image_exist}    SikuliLibrary.Exists    ${exit_yes_button_image}
            
            IF    ${exit_yes_button_image_exist} == ${True}
                SikuliLibrary.Click    ${exit_yes_button_image}
            END
            #---------------------------------Added to exit Jobcard data DMS3 page----------------------------------------#
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${main_DMS_exit_btn}    ${time_out}
            ${main_DMS_exit_btn_exists}    SikuliLibrary.Exists    ${main_DMS_exit_btn}
            IF    ${main_DMS_exit_btn_exists} == ${True}
                SikuliLibrary.Click    ${main_DMS_exit_btn}
            END
            #--------------------------------Added to exit Jobcard data DMS3 page------------------------------------------#
            #---------------------------------Added to exit Main DMS3 page-------------------------------------------------#
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${main_DMS_exit_btn}    ${session_time_out}
            ${main_DMS_exit_btn_exists}    SikuliLibrary.Exists    ${main_DMS_exit_btn}
            IF    ${main_DMS_exit_btn_exists} == ${True}
                Run Keyword And Ignore Error    SikuliLibrary.Click    ${main_DMS_exit_btn}
            END
            #--------------------------------Added to exit Main DMS3 page--------------------------------------------------#
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    
Retry Login And Menu Select Module
    Login To DMS
    Select Jobcard Opening Menu
    
Check For Session Expired
    TRY
        # SikuliLibrary.Wait Until Screen Contain    ${service_menu_image}     ${time_out}
        # ${service_menu_exists}=    SikuliLibrary.Exists    ${service_menu_image}   
        # IF    ${service_menu_exists}==False
            # Show Message Popup    Alert    Service Menu Not Found. Session Expired. Retry Login
            # Fail    Service Menu Not Found. Session Expired. Retry Login
        # ELSE
            # ${session_out_popup_exists}=    SikuliLibrary.Exists    ${session_out_popup} 
            # IF    ${service_menu_exists}==$True
            #     Fail    Session Expired. Retry Login
            # ELSE
            #     return from keyword    True
            # END
        # END

        ${session_out_popup_exists}=    SikuliLibrary.Exists    ${session_out_popup}
        IF    ${session_out_popup_exists} == $True
            Fail    Session Expired. Retry Login
        ELSE
            return from keyword    True
        END

    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END

Select Jobcard Opening Menu
    TRY
        Capture Screenshot
        #------------------------------------------DMS V3 Sikuli Change-----------------------#

        # SikuliLibrary.Wait Until Screen Contain    ${service_menu_image}     ${time_out}
        # SikuliLibrary.Click    ${service_menu_image}

        SikuliLibrary.Wait Until Screen Contain    ${service_btn_v3}     ${time_out}
        SikuliLibrary.Click    ${service_btn_v3}

        #------------------------------------------DMS V3 Sikuli Change-----------------------#

        SikuliLibrary.Wait Until Screen Contain    ${transactions_menu_image}     ${time_out}
        SikuliLibrary.Click    ${transactions_menu_image}
        Capture Screenshot
        SikuliLibrary.Wait Until Screen Contain    ${jobcard_opening_menu_image}     ${time_out}
        SikuliLibrary.Click    ${jobcard_opening_menu_image}

        SikuliLibrary.Wait Until Screen Contain    ${opening_menu_image}     ${time_out}
        SikuliLibrary.Click    ${opening_menu_image}
        Capture Screenshot
        ${session_exp_status}    Run Keyword And Return Status    Check For Session Expired   
        IF    ${session_exp_status} == ${True}
            Return From Keyword   ${True}
        ELSE
            Return From Keyword   ${False} 
        END
        Capture Screenshot

    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END

Check For Authorized Popup And Close
    Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${autorized_popup_image}     ${min_timout}
    ${autorized_popup_image_exists}=    SikuliLibrary.Exists    ${autorized_popup_image}   
    IF    ${autorized_popup_image_exists} == True
        SikuliLibrary.Click    ${auth_ok_btn_image}
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${jobcard_no_label_image}     ${min_timout}
    END


Input Jobcard Number
    [Arguments]    ${jobcard_no}
    TRY
        Check For Authorized Popup And Close

        ${jobcard_no_visible_status}    Capture Jobcard Number Field     ${jobcard_no_label_image}

        IF    ${jobcard_no_visible_status} == ${True}

            SikuliLibrary.Input text   ${jobcard_number_txtbox_image}     ${jobcard_no}

            RPA.Desktop.Press Keys  enter

            ${running_status}    ${entered_jc_no}    Custom Get Text From Image    ${jobcard_no_label_image}    80    0    45    2

            IF    "${entered_jc_no}" == ""
                ${jobcard_no_visible_status}    Capture Jobcard Number Field     ${jobcard_no_label_image}
                SikuliLibrary.Input text   ${jobcard_number_txtbox_image}     ${jobcard_no}
                RPA.Desktop.Press Keys  enter   

                ${running_status}    ${entered_jc_no}    Custom Get Text From Image    ${jobcard_no_label_image}    80    0    45    2
                IF    "${entered_jc_no}" == ""
                    Return From Keyword    ${False}
                END
            END

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${cvms_popup_ok_btn_image}     ${min_wait_time}
            ${cvms_popup_ok_btn_image_exists}=    SikuliLibrary.Exists    ${cvms_popup_ok_btn_image}   
            Capture Screenshot
            IF    ${cvms_popup_ok_btn_image_exists}==True
                Run Keyword And Ignore Error    SikuliLibrary.Click    ${cvms_popup_ok_btn_image} 
            END

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${extended_warrenty_popup_image}     ${min_wait_time}
            ${extended_warrenty_popup_image_exists}=    SikuliLibrary.Exists    ${extended_warrenty_popup_image}   
            Capture Screenshot
            IF    ${extended_warrenty_popup_image_exists}==True
                Run Keyword And Ignore Error    SikuliLibrary.Click    ${extended_warrenty_popup_okbtn_image} 
            END
            
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${msr_app_download_popup_image}     ${min_wait_time}
            ${msr_app_download_popup_image_exists}=    SikuliLibrary.Exists    ${msr_app_download_popup_image}   
            Capture Screenshot
            IF    ${msr_app_download_popup_image_exists}==True
                Run Keyword And Ignore Error    SikuliLibrary.Click    ${msr_app_download_popup_okbtn_image} 
            END

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${service_followup_popup_image}     ${min_wait_time}
            ${service_followup_popup_image_exists}=    SikuliLibrary.Exists    ${service_followup_popup_image} 
            Capture Screenshot  
            IF    ${service_followup_popup_image_exists}==True
                Run Keyword And Ignore Error    SikuliLibrary.Click    ${service_followup_okbtn_image} 
            END


            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${new_ccp_purchase_popup}     ${min_wait_time}
            ${new_ccp_purchase_popup_image_exists}=    SikuliLibrary.Exists    ${new_ccp_purchase_popup} 
            Capture Screenshot  
            IF    ${new_ccp_purchase_popup_image_exists}==True
                Run Keyword And Ignore Error    SikuliLibrary.Click    ${new_ccp_purchase_ok_btn} 
            END
            
            #--------------------------------------------DMS3 new ccp Popup-----------------------------------------#
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${DMS3_ccp_popup}     ${min_wait_time}
            ${DMS3_ccp_popup_exists}=    SikuliLibrary.Exists    ${DMS3_ccp_popup} 
            Capture Screenshot  
            IF    ${DMS3_ccp_popup_exists}==True
                Run Keyword And Ignore Error    SikuliLibrary.Click    ${DMS3_ccp_popup_ok} 
            END
            #--------------------------------------------DMS3 new ccp Popup-----------------------------------------#
            Sleep    ${avg_time}
             # Fetching vehicle Id. Passing the captured image of vehicle id label.
            # ${running_status}    ${extracted_value}    Get Vehicle Id From Image    ${vehicle_id_label_image}    "63"    "0"    "90"    "3"
            # IF    ${running_status} == ${True}

                # IF     ${extracted_value} != ${EMPTY}
                    # Log    Success
                    # Return From Keyword    ${True}
                # ELSE   
                    # Return From Keyword    ${False}
                # END 
            
            # ELSE   
            #     Return From Keyword    ${False}
            # END 
            Return From Keyword    ${True}
        ELSE
            Return From Keyword    ${False}
        END

        # END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END

Capture Jobcard Number Field
    [Arguments]     ${jobcard_no_label_image}
    TRY
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${jobcard_no_label_image}   ${time_out}

        ${jobcard_no_label_image_exists}=    SikuliLibrary.Exists    ${jobcard_no_label_image}   
        IF    ${jobcard_no_label_image_exists}==False
            Return from keyword    ${False}
        ELSE     
            ${label_coordinates}=    SikuliLibrary.Get Image Coordinates    ${jobcard_no_label_image}
            Log    ${label_coordinates}

            ${textfield_value_coordinates}=   increase_cordinates   ${label_coordinates}    80    0    45    2
            
            ${new_textfield_image}=    SikuliLibrary.Capture Region    ${textfield_value_coordinates}
            Sleep    3s
            SikuliLibrary.Click    ${new_textfield_image}

            Return from keyword    ${True}
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
        

Extract Fields
    [Arguments]    ${jobcard_no}
    TRY
    
        Capture Screenshot

        ${extracted_data}=    Create Dictionary

        Set To Dictionary    ${extracted_data}    Job Card No=${jobcard_no}

        ${dms_full_image_coordinates}=    SikuliLibrary.Get Image Coordinates    ${dms_full_image}
        Log    ${dms_full_image_coordinates}
        ${captured_path}=    SikuliLibrary.Capture Region    ${dms_full_image_coordinates}
        ${new_image_path}    Set Variable    ${dms_captured_folder}\\jobcard_image_${jobcard_no}.png

        # Move the captured image to the new location
        OperatingSystem.Move File    ${captured_path}    ${new_image_path}
        ${output}    extract_image_data    ${new_image_path}
        Log    ${output}

        ${omr_code}=    Get From Dictionary    ${output}    omr
        Set To Dictionary    ${extracted_data}    OMR=${omr_code}
        
        #---------- Vehicle Details -----------------------------------

        # ${service_type_id}=    Get From Dictionary    ${output}    service_type_id
        # Set To Dictionary    ${extracted_data}    Service Type Code=${service_type_id}


        
        ${sub_service_type}=    Get From Dictionary    ${output}    sub_service_type
        Set To Dictionary    ${extracted_data}    Sub Service Type=${sub_service_type}


        
        ${service_advisor_code}=    Get From Dictionary    ${output}    service_advisor_code
        Set To Dictionary    ${extracted_data}    Service Advisor Code=${service_advisor_code}


        # Fetching vehicle Id. Passing the captured image of vehicle id label.
        ${running_status}    ${extracted_value}    Get Vehicle Id From Image    ${vehicle_id_label_image}    "63"    "0"    "90"    "3"
        IF    ${running_status} == ${False}
            Return From Keyword    ${False}    ${extracted_data}
        ELSE 
            # ${vehicle_id}    remove_space    ${extracted_value}
            Set To Dictionary    ${extracted_data}    Vehicle ID=${extracted_value} 
        END
       

        ${vehicle_model_id}=    Get From Dictionary    ${output}    vehicle_model_id
        Set To Dictionary    ${extracted_data}    Vehicle Model=${vehicle_model_id}

        
        ${vehicle_variant_id}=    Get From Dictionary    ${output}    vehicle_variant_id
        Set To Dictionary    ${extracted_data}    Vehicle Variant=${vehicle_variant_id}

       
        ${color_id}=    Get From Dictionary    ${output}    color_id
        Set To Dictionary    ${extracted_data}    Color=${color_id}


        #Fetching vehicle sale date. Passing the captured image of vehicle sale date label.
        ${running_status}    ${extracted_value}    Custom Get Text From Image    ${vehicle_sale_date_label_image}    "105"    "0"    "13"    "2"
        IF    ${running_status} == ${False}
            Return From Keyword    ${False}    ${extracted_data}
        ELSE
            # Set To Dictionary    ${extracted_data}    Vehicle Sales Date=${extracted_value}
            #------------------------------------------DMS3 change-------------------------------------#
            ${clean_value}=    Replace String Using Regexp    ${extracted_value}    [^0-9\-]    ""
            Set To Dictionary    ${extracted_data}    Vehicle Sales Date=${clean_value}
            #------------------------------------------DMS3 change-------------------------------------#
        END
        
        
        ${is_extended_warranty}=    Get From Dictionary    ${output}    is_extended_warranty
        Set To Dictionary    ${extracted_data}    Extended Warranty=${is_extended_warranty}


        
        ${is_mcp}=    Get From Dictionary    ${output}    is_mcp
        Set To Dictionary    ${extracted_data}    MCP=${is_mcp}


       
        ${technical_campaing}=    Get From Dictionary    ${output}    technical_campaing
        Set To Dictionary    ${extracted_data}    Technical Campaign=${technical_campaing}
       
        
        #---------------- Customer Details -------------------------------
        #Fetching Customer ID. Passing the captured image of Customer ID label.
        ${running_status}    ${extracted_value}    Custom Get Text From Image    ${customer_id_label_image}    "76"    "-6"    "49"    "4"
        IF    ${running_status} == ${False}
            Return From Keyword    ${False}    ${extracted_data}
        ELSE
            Set To Dictionary    ${extracted_data}    Customer ID=${extracted_value}
        END

        #Fetching city. Passing the captured image of city label.
        ${running_status}    ${extracted_value}    Custom Get Text From Image    ${city_label_image}    "73"    "0"    "119"    "0"
        IF    ${running_status} == ${False}
            Return From Keyword    ${False}    ${extracted_data}
        ELSE
            Set To Dictionary    ${extracted_data}    City=${extracted_value}
        END

        #Fetching state. Passing the captured image of state label.
        ${running_status}    ${extracted_value}    Custom Get Text From Image    ${state_label_image}    "36"    "0"    "100"    "0"
        IF    ${running_status} == ${False}
            Return From Keyword    ${False}    ${extracted_data}
        ELSE
            Set To Dictionary    ${extracted_data}    State=${extracted_value}
        END

        #Fetching phone. Passing the captured image of phone label.
        ${running_status}    ${extracted_value}    Custom Get Text From Image    ${phone_label_image}    "43"    "-2"    "135"    "1"
        IF    ${running_status} == ${False}
            Return From Keyword    ${False}    ${extracted_data}
        ELSE
            Set To Dictionary    ${extracted_data}    Phone=${extracted_value}
        END

        #Fetching mobile. Passing the captured image of mobile label.
        ${running_status}    ${extracted_value}    Custom Get Text From Image    ${mobile_label_image}    "43"    "0"    "90"    "2"
        IF    ${running_status} == ${False}
            Return From Keyword    ${False}    ${extracted_data}
        ELSE
        #-------------------------------------added for DMS 3--------------------------------#
            ${clean_mobile_data}=    Extract Valid Mobile    ${extracted_value}
            Set To Dictionary    ${extracted_data}    Mobile=${clean_mobile_data}
        #-------------------------------------added for DMS 3--------------------------------#    
            
            # Set To Dictionary    ${extracted_data}    Mobile=${extracted_value}
        END

        ${empty_keys}    Check For Empty Data Extracted     ${extracted_data}
        Log    ${extracted_data}

    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    Return From Keyword    ${True}    ${extracted_data}     ${empty_keys}


Get Coordinate
    [Arguments]    ${image}
    ${test_co}=    SikuliLibrary.Get Image Coordinates    ${image}
    Log    ${test_co}


Get Demand Code
    [Arguments]      ${form_label_image}    ${x}    ${y}    ${w}    ${h} 
    TRY 

        ${running_status}    Set Variable    ${True}
        ${extracted_value}    Set Variable    ${EMPTY}

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${form_label_image}   ${time_out}

        ${field_label_image_exists}=    SikuliLibrary.Exists    ${form_label_image}   
        IF    ${field_label_image_exists}==False
            # Messagebox.Show Message Popup    Alert    Srl No Not Found.
            ${running_status}    Set Variable    ${False}
        ELSE
                
            ${label_coordinates}=    SikuliLibrary.Get Image Coordinates    ${form_label_image}
            Log    ${label_coordinates}

            ${textfield_value_coordinates}=   increase_cordinates   ${label_coordinates}    ${x}    ${y}    ${w}    ${h}
            
            # ${new_textfield_value_image}=    SikuliLibrary.Capture Region    ${textfield_value_coordinates}

            SikuliLibrary.Double Click On Region    ${textfield_value_coordinates}
            
            ${extracted_value}     Copy And Get Text From Clipboard
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    [Return]    ${running_status}    ${extracted_value}
    

Copy And Get Text From Clipboard
    RPA.Desktop.Press keys    ctrl    a

    RPA.Desktop.Press keys    ctrl    c
    Sleep    0.5s

    ${extracted_value}    Get Clipboard Value
    Log    ${extracted_value}

    Return From Keyword    ${extracted_value}

Extract Demand Codes
    TRY
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${job_card_details_title_image}     ${time_out}
        ${job_card_details_title_exists}=    SikuliLibrary.Exists    ${job_card_details_title_image}   
        IF    ${job_card_details_title_exists}==False
            # Messagebox.Show Message Popup    Alert    Job Card Details Title Not Found In The Existing Page. Need To Login Again.
            Return From Keyword    ${False}
        ELSE
            SikuliLibrary.Wheel Down    5    ${drt_details_title_image}

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${demand_code_button_image}     ${time_out}
            ${demand_code_button_image_exists}=    SikuliLibrary.Exists    ${demand_code_button_image}   
            IF    ${demand_code_button_image_exists}==False
                
                SikuliLibrary.Wheel Down    5    ${drt_details_title_image}

                SikuliLibrary.Click    ${demand_code_button_image}
            ELSE
                SikuliLibrary.Click    ${demand_code_button_image} 

                # Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${ok_new_popup_image}     ${minimum_time_out}
                # ${ok_new_popup_image_exists}=    SikuliLibrary.Exists    ${ok_new_popup_image}   
                # IF    ${ok_new_popup_image_exists}==True
                #     SikuliLibrary.Click    ${ok_new_popup_image} 
                # END
            END
            Sleep    ${normal_sleep}
            ${demand_code_error_popup_image_exists}=    SikuliLibrary.Exists    ${demand_code_error_popup_image}
            
            IF    ${demand_code_error_popup_image_exists}==False
                ${demand_code_str}     Extract And Update Demand Codes
                Return From Keyword    ${demand_code_str}
            ELSE
                #----------------------------------changed for DMS3-----------------------------------------#
                # ${popup_ok_btn_image_exists}=    SikuliLibrary.Exists    ${popup_ok_btn_image}
                ${popup_ok_btn_image_exists}=    SikuliLibrary.Exists    ${demand_code_error_popup_ok_dms3}
                IF    ${popup_ok_btn_image_exists}==True
                    SikuliLibrary.Click    ${demand_code_error_popup_ok_dms3}
                END
                #----------------------------------changed for DMS3-----------------------------------------#

                Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${service_menu_button_image}     ${time_out}
                ${service_menu_button_image_exists}=    SikuliLibrary.Exists    ${service_menu_button_image}
                IF    ${service_menu_button_image_exists}==True
                    SikuliLibrary.Click    ${service_menu_button_image}
                END

                Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${back_btn_image}     ${time_out}
                ${back_btn_image_exists}=    SikuliLibrary.Exists    ${back_btn_image}
                IF    ${back_btn_image_exists}==True
                    SikuliLibrary.Click    ${back_btn_image}
                END

                #----------------------------------changed for DMS3-----------------------------------------#
                Run Keyword And Ignore Error    SikuliLibrary.Wheel Down    5    ${drt_details_title_image}
                #----------------------------------changed for DMS3-----------------------------------------#

                Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${demand_code_button_image}     ${time_out}
                ${demand_code_button_image_exists}=    SikuliLibrary.Exists    ${demand_code_button_image}   
                IF    ${demand_code_button_image_exists}==False
                    Return From Keyword    ${False}                
                ELSE
                    
                    SikuliLibrary.Click    ${demand_code_button_image}

                    Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${ok_new_popup_image}     ${Min_time}
                    ${ok_new_popup_image_exists}=    SikuliLibrary.Exists    ${ok_new_popup_image}   
                    IF    ${ok_new_popup_image_exists}==True
                        SikuliLibrary.Click    ${ok_new_popup_image} 
                    END

                    SikuliLibrary.Wheel Up    5

                    Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${customer_demands_title_image}     ${time_out}
                    ${customer_demands_title_image_exists}=    SikuliLibrary.Exists    ${customer_demands_title_image}   
                    IF    ${customer_demands_title_image_exists}==False
                        Return From Keyword    ${False} 
                    END
                END
                ${demand_code_str}     Extract And Update Demand Codes
                Return From Keyword    ${demand_code_str}
                END
        END 
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    
Extract And Update Demand Codes
    TRY
        
        SikuliLibrary.Wheel Up    5

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${demand_code_heading_image}     ${time_out}
        ${demand_code_heading_image_exists}=    SikuliLibrary.Exists    ${demand_code_heading_image}   
        IF    ${demand_code_heading_image_exists}==False
            Return From Keyword    ${False} 
        ELSE
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${demand_code_click_image}     ${time_out}
            # SikuliLibrary.Click    ${demand_code_click_image}

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${valid_demand_code_enter_popup_image}     ${min_wait_time}
            ${valid_demand_code_enter_popup_image_exists}=    SikuliLibrary.Exists    ${valid_demand_code_enter_popup_image}   
            IF    ${valid_demand_code_enter_popup_image_exists}==True
                SikuliLibrary.Click    ${demand_code_ok_btn_image}
            END          

            ${demand_code_str}   Set Variable     ${EMPTY}
            # ${demand_code_list}   Create List

            #Fetching Demand Codes Passing the captured image of Demand Code label.
            # ${running_status}    ${extracted_value}    Get Demand Code    ${sl_no_field_image}    "47"    "0"    "23"    "0"
            Clear Clipboard
            ${extracted_value}     Copy And Get Text From Clipboard

            IF    '${extracted_value}' == ''
                Return From Keyword    ${False}
                # Return From Keyword    ${False}    ${extracted_data}
            ELSE
                WHILE    '${extracted_value}' != '${EMPTY}'
                    # Append To List    ${demand_code_list}    ${extracted_value}
                    IF    '${demand_code_str}' == '${EMPTY}'
                        ${demand_code_str}    Set Variable    ${extracted_value}
                    ELSE
                        ${demand_code_str}    Set Variable    ${demand_code_str},${extracted_value}
                    END
                    # Set To Dictionary    ${extracted_data}    Mobile=${extracted_value}
                    RPA.Desktop.Press Keys    Down

                    
                    #------------------------------------------added for DMS Sikuli new popup-----------------------------------------------#
                    Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${demandcode_DMS3_new_popup}     ${min_wait_time}
                    ${demandcode_DMS3_new_popup_exists}=    SikuliLibrary.Exists    ${demandcode_DMS3_new_popup}   
                    IF    ${demandcode_DMS3_new_popup_exists}==True
                        # Press Keys Action    enter
                        SikuliLibrary.Click    ${demandcode_DMS3_new_popup_ok}
                        # Return From Keyword    ${False}
                    END

                    #------------------------------------------added for DMS Sikuli new popup-----------------------------------------------#



                    ${valid_demand_code_enter_popup_image_exists}=    SikuliLibrary.Exists    ${valid_demand_code_enter_popup_image}   
                    IF    ${valid_demand_code_enter_popup_image_exists}==True
                        SikuliLibrary.Click    ${demand_code_ok_btn_image}
                    END 
                    Clear Clipboard
                    ${extracted_value}     Copy And Get Text From Clipboard
                END 
            END
            Return From Keyword    ${demand_code_str}
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    
    
    
Update Extracted Data To Summary Excel
    [Arguments]    ${extracted_data_dict}    ${summary_file}    ${jobcard_no}    ${carry_over_ratio_report_path}     ${empty_keys}
    TRY
        update_output_excel_with_extracted_values    ${extracted_data_dict}    ${summary_file}    ${jobcard_no}    ${carry_over_ratio_report_path}     ${empty_keys}
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
        
# Reusable method used to extract text from each fields in a Jobcard details page.
# Here the argument is the image of each field.
# After process the method returns curresponding text from the image. 
Custom Get Text From Image
    [Arguments]      ${form_label_image}    ${x}    ${y}    ${w}    ${h}  
    TRY

        ${running_status}    Set Variable    ${True}
        ${extracted_value}    Set Variable    ${EMPTY}

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${form_label_image}   ${minimum_time_out}

        ${field_label_image_exists}=    SikuliLibrary.Exists    ${form_label_image}   
        IF    ${field_label_image_exists}==False
            ${running_status}    Set Variable    ${False}
        ELSE
            # Custom Get Text From Image     
            ${label_coordinates}=    SikuliLibrary.Get Image Coordinates    ${form_label_image}
            Log    ${label_coordinates}

            ${textfield_value_coordinates}=   increase_cordinates   ${label_coordinates}    ${x}    ${y}    ${w}    ${h}
            
            ${new_textfield_image}=    SikuliLibrary.Capture Region    ${textfield_value_coordinates}

            ${extracted_value}=    SikuliLibrary.Get Text    ${new_textfield_image}
            Log    ${extracted_value}
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    [Return]    ${running_status}    ${extracted_value}
    
Get Vehicle Id From Image
    [Arguments]      ${form_label_image}    ${x}    ${y}    ${w}    ${h}  
    TRY

        ${running_status}    Set Variable    ${True}
        ${extracted_value}    Set Variable    ${EMPTY}

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${form_label_image}   ${time_out}

        ${field_label_image_exists}=    SikuliLibrary.Exists    ${form_label_image}   
        IF    ${field_label_image_exists}==False
            # Messagebox.Show Message Popup    Alert    Jobcard Number Textbox Not Found.
            ${running_status}    Set Variable    ${False}
        ELSE
            # Custom Get Text From Image     
            ${label_coordinates}=    SikuliLibrary.Get Image Coordinates    ${form_label_image}
            Log    ${label_coordinates}

            ${textfield_value_coordinates}=   increase_cordinates   ${label_coordinates}    ${x}    ${y}    ${w}    ${h}
            
            ${new_textfield_image}=    SikuliLibrary.Capture Region    ${textfield_value_coordinates}

            # SikuliLibrary.Click    ${new_textfield_image}
            SikuliLibrary.Double Click    ${new_textfield_image}
            
            Clear Clipboard
            ${extracted_value}     Copy And Get Text From Clipboard

            ${length}=    Get Length    ${extracted_value}
            IF    '${extracted_value}' == '' or ${length} != 17
                Clear Clipboard
                ${extracted_value}     Copy And Get Text From Clipboard

                ${length}=    Get Length    ${extracted_value}
                IF    '${extracted_value}' == '' or ${length} != 17
                    # SikuliLibrary.Click    ${new_textfield_image}
                    Clear Clipboard
                    ${extracted_value}     Copy And Get Text From Clipboard
                    
                    # ${length}=    Get Length    ${extracted_value}
                    # IF    '${extracted_value}' == '' or ${length} != 17
                        # Fail    Vehicle ID format is not standard
                    # END

                END

            END

            Log    ${extracted_value}
        END
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END
    [Return]    ${running_status}    ${extracted_value}

Prepare Failure Report Path
    [Arguments]    ${curr_date}
    TRY
        ${project_root}    business_operations.get_process_root_directory
        ${failure_report_path}    Set Variable     ${project_root}\\${process_related_folder_name}\\${curr_date}\\${inprogress_dir_name}\\consolidated_failed_report.xlsx
        Log    ${failure_report_path} 
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END  
    [Return]    ${failure_report_path}

Move Consolidated Report To Processed Folder
    [Arguments]    ${consolidated_report_path}

    ${project_root}     business_operations.get_process_root_directory
    ${processed_erp_dir_path}    Set Variable     ${project_root}\\${process_related_folder_name}\\${curr_date}\\${processed_erp_dir_name}
    
    move_report_to_destination_folder    ${consolidated_report_path}    ${processed_erp_dir_path}

    ${file_name}    Get File Name    ${consolidated_report_path}
    Log To Console    ${file_name}

    ${consolid_report_path_in_processed_erp_dir}    Set Variable     ${project_root}\\${process_related_folder_name}\\${curr_date}\\${processed_erp_dir_name}\\${file_name}
    Return From Keyword    ${consolid_report_path_in_processed_erp_dir}

Prepare Results Folder Path
    [Arguments]    ${curr_date}    
    TRY
        ${project_root}     business_operations.get_process_root_directory
        ${results_download_folder}    Set Variable     ${project_root}\\${results_dir}\\${curr_date}\\${downloads_dir_name}
        Log    ${results_download_folder} 
    EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message} 
    END  
    [Return]    ${results_download_folder}


# *** Tasks ***
# Demo
#     ${curr_date}    ${date_timestamp}    get_current_date_fun
# #     Send Email Output     ${curr_date}    Job Card Creation Process    ${EMPTY}    E:\\JobcardOpeningIntegrated\\ProcessRelatedFolders\\2025-04-27\\ProcessedERP\\consolidated_jobcard_report20250427170351.xlsx 
#     Set Global Variable    ${curr_date}
# #     # ERP Login
#     # Login To DMS
# #     bring_to_foreground    Wings ERP 23E - Web Client - PRO
#     ${time_out}    Set Variable    45
#     Set Global Variable    ${time_out}

# #     ${full_title}    get_the_full_title    Wings ERP 23E
# #     Control Window    name:"${full_title}"

#     ${consolidated_report_path}    Set Variable         E:\\JobcardOpeningIntegrated\\ProcessRelatedFolders\\2025-05-02\\InProgress\\consolidated_jobcard_report20250502203957.xlsx
#     ${carry_over_ratio_report_path}    Set Variable         E:\\JobcardOpeningIntegrated\\ProcessRelatedFolders\\2025-05-02\\Downloads\\ELM.FOB20250502083951.xlsx
#     Jobcard Extraction Process    ${consolidated_report_path}    ${carry_over_ratio_report_path}    ${EMPTY}    ${date_timestamp}
   
# 
    # ${dms_full_image}      Set Variable         E:\\JobcardOpeningIntegrated\\Locators\\Screenshot_manual.png
    # ${JC_NUMBER}           Set Variable    JC25000575
    # # Extract Fields    ${JC_NUMBER}
    # &{output}    extract_image_data    ${dms_full_image}
    # Log To Console  ${output}
    # ${service_advisor_code}=    Get From Dictionary    ${output}    service_advisor_code
    # Log    ${service_advisor_code}
# *** Tasks ***    
# Demo    
#     Extract Demand Codes



    
    


    
    
    

    

    
    
    
    


    
    
    