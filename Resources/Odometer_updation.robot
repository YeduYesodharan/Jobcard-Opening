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
${odometer_update_img}    ${imagerootfolder}\\odometer_update_img.png
${odometer_branch_path}    3|1|1|1|1|1|1|1|2|2|1|1
${odometer_search_vehicle_button}    3|1|1|1|1|1|1|1|3|1
${odo_chas_no_radio}    1|2|1|2
${mob_no_odo_radio}    1|2|1|4
${Updated_reading}    3|1|1|1|2|1|1|1|1|1|1|7|2|1
${index}    1
${search_input_odo}    1|2|1|5|2
${go_odo}    1|2|1|5|3
${Odo_Branch_path}    3|1|1|1|1|1|1|1|2|2|1|1
# ${log_folder}     ${CURDIR}${/}..\\Log
${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
# ${log_folder}     C:\\Users\popular\Desktop\JobcardOpeningIntegrated\Jobcard-Opening\Screenshot
${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx
*** Keywords ***  

Odometer Updation Navigation
    [Arguments]    ${Chassis No.}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Odometer_reading}    ${Input_Sheet_Path}

    TRY

        Window Navigation    Wings ERP 23E - Web Client

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
        ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
        IF    ${trans_button_exists}==${True}            
            SikuliLibrary.Click    ${transaction_button}            
        END

        Click Action   name:${menu_transactions} > ${menu_AutoDms}
        Click Action   name:${menu_transactions} > ${menu_service} 
        Click Action   name:${menu_transactions} > ${menu_value_added_services}   
        Click Action   name:${menu_transactions} > ${menu_odometer_update}    #menu added in variables
            
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${odometer_update_img}     ${sik_max_time}
        ${odometer_update_img_heading_exists}=    SikuliLibrary.Exists    ${odometer_update_img}  #sikuli image added in locators
        
        IF    ${odometer_update_img_heading_exists}==${True}
            Click Action Maximum Retry    path:${Odo_Branch_path}
            # Double Click Action    path:${odometer_search_vehicle_button}
            # Log    odometer entered
        END
        Capture Screenshot
    EXCEPT  AS   ${error_message}

        Log    ${error_message} 
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Odometer Updation Navigation    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}     
        # Exit For Loop
    END 


Odometer Updation
    [Arguments]    ${Chassis No.}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Odometer_reading}    ${Input_Sheet_Path}    ${Branch}    ${Service Type Description}

    TRY

        Click Action Maximum Retry    path:${Odo_Branch_path}
        ${branch_code}    ERP_methods.Get Erp Branch Location Code    ${branch_mapping}    ${Branch}
        # RETURN  ${branch_code}
        IF    '${branch_code}' != 'branch code not available'
            RPA.Desktop.Press Keys    ctrl  a
            Type Text Action    ${branch_code}
            Press Keys Action    enter
            Double Click Action    path:${odometer_search_vehicle_button}
            Capture Screenshot

        ELSE
            RETURN  ${branch_code}
        END    

        ${alert_box_visible_1}    Odometer Vehicle Details Search with Chassis Number    ${Chassis No.}    ${Job Card No.}    ${Input_Sheet_Path}
        # ${alert_box_visible_1_lower}    Convert To Lower Case    ${alert_box_visible_1}   

        # IF  '${alert_box_visible_1_lower}' == 'message'
        IF  '${alert_box_visible_1}' != ''     #or '${alert_box_visible_1}' == 'Message'
            Capture Screenshot
            Click Action  path:${no_rec_ok_button}
            Click Action  path:${odo_chas_no_radio}            

            ${alert_box_visible_2}    Odometer Vehicle Details Search with Registration Number    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}
            # ${alert_box_visible_2_lower}    Convert To Lower Case    ${alert_box_visible_2}   

            # IF  '${alert_box_visible_2_lower}' == 'message'
            IF  '${alert_box_visible_2}' != ''     #or '${alert_box_visible_2}' == 'Message'
                Capture Screenshot
                Click Action  path:${no_rec_ok_button}
                Click Action  path:${mob_no_odo_radio} 
                ${alert_box_visible_3}    Odometer Vehicle Details Search with Mobile Number    ${Phone & Mobile No.}    ${Job Card No.}    ${Input_Sheet_Path}
                # ${alert_box_visible_3_lower}    Convert To Lower Case    ${alert_box_visible_3}  

                # IF  '${alert_box_visible_3_lower}' == 'message'
                IF  '${alert_box_visible_3}' != ''     #or '${alert_box_visible_3}' == 'Message'
                    Capture Screenshot
                    Click Action  path:${no_rec_ok_button}   
                    #odometer Updation
                    Sleep    ${Min_time}
                    Press Keys Action    esc
                    Press Keys Action    esc
                    Press Keys Action    enter                    
                                    
                ELSE    

                    ${chassis_search_value}    ${reg_search_value}    Validation in Mobile Number Search    ${Chassis No.}    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}    
                    # ${validation_chassis_stripped}=    Strip String    ${chassis_search_value}
                    # ${validation_reg_stripped}=    Strip String    ${reg_search_value}
                    # ${reg_stripped}=    Strip String    ${Registration No.}
                    # ${chassis_stripped}=    Strip String    ${Chassis No.}
                    

                    # IF    '${validation_chassis_stripped}' == '${chassis_stripped}' and '${validation_reg_stripped}' == '${reg_stripped}'
                    IF    '${chassis_search_value}' == '${Chassis No.}' and '${reg_search_value}' == '${Registration No.}'
                        Click Action Maximum Retry  path:${ok_to_tabs}
                        #odometer Updation
                        Sleep    ${Min_time}
                        Click Action    path:${Updated_reading}
                        RPA.Desktop.Press Keys    shift  right
                        Type Text Action    ${Odometer_reading}
                        Sleep    1
                        Capture Screenshot
                    ELSE
                        Press Keys Action    esc
                        Press Keys Action    esc
                        Press Keys Action    enter
                    END
                END           
                                
            ELSE  
                
                ${validation_status}    Validation in Registration Number Search    ${Chassis No.}    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}    
                # ${validation_status_stripped}=    Strip String    ${validation_status}
                # ${vehicle_id_stripped}=          Strip String    ${Chassis No.}
                # IF    '${validation_status_stripped}' == '${vehicle_id_stripped}'
                IF    '${validation_status}' == '${Chassis No.}'
                    Click Action Maximum Retry  path:${ok_to_tabs}
                    #odometer Updation
                    Sleep    ${Min_time}
                    Click Action    path:${Updated_reading}
                    Press Keys Action  delete
                    Type Text Action    ${Odometer_reading}
                    Sleep    1
                    Capture Screenshot
                ELSE
                    Press Keys Action    esc
                    Press Keys Action    esc
                    Press Keys Action    enter
                END
            END  
        ELSE         
            ${validation_status}    Validation in Chassis Number Search    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}    
            Capture Screenshot
            # ${validation_status_stripped}=    Strip String    ${validation_status}
            # ${reg_no_stripped}=          Strip String    ${Registration No.}
            # IF    '${validation_status_stripped}' == '${reg_no_stripped}'
            IF    '${validation_status}' == '${Registration No.}'
                Click Action Maximum Retry  path:${ok_to_tabs}
                #odometer Updation
                Sleep    ${Min_time}
                Click Action    path:${Updated_reading}
                RPA.Desktop.Press Keys    shift  right
                Type Text Action    ${Odometer_reading}
                Sleep    1
                Capture Screenshot
            ELSE
                Press Keys Action    esc
                Press Keys Action    esc
                Press Keys Action    enter
            END
                

        END

        Capture Screenshot

    EXCEPT  AS   ${error_message}
        Log    ${error_message} 
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Odometer Vehicle Search and Updations    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    ${search_window_close}    
        Press Keys Action    esc  
        Press Keys Action    enter
        Press Keys Action    esc  
        Press Keys Action    enter
        # Exit For Loop  
    END 



Odometer Vehicle Details Search with Chassis Number
    [Arguments]    ${Chassis No.}   ${Job Card No.}    ${Input_Sheet_Path}
    TRY        
        
        Click Action    path:${search_input_odo}
        Type Text Action    ${Chassis No.}
        Sleep    ${Min_time}
        Click Action    path:${go_odo}
        Sleep    ${Min_time}
        ${alert_box_visible_1}   Get Text Action Maximum Retry    path:${alert_box}
        Capture Screenshot
        RETURN    ${alert_box_visible_1}
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Odometer Vehicle Search with Chassis Number    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    ${search_window_close}
        # Exit For Loop
    END 


Odometer Vehicle Details Search with Registration Number
    [Arguments]    ${Registration No.}   ${Job Card No.}    ${Input_Sheet_Path}
    TRY        
        
        Click Action    path:${search_input_odo}
        Type Text Action    ${Registration No.}
        Sleep    ${Min_time}
        Click Action    path:${go_odo}
        Sleep    ${Min_time}
        ${alert_box_visible_2}   Get Text Action Maximum Retry    path:${alert_box}
        Capture Screenshot
        RETURN    ${alert_box_visible_2}
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Odometer Vehicle Search with Registration Number    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    ${search_window_close}
        # Exit For Loop
    END 

Odometer Vehicle Details Search with Mobile Number
    [Arguments]    ${Phone & Mobile No.}   ${Job Card No.}    ${Input_Sheet_Path}
    TRY        
        
        Click Action    path:${search_input_odo}
        Type Text Action    ${Phone & Mobile No.}
        Sleep    ${Min_time}
        Click Action    path:${go_odo}
        Sleep    ${Min_time}
        ${alert_box_visible_3}   Get Text Action Maximum Retry    path:${alert_box}
        Capture Screenshot
        RETURN    ${alert_box_visible_3}
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error Occured in Odometer Vehicle Search with Mobile Number    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    ${search_window_close}
        # Exit For Loop
    END 



    
    
