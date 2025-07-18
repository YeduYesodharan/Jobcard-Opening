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
Resource            Main_Flow.robot
Resource            Odometer_updation.robot



*** Variables ***
${price_list}    3|1|1|1|1|1|1|1|13|2|1|1  
${service_price_list}    3|1|1|1|1|1|1|1|14|2|1|1
${check_list_type}    3|1|1|1|1|1|1|1|15|2|1|1
${manual_jobcard_number}    3|1|1|1|2|1|1|1|1|1|1|1|2|1
${expected_delivery_date}    3|1|1|1|2|1|1|1|1|1|1|2|2|1
${type_of_service}    3|1|1|1|2|1|1|1|1|1|1|5|2|1|1
${service_type}    3|1|1|1|2|1|1|1|1|1|1|6|2|1|1
${service_advisor}      3|1|1|1|2|1|1|1|1|1|1|11|2|1|1    #3|1|1|1|2|1|1|1|1|1|1|12|2|1|1         
${fuel_level}    3|1|1|1|2|1|1|1|1|1|1|17|2|1|1    #3|1|1|1|2|1|1|1|1|1|1|16|2|1|1    
${odometer_reading_inp}    3|1|1|1|2|1|1|1|1|1|1|16|2|1    #3|1|1|1|2|1|1|1|1|1|1|15|2|1    
${expected_delivery time}    3|1|1|1|2|1|1|1|1|1|1|3|2|1
${edtime_title}     3|1|1|1|2|1|1|1|1|1|1|3|1
${type_of_service_title}    3|1|1|1|2|1|1|1|1|1|1|5|1
${service_type_title}    3|1|1|1|2|1|1|1|1|1|1|6|1
${tos_threedots}    3|1|1|1|2|1|1|1|1|1|1|5|2|2|2
${advisor_name_title}    3|1|1|1|2|1|1|1|1|1|1|11|1
${odometer title}    3|1|1|1|2|1|1|1|1|1|1|16|1    #3|1|1|1|2|1|1|1|1|1|1|15|1    
${fuel_level title}    3|1|1|1|2|1|1|1|1|1|1|17|1    #3|1|1|1|2|1|1|1|1|1|1|16|1    
${job_code}    3|1|1|1|2|1|1|1|1|1|1|1|1|1|2|3
${step_details_working_btn}    3|1|1|1|2|1|1|1|1|1|1|1|1|1|1|3                                 
${check_all}    1|1
${check_confirm}    1|1
${final_save}    3|1|1|3|1|2|1|1|2    
${previous_odometer title}    3|1|1|1|2|1|1|1|1|1|1|15|1    #3|1|1|1|2|1|1|1|1|1|1|14|1
${previous_omr_input}    3|1|1|1|2|1|1|1|1|1|1|15|2|1    #3|1|1|1|2|1|1|1|1|1|1|14|2
${working_btn}   ${imagerootfolder}\\working_btn.png
${final_save_content}    1|3
${final_save_ok}    1|1
# ${Service_Type_Sheet}       ${CURDIR}${/}..\\Mapping\\Service Type Cumulative.xlsx
${Service_Type_Sheet}    C:\\JobcardOpeningIntegrated\\Mapping\\Service Type Cumulative.xlsx
${advisor_list_sheet}    C:\\JobcardOpeningIntegrated\\Mapping\\Service Advisor List ERP SRM_ELM.xlsx
# ${advisor_list_sheet}       ${CURDIR}${/}..\\Mapping\\Service Advisor List ERP SRM_ELM.xlsx
${log_folder}     ${CURDIR}${/}..\\Screenshot
${prefix}    Jobcard already opened for
${body_check_list_type}    3|1|1|1|1|1|1|1|9|2|1|1        
${body_previous_odometer title}    3|1|1|1|2|1|1|1|1|1|1|7|1
${body_prv_omr_input}    3|1|1|1|2|1|1|1|1|1|1|7|2|1 
${body_odometer_title}    3|1|1|1|2|1|1|1|1|1|1|8|1 
${body_odometer_reading_inp}    3|1|1|1|2|1|1|1|1|1|1|8|2|1    
${body_manual_jobcard_number_inp}    3|1|1|1|2|1|1|1|1|1|1|1|2|1
${body_eddate_title}    3|1|1|1|2|1|1|1|1|1|1|5|1
${body_expected_delivery_date_inp}    3|1|1|1|2|1|1|1|1|1|1|5|2|1
${body_edtime_title}    3|1|1|1|2|1|1|1|1|1|1|6|1 
${body_expected_delivery time_inp}    3|1|1|1|2|1|1|1|1|1|1|6|2|1
${body_mjc_title}    3|1|1|1|2|1|1|1|1|1|1|1|1
${body_type_of_service_title}    3|1|1|1|2|1|1|1|1|1|1|10|1  
${body_tos_inp}    3|1|1|1|2|1|1|1|1|1|1|10|2|1|1  
${body_service_advisor_inp}    3|1|1|1|2|1|1|1|1|1|1|12|2|1|1  
${substring}    Jobcard already opened 
# ${branch_mapping}    ${CURDIR}${/}..\\Mapping\\Location Mapping DMS ERP.xlsx
${branch_mapping}    C:\\JobcardOpeningIntegrated\\Mapping\\Location Mapping DMS ERP.xlsx



*** Keywords ***

JobCard Tab Entry Activation

    [Arguments]  ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}

    ${Service Type Description_value}    Clean String    ${Service Type Description}
    
    IF    "${Service Type Description_value}" != "bodyrepair" and "${Service Type Description_value}" != "bandp"

        TRY
            
            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg1}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg1}
            Log    ${used_jc_msg1}
            IF  '${used_jc_msg1}' != '${None}'
                IF  '${used_jc_msg1}' != 'Dashboard'
                    RETURN  ${used_jc_msg1}
                END
            END
            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg2}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg1}
            Log    ${used_jc_msg2}
            IF  '${used_jc_msg2}' != '${None}'
                IF  '${used_jc_msg2}' != 'Dashboard'
                    RETURN  ${used_jc_msg2}
                END
            END
            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg3}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg1}
            Log    ${used_jc_msg3}
            IF  '${used_jc_msg3}' != '${None}'
                IF  '${used_jc_msg3}' != 'Dashboard'
                    RETURN  ${used_jc_msg3}
                END
            END


            Click Action    path:${price_list}  
            
            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg4}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg1}
            Log    ${used_jc_msg4}
            IF  '${used_jc_msg4}' != '${None}'
                IF  '${used_jc_msg4}' != 'Dashboard'
                    RETURN  ${used_jc_msg4}
                END
            END
            Run Keyword And Ignore Error    Additional Windows
            Click Action    path:${service_price_list}  
            ${used_jc_msg5}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg2}
            Log    ${used_jc_msg5}
            IF  '${used_jc_msg5}' != '${None}'
                IF  '${used_jc_msg5}' != 'Dashboard'
                    RETURN  ${used_jc_msg5}
                END
            END
       
            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg6}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg1}
            Log    ${used_jc_msg6}
            IF  '${used_jc_msg6}' != '${None}'
                IF  '${used_jc_msg6}' != 'Dashboard'
                    RETURN  ${used_jc_msg6}
                END
            END
            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg7}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg1}
            Log    ${used_jc_msg7}
            IF  '${used_jc_msg7}' != '${None}'
                IF  '${used_jc_msg7}' != 'Dashboard'
                    RETURN  ${used_jc_msg7}
                END
            END
            Click Action    path:${check_list_type}   
            
        EXCEPT  AS  ${click_error}
        
            Log  ${click_error}
            Capture Screenshot
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while activating the JobCard Tab Entry.    ${exception_reason_column_name}
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
            # Exit For Loop
        END 

    ELSE
        
            TRY

            Run Keyword And Ignore Error    Additional Windows  
            Run Keyword And Ignore Error    Additional Windows
            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg1}    Used Jobcard
            IF  '${used_jc_msg1}' != '${None}'
                IF  '${used_jc_msg1}' != 'Dashboard'
                    RETURN  ${used_jc_msg1}
                END
            END
            Click Action    path:${body_check_list_type} 


            #added when there is ui reversion on bodyshop
            Run Keyword And Ignore Error    Additional Windows  

            ${used_jc_msg2}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg4}
            Log    ${used_jc_msg2}
            IF  '${used_jc_msg2}' != '${None}'
                IF  '${used_jc_msg2}' != 'Dashboard'
                    RETURN  ${used_jc_msg2}
                END
            END
            Run Keyword And Ignore Error    Additional Windows  

        EXCEPT  AS  ${click_error}
        
            Log  ${click_error}
            Capture Screenshot
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while activating the JobCard Tab Entry.    ${exception_reason_column_name}
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
            # Exit For Loop
        END 
        

    END


Service Details Tab

    [Arguments]  ${Job Card No.}    ${Promised Dt.}    ${Type of Service}     ${S.A}    ${Odometer Reading}     
    ...    ${Demand Code}    ${Service Type Code}    ${Input_Sheet_Path}   ${Registration No.}    ${Chassis No.}
    ...    ${Phone & Mobile No.}    ${Service Type Description}    ${Sub Service Type}    ${Branch}
     
    ${Service Type Description_value}    Clean String    ${Service Type Description}
     
    IF    "${Service Type Description_value}" != "bodyrepair" and "${Service Type Description_value}" != "bandp"

        TRY

            # Press Keys Action    f6
            # Press Keys Action    f8
            Press Keys Action    f5
               
            ${prvious_omr_title_value}    Get Text Action    path:${previous_odometer title}
            Log    ${prvious_omr_title_value}
            ${prvious_omr_title_value_lower}    Convert To Lower Case    ${prvious_omr_title_value}
            IF    "${prvious_omr_title_value_lower}" == "previous reading"
        
                Click Action    path:${previous_omr_input}       
                ${previous_omr_value}    Get Value Action    path:${previous_omr_input}

            END 
            ${used_jc_msg5}    Used Jobcard
            IF  '${used_jc_msg5}' != '${None}'
                
                IF  '${used_jc_msg5}' != 'Dashboard'
                    RETURN  ${used_jc_msg5}
                END
                
            END
            # Show Message Box    title    ${used_jc_msg5}
            Log    ${used_jc_msg5}
            IF    ${previous_omr_value} == 0

                ${odometer_title_value}   Get Text Action    path:${odometer_title}
                Log    ${odometer_title_value}
                ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                IF    "${odometer_title_value_lower}" == "odometer reading *"
                    # Press Keys Action    tab
                    Double Click Action    path:${odometer_reading_inp}  
                    # RPA.Desktop.Press Keys  Ctrl  a
                    Type Text Action    ${Odometer Reading}
                    
                END  
                Capture Screenshot
            ELSE IF  '${previous_omr_value}' == '${Odometer_Reading}'

                ${odometer_title_value}   Get Text Action    path:${odometer_title}
                Log    ${odometer_title_value}
                ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                IF    "${odometer_title_value_lower}" == "odometer reading *"
                    # Press Keys Action    tab
                    Double Click Action    path:${odometer_reading_inp}  
                    # RPA.Desktop.Press Keys  Ctrl  a
                    Type Text Action    ${Odometer Reading}
                END  
                Capture Screenshot

            #---------------------------------Odo new change added-------------------------------------------------------#
            ELSE IF  '${previous_omr_value}' < '${Odometer_Reading}'

                ${odometer_title_value}   Get Text Action    path:${odometer_title}
                Log    ${odometer_title_value}
                ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                IF    "${odometer_title_value_lower}" == "odometer reading *"
                    # Press Keys Action    tab
                    Double Click Action    path:${odometer_reading_inp}  
                    # RPA.Desktop.Press Keys  Ctrl  a
                    Type Text Action    ${Odometer Reading}
                END  
                Capture Screenshot
            #---------------------------------Odo new change added-------------------------------------------------------#
            
            ELSE  

            # END
            # IF    ${previous_omr_value} < ${Odometer_Reading}
                
            #     Log   pass as exception and go for next iteration
                # RETURN    Odometer Exception Occured:DMS reading is higher than ERP reading
            
            # ELSE IF  ${previous_omr_value} > ${Odometer_Reading}
                Press Keys Action    esc  
                Press Keys Action    enter
                Odometer Updation Navigation  ${Chassis No.}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Odometer_reading}    ${Input_Sheet_Path}    
                ${odm_updtn_status}    Odometer Updation  ${Chassis No.}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Odometer_reading}    ${Input_Sheet_Path}   ${Branch}    ${Service Type Description}
                # IF    '${odm_updtn_status}' == 'branch code not available'
                #     RETURN  ${odm_updtn_status}

                # #for testing != changed to ==
                # ELSE IF  '${odm_updtn_status}' != 'Transaction saved.'                               
                #     RETURN   'odometer updation failed' 
                # END
                IF    '${odm_updtn_status}' == 'branch code not available'
                    RETURN  ${odm_updtn_status} 
                END
                #just commented--------------------------------#
                ${save_status}    ${save_message}    Final Save    ${Job Card No.}    ${Input_Sheet_Path}
                IF    ${save_status} == ${False}
                    Press Keys Action    esc 
                    Press Keys Action    enter
                    RETURN    odometer updation failed
                ELSE
                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    ${save_message}   ${exception_reason_column_name}
                    Press Keys Action    esc 
                    Press Keys Action    enter
                END
                 #just commented--------------------------------#
                # reopen vehicle details for JC creation

                IF    "${Service Type Description_value}" != "bodyrepair" and "${Service Type Description_value}" != "bandp"

                    ${nav_status2}    Service Job Card Navigation
                    IF    '${nav_status2}' == '${True}'                        
                        ${status_val}    Service Jobcard Branch and Location path    ${Job Card No.}    ${Input_Sheet_Path}    ${Branch}
                    END

                ELSE
                    
                    ${nav_status1}    BodyShop Job Card Navigation   
                    Capture Screenshot
                    IF    '${nav_status1}' == '${True}'                                       
                        ${status_val}    Bodyshop Branch and Location path    ${Job Card No.}    ${Input_Sheet_Path}    ${Branch}
                    END

                END

                IF    '${status_val}' != 'branch code not available'
                    ${alert_box_visible_1}    Vehicle Details Search with Registration Number    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}
                    ${alert_box_visible_1_lower}    Convert To Lower Case    ${alert_box_visible_1}
                    
                    IF  '${alert_box_visible_1_lower}' == 'message'
                    # IF  '${alert_box_visible_1_lower}' != ''
                        Click Action  path:${no_rec_ok_button}
                        Click Action  path:${Chassis_No} 
                        
                        ${alert_box_visible_2}    Vehicle Details Search with Chassis Number    ${Chassis No.}    ${Job Card No.}    ${Input_Sheet_Path}
                        ${alert_box_visible_2_lower}    Convert To Lower Case    ${alert_box_visible_2}

                        IF  '${alert_box_visible_2_lower}' == 'message'
                        # IF  '${alert_box_visible_2_lower}' != ''
                            Click Action  path:${no_rec_ok_button}
                            Click Action  path:${Mobile_no}

                            ${alert_box_visible_3}    Vehicle Details Search with Mobile Number    ${Phone & Mobile No.}    ${Job Card No.}    ${Input_Sheet_Path}
                            ${alert_box_visible_3_lower}    Convert To Lower Case    ${alert_box_visible_3}
                    
                            IF  '${alert_box_visible_3_lower}' == 'message'
                            # IF  '${alert_box_visible_3_lower}' != ''
                                Click Action  path:${no_rec_ok_button}
                            ELSE   
                                Click Action Maximum Retry  path:${ok_to_tabs}
                                Sleep    ${Min_time}

                            END
                        
                        ELSE   
                            Click Action Maximum Retry  path:${ok_to_tabs}
                            Sleep    ${Min_time}
                            
                        END

                    ELSE   
                        Click Action Maximum Retry  path:${ok_to_tabs}
                        Sleep    ${Min_time}
                        
                    END

                ELSE
                    RETURN   ${status_val}
                END 
                JobCard Tab Entry Activation    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}  
                Run Keyword And Ignore Error    Additional Windows

                ${used_jc_msg7}    Used Jobcard
                # Show Message Box    title    ${used_jc_msg7}
                Log    ${used_jc_msg7}
                IF  '${used_jc_msg7}' != '${None}'
                    IF  '${used_jc_msg7}' != 'Dashboard'
                        RETURN  ${used_jc_msg7}
                    END
                END
                # Press Keys Action    f8 

                ${pickup_status}    Pickup Details Tab     ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}

                IF  """${pickup_status}""" == """Some mandatory values except Pickup Driver is missing"""
                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Some mandatory values except Pickup Driver is missing    ${exception_reason_column_name}
                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
                    Press Keys Action    esc 
                    Press Keys Action    enter

                ELSE
                        Press Keys Action    f5
                        
                        Run Keyword And Ignore Error    Additional Windows

                        ${used_jc_msg7}    Used Jobcard
                        # Show Message Box    title    ${used_jc_msg7}
                        Log    ${used_jc_msg7}
                        IF  '${used_jc_msg7}' != '${None}'
                            IF  '${used_jc_msg7}' != 'Dashboard'
                                RETURN  ${used_jc_msg7}
                            END
                        END

                        Run Keyword And Ignore Error    Additional Windows


                        #enter updated odometer value in odmeter reading field 
                        ${odometer_title_value}   Get Text Action    path:${odometer_title}
                        Log    ${odometer_title_value}
                        ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                        IF    "${odometer_title_value_lower}" == "odometer reading *"
                            # Press Keys Action    tab
                            Double Click Action    path:${odometer_reading_inp}  
                            # RPA.Desktop.Press Keys  Ctrl  a
                            # Run Keyword And Ignore Error    Additional Windows
                            ${used_jc_msg7}    Used Jobcard
                            # Show Message Box    title    ${used_jc_msg7}
                            Log    ${used_jc_msg7}
                            IF  '${used_jc_msg7}' != '${None}'
                                IF  '${used_jc_msg7}' != 'Dashboard'
                                    RETURN  ${used_jc_msg7}
                                END
                            END
                            Type Text Action    ${Odometer Reading}
                        END 

                END            
                    
            END
                Capture Screenshot
                Click Action    path:${manual_jobcard_number}   
                Type Text Action    ${Job Card No.}
            

                ${Promised Date.}   ${Promised Time.}    ERP_methods.Extract Date And Time    ${Promised_Dt.}    
                Log    ${Promised Date.}
                Log    ${Promised Time.}
                Click Action    path:${expected_delivery_date}    
                Type Text Action    ${Promised Date.}
                # ERP_methods.Show Message Box    title    ${Promised Date.}
                
                ${ed_time_title_value}    Get Text Action    path:${edtime_title}
                ${ed_time_title_value_lower}    Convert To Lower Case    ${ed_time_title_value}
                IF    "${ed_time_title_value_lower}" == "expected delivery time *"

                    Press Keys Action    tab
                    Press Keys Action    backspace
                    Click Action    path:${expected_delivery time}
                    # ERP_methods.Show Message Box    title    ${Promised Time.}
                    Type Text Action    ${Promised Time.}
                    
                    
                END

            
                ${typofservice_title_value}    Get Text Action    path:${type_of_service_title}    
                Log    ${typofservice_title_value}
                ${typofservice_title_value_lower}    Convert To Lower Case    ${typofservice_title_value}
                IF    "${typofservice_title_value_lower}" == "type of service *"
                    
                    ${type_of_service_output}    ERP_methods.Map Type Of Service    ${Service_Type_Sheet}    ${Service Type Code}   ${Sub Service Type}
                    
                    IF    '${type_of_service_output}' != 'no match found for type of service'

                        Press Keys Action    tab                    
                        Type Text Action    ${type_of_service_output}
                        Capture Screenshot
                        Press Keys Action    enter    #will take to Service_type_tab
                        # RETURN  ${True}
                    ELSE
                        RETURN    ${type_of_service_output}
                    END         
                END
            
                ${servicetype_title_value}    Get Text Action    path:${service_type_title}       
                Log    ${servicetype_title_value}
                ${servicetype_title_value_lower}    Convert To Lower Case    ${servicetype_title_value}
                IF    "${servicetype_title_value_lower}" == "service type *"
                    
                    ${ServiceType_Output}    ERP_methods.Sort And Get Service Type From Excel    ${Service_Type_Sheet}    ${type_of_service_output}    ${Service Type Code}
                    IF    "${ServiceType_Output}" == "no match found for service type" 

                        ${ServiceType_Output}    ERP_methods.Sort And Get Service Type From Excel    ${Service_Type_Sheet}    ${type_of_service_output}    ${Sub Service Type}
                                                
                    END
                    
                    IF    "${ServiceType_Output}" != "no match found for service type"  

                        Type Text Action    ${ServiceType_Output}
                        Capture Screenshot
                        Log    ${ServiceType_Output}
                        Press Keys Action    enter    #will take to advisor tab

                        ${used_jc_msgs}    Used Jobcard
                        IF  '${used_jc_msgs}' != '${None}'
                                            
                            IF  '${used_jc_msgs}' != 'Dashboard'
                                RETURN  ${used_jc_msgs}
                            END

                        END  
    
                    ELSE
                        RETURN  ${ServiceType_Output}
                    END   
                END

                ${advisor_output}    ERP_methods.Advisor Name From Dms    ${advisor_list_sheet}    ${S.A}
                Log    ${advisor_output}
                  

                ${used_jc_msgg}    Used Jobcard
                IF  '${used_jc_msgg}' != '${None}'
                    IF  '${used_jc_msgg}' != 'Dashboard'
                        RETURN  ${used_jc_msgg}
                    END
                END  

                ${used_jc_msgss}    Used Jobcard
                IF  '${used_jc_msgss}' != '${None}'                       
                    IF  '${used_jc_msgss}' != 'Dashboard'
                        RETURN  ${used_jc_msgss}
                    END
                END

                
                IF    '${advisor_output}' != 'no match found for service advisor'
                    Click Action    path:${service_advisor} 
                    
                    
                    #CHECK for km exceeds exception
                    ${used_jc_msg}    Used Jobcard
                    IF  '${used_jc_msg}' != '${None}'                       
                        IF  '${used_jc_msg}' != 'Dashboard'
                            RETURN  ${used_jc_msg}
                        END
                    END  

                    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  a   
                    # ERP_methods.Show Message Box    tos    ${advisor_output}
                    Type Text Action    ${advisor_output}
                    Capture Screenshot
                    # Click Action    path:${service_advisor} 
                    Press Keys Action    enter
                    #CHECK for km exceeds exception
                    ${used_jc_msg}    Used Jobcard
                    IF  '${used_jc_msg}' != '${None}'
                        IF  '${used_jc_msg}' != 'Dashboard'
                            RETURN  ${used_jc_msg}
                        END
                    END  
                ELSE
                    RETURN    ${advisor_output}
                END 

                ${fuel_level title_value}   Get Text Action    path:${fuel_level title}
                Log    ${fuel_level title_value}
                ${fuel_level title_value_lower}    Convert To Lower Case    ${fuel_level title_value}
                IF    "${fuel_level_title_value_lower}" == "fuel level *"
                    Press Keys Action    tab
                    Press Keys Action    tab
                    Press Keys Action    tab
                    Type Text Action    ${fuel_level}
                    Press Keys Action    enter
                    Capture Screenshot
                END
                Capture Screenshot
                
                # RETURN    ${False}
    
        EXCEPT  AS   ${Service_Detail_Tab_error}
            Log  ${Service_Detail_Tab_error}
            Capture Screenshot
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred in the Service Details Tab.    ${exception_reason_column_name}
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
            # Exit For Loop
        END   
    
    ELSE   #if its body repair

        TRY 

            # Press Keys Action    f6
            # Press Keys Action    f8
            Press Keys Action    f5

            Run Keyword And Ignore Error    Additional Windows
            ${used_jc_msg7}    Used Jobcard
            # Show Message Box    title    ${used_jc_msg1}
            Log    ${used_jc_msg7}
            IF  '${used_jc_msg7}' != '${None}'
                IF  '${used_jc_msg7}' != 'Dashboard'
                    RETURN  ${used_jc_msg7}
                END
            END
            
            ${prvious_omr_title_value}    Get Text Action    path:${body_previous_odometer title}
            Log    ${prvious_omr_title_value}
            ${prvious_omr_title_value_lower}    Convert To Lower Case    ${prvious_omr_title_value}
            IF    "${prvious_omr_title_value_lower}" == "previous reading"
           
                Click Action    path:${body_prv_omr_input}     
                ${previous_omr_value}    Get Value Action    path:${body_prv_omr_input}  

            END 

            IF    ${previous_omr_value} == 0
                
                ${odometer_title_value}   Get Text Action    path:${body_odometer_title}
                Log    ${odometer_title_value}
                ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                IF    "${odometer_title_value_lower}" == "present odo meter reading *"
                    # Press Keys Action    tab
                    Double Click Action    path:${body_odometer_reading_inp}  
                    Type Text Action    ${Odometer Reading}
                    Capture Screenshot
                END 

            ELSE IF  '${previous_omr_value}' == '${Odometer_Reading}'    
                
                ${odometer_title_value}   Get Text Action    path:${body_odometer_title}
                Log    ${odometer_title_value}
                ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                IF    "${odometer_title_value_lower}" == "present odo meter reading *"
                    # Press Keys Action    tab
                    Double Click Action    path:${body_odometer_reading_inp}  
                    Type Text Action    ${Odometer Reading}
                    Capture Screenshot
                END 

            #---------------------------------Odo new change added-------------------------------------------------------#
            ELSE IF  '${previous_omr_value}' < '${Odometer_Reading}'

                ${odometer_title_value}   Get Text Action    path:${body_odometer_title}
                Log    ${odometer_title_value}
                ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                IF    "${odometer_title_value_lower}" == "present odo meter reading *"
                    # Press Keys Action    tab
                    Double Click Action    path:${body_odometer_reading_inp}  
                    Type Text Action    ${Odometer Reading}
                    Capture Screenshot
                END 
            #---------------------------------Odo new change added-------------------------------------------------------#

            
            ELSE
                
                Press Keys Action    esc  
                Press Keys Action    enter
                Odometer Updation Navigation  ${Chassis No.}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Odometer_reading}    ${Input_Sheet_Path}
                ${odm_updtn_status}    Odometer Updation  ${Chassis No.}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Odometer_reading}    ${Input_Sheet_Path}    ${Branch}    ${Service Type Description}
                # IF    '${odm_updtn_status}' == 'branch code not available'
                #     RETURN  ${odm_updtn_status}                                              
                # #for testing != changed to ==
                # ELSE IF  '${odm_updtn_status}' != 'Transaction saved.'                               
                #     RETURN   'odometer updation failed' 
                # END
                IF    '${odm_updtn_status}' == 'branch code not available'
                    RETURN  ${odm_updtn_status} 
                END
                 #just commented--------------------------------#
                ${save_status}    ${save_message}    Final Save    ${Job Card No.}    ${Input_Sheet_Path}
                IF    ${save_status} == ${False}
                    Press Keys Action    esc 
                    Press Keys Action    enter
                    RETURN    odometer updation failed
                ELSE
                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    ${save_message}   ${exception_reason_column_name}
                    Press Keys Action    esc 
                    Press Keys Action    enter
                END
                 #just commented--------------------------------#
                # reopen vehicle details for JC creation

                IF    "${Service Type Description_value}" != "bodyrepair" and "${Service Type Description_value}" != "bandp"

                    ${nav_status2}    Service Job Card Navigation
                    IF    '${nav_status2}' == '${True}'                        
                        ${status_val}    Service Jobcard Branch and Location path    ${Job Card No.}    ${Input_Sheet_Path}    ${Branch}
                    END

                ELSE
                    
                    ${nav_status1}    BodyShop Job Card Navigation   
                    IF    '${nav_status1}' == '${True}'                                       
                        ${status_val}    Bodyshop Branch and Location path    ${Job Card No.}    ${Input_Sheet_Path}    ${Branch}
                    END

                END

                IF    '${status_val}' != 'branch code not available'
                    ${alert_box_visible_1}    Vehicle Details Search with Registration Number    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}
                    ${alert_box_visible_1_lower}    Convert To Lower Case    ${alert_box_visible_1}
                    
                    IF  '${alert_box_visible_1_lower}' == 'message'
                    # IF  '${alert_box_visible_1_lower}' != ''
                        Click Action  path:${no_rec_ok_button}
                        Click Action  path:${Chassis_No} 
                        
                        ${alert_box_visible_2}    Vehicle Details Search with Chassis Number    ${Chassis No.}    ${Job Card No.}    ${Input_Sheet_Path}
                        ${alert_box_visible_2_lower}    Convert To Lower Case    ${alert_box_visible_2}

                        IF  '${alert_box_visible_2_lower}' == 'message'
                        # IF  '${alert_box_visible_2_lower}' != ''
                            Click Action  path:${no_rec_ok_button}
                            Click Action  path:${Mobile_no}

                            ${alert_box_visible_3}    Vehicle Details Search with Mobile Number    ${Phone & Mobile No.}    ${Job Card No.}    ${Input_Sheet_Path}
                            ${alert_box_visible_3_lower}    Convert To Lower Case    ${alert_box_visible_3}
                    
                            IF  '${alert_box_visible_3_lower}' == 'message' 
                            # IF  '${alert_box_visible_3_lower}' != ''
                                Click Action  path:${no_rec_ok_button}
                            ELSE   
                                Click Action Maximum Retry  path:${ok_to_tabs}
                                Sleep    ${Min_time}

                            END
                        
                        ELSE   
                            Click Action Maximum Retry  path:${ok_to_tabs}
                            Sleep    ${Min_time}
                            
                        END

                    ELSE   
                        Click Action Maximum Retry  path:${ok_to_tabs}
                        Sleep    ${Min_time}
                        
                    END

                ELSE
                    RETURN   ${status_val}
                END 
                JobCard Tab Entry Activation    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}  
                # Press Keys Action    f8 

                ${pickup_status}    Pickup Details Tab     ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}

                IF  """${pickup_status}""" == """Some mandatory values except Pickup Driver is missing"""
                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Some mandatory values except Pickup Driver is missing    ${exception_reason_column_name}
                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
                    Press Keys Action    esc 
                    Press Keys Action    enter
                    
                ELSE

                    Press Keys Action    f5
                    
                    Run Keyword And Ignore Error    Additional Windows
                    ${used_jc_msg7}    Used Jobcard
                    # Show Message Box    title    ${used_jc_msg1}
                    Log    ${used_jc_msg7}
                    IF  '${used_jc_msg7}' != '${None}'
                        IF  '${used_jc_msg7}' != 'Dashboard'
                            RETURN  ${used_jc_msg7}
                        END
                    END


                    #enter updated odometer value in odmeter reading field 
                    ${odometer_title_value}   Get Text Action    path:${body_odometer_title}
                    Log    ${odometer_title_value}
                    ${odometer_title_value_lower}    Convert To Lower Case    ${odometer_title_value}
                    IF    "${odometer_title_value_lower}" == "present odo meter reading *"
                        # Press Keys Action    tab
                        Double Click Action    path:${body_odometer_reading_inp}  
                        Type Text Action    ${Odometer Reading}
                        Capture Screenshot
                    END 
                END 

            END
        
            
                ${body_mjc_title_value}    Get Text Action    path:${body_mjc_title}
                ${body_mjc_title_value_lower}    Convert To Lower Case    ${body_mjc_title_value}
                IF    "${body_mjc_title_value_lower}" == "manual jobcard no *"

                    Click Action    path:${body_manual_jobcard_number_inp}   
                    Type Text Action    ${Job Card No.}
                
                END
                           
                ${ed_date_title_value}    Get Text Action    path:${body_eddate_title}
                ${ed_date_title_value_lower}    Convert To Lower Case    ${ed_date_title_value}
                IF    "${ed_date_title_value_lower}" == "expected delivery date *"

                    ${Promised Date.}   ${Promised Time.}    ERP_methods.Extract Date And Time    ${Promised_Dt.}    
                    Click Action    path:${body_expected_delivery_date_inp}    
                    Type Text Action    ${Promised Date.}
                    
                    
                END

            
                ${ed_time_title_value}    Get Text Action    path:${body_edtime_title}
                ${ed_time_title_value_lower}    Convert To Lower Case    ${ed_time_title_value}
                IF    "${ed_time_title_value_lower}" == "expected delivery time *"

                    Press Keys Action    tab
                    Press Keys Action    backspace
                    Click Action    path:${body_expected_delivery time_inp}
                    Type Text Action    ${Promised Time.}
                    
                    
                END


                # ${type_of_service_output}    ERP_methods.Map Type Of Service    ${Input_Sheet_Path}    ${Service_Type_Sheet}
                ${typofservice_title_value}    Get Text Action    path:${body_type_of_service_title}    
                Log    ${typofservice_title_value}
                ${typofservice_title_value_lower}    Convert To Lower Case    ${typofservice_title_value}
                IF    "${typofservice_title_value_lower}" == "type of service *"

                    ${type_of_service_output}    ERP_methods.Map Type Of Service    ${Service_Type_Sheet}    ${Service Type Code}   ${Sub Service Type}                 
                    IF    '${type_of_service_output}' != 'no match found for type of service'

                        Click Action    path:${body_tos_inp}
                        Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys  ctrl  a
                        Type Text Action    ${type_of_service_output}
                        Capture Screenshot
                        Press Keys Action    enter

                        #CHECK for km exceeds exception
                        ${used_jc_msg}    Used Jobcard
                        IF  '${used_jc_msg}' != '${None}'
                            IF  '${used_jc_msg}' != 'Dashboard'
                                RETURN  ${used_jc_msg}
                            END
                        END  

                    ELSE
                        RETURN    ${type_of_service_output}
                    END         
                END                 
                #     Click Action    path:${body_tos_inp}
                #     RPA.Desktop.Press Keys  ctrl  a
                #     Type Text Action    ${type_of_service_output}
                #     Press Keys Action    enter
                    
                # END

                ${advisor_output}    ERP_methods.Advisor Name From Dms    ${advisor_list_sheet}    ${S.A}
                Log    ${advisor_output}
                  

                ${used_jc_msg}    Run Keyword And Ignore Error    Used Jobcard
                ${used_jc_msg}    Used Jobcard
                IF  '${used_jc_msg}' != '${None}'
                    IF  '${used_jc_msg}' != 'Dashboard'
                        RETURN  ${used_jc_msg}
                    END
                END 
                
                IF    '${advisor_output}' != 'no match found for service advisor'

                    # Click Action    path:${body_service_advisor_inp}  
                    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  a   
                    Type Text Action    ${advisor_output}
                    Capture Screenshot
                    # Click Action    path:${body_service_advisor_inp} 
                    Press Keys Action    enter
                    Sleep    ${Min_time}
                ELSE
                        RETURN    ${advisor_output}
                END   
                # RETURN    ${False}

                Capture Screenshot
        
        EXCEPT  AS   ${Service_Detail_Tab_error}
            Log  ${Service_Detail_Tab_error}
            Capture Screenshot
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred in the Service Details Tab.    ${exception_reason_column_name}
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
            # Exit For Loop
        END 
    END




Suggested Jobs Tab
    [Arguments]  ${Job Card No.}    ${Demand Code}    ${Input_Sheet_Path}    ${Service Type Description}
    
    TRY
        ${Service Type Description_value}    Clean String    ${Service Type Description}
        IF    "${Service Type Description_value}" == "bodyrepair" or "${Service Type Description_value}" == "bandp"

            # Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f3
            # Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f9
            Press Keys Action    f9


        ELSE
            # Press Keys Action    f7
            # Press Keys Action    f9
            Press Keys Action    f6
        
        END
            
        ${demand_codes_list}    ERP_methods.Extract Demand Codes    ${Demand Code}

        # Click on the job code element
        Click Action    path:${job_code}  
        FOR    ${each_demand_code}    IN    @{demand_codes_list}  

            # # Type the current demand code and press enter
            # Type Text Action    ${each_demand_code}
            # Capture Screenshot
            # Press Keys Action    enter
            # # Press Tab to move to the next field (if needed)
            # Press Keys Action    tab
            # Capture Screenshot
            # # Sleep    ${Min_time}

            # Type the current demand code and press enter
            Type Text Action    ${each_demand_code}
            Capture Screenshot
            Sleep    1    
            #Press tab for entering the value  
            Press Keys Action    tab
            Capture Screenshot
            Sleep    1  
            #Press ener key for moving to next jobcode input field
            Press Keys Action    enter
            Capture Screenshot  
            Sleep    1
        END            
        
        Capture Screenshot
            
    EXCEPT  AS  ${Suggested Jobs_Tab_error}
        Log  ${Suggested Jobs_Tab_error}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while interacting with the Suggested Jobs Tab.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        # Exit For Loop
        
    END       

Step Details Tab
    [Arguments]  ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}   
    
    TRY
        ${Service Type Description_value}    Clean String    ${Service Type Description}
        
        IF    "${Service Type Description_value}" == "bodyrepair" or "${Service Type Description_value}" == "bandp"

            # Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f5
            # Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f5
            #  Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f12
            Press Keys Action    f12

        ELSE
            # Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f9
            # Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f11
            Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f7
        
        END

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${working_btn}    ${sik_max_time}
        ${working_btn_exists}=    SikuliLibrary.Exists    ${working_btn}
        IF    ${working_btn_exists} == ${True}            
            SikuliLibrary.Right Click    ${working_btn}
            Sleep    ${Min_time}
            Press Keys Action    tab  
            Press Keys Action    enter  
            Sleep    ${Min_time}   
            Press Keys Action    enter               
            # Click Action    path:${check_confirm}                   
        END  
        Capture Screenshot
        
    
    EXCEPT  AS  ${Step_Details_Tab_error}
        Log    ${Step_Details_Tab_error}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while interacting with the Step Details Tab.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        # Exit For Loop
    END

Final Save
    [Arguments]  ${Job Card No.}    ${Input_Sheet_Path}
    TRY

        #Finale Save click
        Click Action     path:${final_save}
        # # Sleep    ${Min_time}   
        # #fetching the popup header content  
        # ${save_confirm_message}    Get Text Action    path:1 

        #-------------------------------------------------------------------------------------------------------#
        ${count}    Set Variable    1    
        ${MAX_ATTEMPTS}    Set Variable    3
        
        WHILE    ${count} < ${MAX_ATTEMPTS}
                    
            ${save_confirm_message}    Get Text Action    path:1
            IF   """${save_confirm_message}""" == """Transaction saved."""  
                BREAK
            ELSE
                Sleep    ${Min_time}
                ${count}=    Evaluate    ${count} + 1
            END
        END
        #-------------------------------------------------------------------------------------------------------#
      
        #Transacion Saving Options
        IF   """${save_confirm_message}""" == """Transaction saved."""  
                   
            ${confirm_message_content}    Get Attribute Action    path:1|3    Name    
            log    ${confirm_message_content}
            Capture Screenshot
            Run Keyword And Ignore Error    Click Action    path:1|1
            RETURN   ${True}     ${confirm_message_content}
 
        ELSE IF  """${save_confirm_message}""" == """Error"""
           
            ${confirm_message_content}    Get Attribute Action    path:1|3    Name    
            log    ${confirm_message_content}
            Capture Screenshot
            Run Keyword And Ignore Error    Click Action    path:1|1
            RETURN   ${False}     ${confirm_message_content}
 
        ELSE IF  """${save_confirm_message}""" == """Exception"""            
             
            ${confirm_message_content}    Get Attribute Action    path:1|3    Name    
            log    ${confirm_message_content}
            Capture Screenshot
            Run Keyword And Ignore Error    Click Action    path:1|1|1|1|2|4
            RETURN    ${False}    ${confirm_message_content}

        # ELSE IF  """${save_confirm_message}""" == """Message"""            
             
        #     ${confirm_message_content}    Get Attribute Action    path:1|3    Name    
        #     log    ${confirm_message_content}
        #     Capture Screenshot
        #     Run Keyword And Ignore Error    Click Action    path:1|1|1|1|2|4
        #     RETURN    ${False}    ${confirm_message_content}

        ELSE IF  """${save_confirm_message}""" == """Close"""
           
            Run Keyword And Ignore Error    Click Action    path:1|2
            # Sleep    ${Min_time}
            ${save_confirm_message}    Get Text Action    path:1
            Capture Screenshot
        
        ELSE IF  """${save_confirm_message}""" == """Message""" 
        
            ${confirm_message_content}    Get Attribute Action    path:1|3    Name    
            log    ${confirm_message_content}
            Capture Screenshot
            #added 963 line extra and commented very next
            
            IF    """${confirm_message_content}""" == """Present Odometer Reading Should not be less than Previous Reading"""
                
                #for handling above odometer bug popup
                Run Keyword And Ignore Error    Click Action    path:1|1 
                Sleep    1
                #waiting for transaction saved popup
                ${saved_status}    ${saved_message}    Transaction Saving Popup
                IF    """${saved_status}""" == """${True}"""
                    RETURN   ${True}     ${saved_message}
                ELSE
                    RETURN    ${False}    ${saved_message}
                END

            ELSE
                # Run Keyword And Ignore Error    Click Action    path:1|1|1|1|2|4
                Run Keyword And Ignore Error    Click Action    path:1|1
                RETURN    ${True}    ${confirm_message_content}
            END           

        ELSE  
            ${confirm_message_content}    Get Attribute Action    path:1|3    Name    
            log    ${confirm_message_content}
            Capture Screenshot
            #added 963 line extra and commented very next
            
            IF    """${confirm_message_content}""" == """Present Odometer Reading Should not be less than Previous Reading"""
                
                #for handling above odometer bug popup
                Run Keyword And Ignore Error    Click Action    path:1|1 
                Sleep    1
                #waiting for transaction saved popup
                ${saved_status}    ${saved_message}    Transaction Saving Popup
                IF    """${saved_status}""" == """${True}"""
                    RETURN   ${True}     ${saved_message}
                ELSE
                    RETURN    ${False}    ${saved_message}
                END

            ELSE
                # Run Keyword And Ignore Error    Click Action    path:1|1|1|1|2|4
                Run Keyword And Ignore Error    Click Action    path:1|1
                RETURN    ${True}    ${confirm_message_content}
            END           
        
        END    
 
    EXCEPT  AS  ${Save_error}
        Log    ${Save_error}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while saving the data    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        # Exit For Loop
    END

Transaction Saving Popup
    ${confirm_message_content}    Get Attribute Action    path:1|3    Name    
    log    ${confirm_message_content}
    Capture Screenshot
    Run Keyword And Ignore Error    Click Action    path:1|1
    RETURN   ${True}     ${confirm_message_content}



Additional Windows  
    ${msg_content}    Get Attribute Action    path:1|3    Name
    Log    ${msg_content}
    # Show Message Box    title   ${msg_content}
    IF    """${msg_content}""" == """Vehicle have no appointments"""
        Click Action    path:1|1
    ELSE IF  """${msg_content}""" == """Present Odometer Reading Should not be less than Previous Reading"""
        Click Action    path:1|1
    ELSE IF  """${msg_content}""" == """Dashboard""" 
        #Click Action    path:1|1 
        Log    dashboard

    END       


Used Jobcard

    ${used_jc_msg_content}    Get Attribute Action    path:1|3    Name
    Log    ${used_jc_msg_content} 
    # Show Message Box    title    ${used_jc_msg_content} 
    ${additional_window_check}   ERP_methods.Contains Substring    ${used_jc_msg_content}    ${sub_string}
    IF    ${additional_window_check} == True
            Click Action    path:1|1 
            RETURN  ${sub_string} 
    
    ELSE IF  """${used_jc_msg_content}""" == """This Service Type already used on this chassis number""" 
        Click Action    path:1|1 
        RETURN    ${used_jc_msg_content}       
        # ${opened_JC_msg_content}=    Get Attribute Action    path:1|3    Name
    ELSE IF  """${used_jc_msg_content}""" == """Present Odometer Reading Should not be less than Previous Reading"""
        Click Action    path:1|1

    ELSE IF    """${used_jc_msg_content}""" == """Vehicle have no appointments"""
        Click Action    path:1|1

    ELSE IF  """${used_jc_msg_content}""" == """Password has expired, please contact administrator or super user.""" 
        Click Action    path:1|1 
        RETURN    ${used_jc_msg_content}

    ELSE IF  """${used_jc_msg_content}""" == """Dashboard""" 
        #Click Action    path:1|1 
        Log    dashboard
    ELSE
        Click Action    path:1|1 
        RETURN    ${used_jc_msg_content}                 
    END
   
        
Password Expiry

        ${used_jc_msg_content}    Get Attribute Action    path:1|3    Name
        Log    ${used_jc_msg_content} 
        # Show Message Box    title    ${used_jc_msg_content} 

        IF  """${used_jc_msg_content}""" == """Password has expired, please contact administrator or super user.""" 
            Click Action    path:1|1 
            RETURN    ${used_jc_msg_content}
        END




Merge And Copy Consolidated Report To Results Folder
    [Arguments]    ${Merged_report_path}
    TRY
        ${project_root}     Set Variable    E:\\SFTP\\Recall
        # E:\SFTP\Recall\2025-03-28\Results
        ${results_folder}    Set Variable     ${project_root}\\${curr_date}\\${results_dir}       
        ${consolidated_report_path_in_results_dir}    ERP_methods.copy_report_to_destination_folder    ${Merged_report_path}    ${results_folder}

    EXCEPT  AS   ${error_message}          
        log    ${error_message}
        Capture Screenshot
        Fail    ${error_message}
    END
    [Return]    ${consolidated_report_path_in_results_dir}