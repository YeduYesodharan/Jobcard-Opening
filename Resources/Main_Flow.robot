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
Resource            Vehicle_Details_Search.robot
Resource            Jobcard_Navigations.robot
Resource            Validation_in_Vehicle_Search.robot
Resource            utility.robot
Resource    Pickup_Details_Tab.robot

*** Variables ***
${Search_get_Vehicle_details}    3|1|1|1|1|1|1|1|5|1
${Reg_No}    1|2|1|1
${Chassis_No}    1|2|1|2
${Mobile_no}    1|2|1|4

${no_record_alert}    1|1|3
${no_rec_ok_button}    1|1|1
${no_rec_header}    1|1|2
${no_rec_close_button}    1|1|2|1
${ok_to_tabs}    1|3|2
${chassis_creation_btn}    1|2|1|7
${create_chssis_chckbox}    1|1|1|1|1|1
${alert_box}    1|1

${testdb_Submit}    1|1|1|1|1|2|3|2
${testdb_hyperlink}    2|1|1|2|2|1|1|2|1|1
${Branch_path}    3|1|1|1|1|1|1|1|3|2|1|1  
${Location_path}   3|1|1|1|1|1|1|1|4|2|1|1
${close_erp_window_approval}   1|1
${imagerootfolder}            ${CURDIR}${/}..\\Locators
${current_month_button}          ${imagerootfolder}\\current_month_button.png
${jobcard_button}          ${imagerootfolder}\\jobcard_btn.png
${service_jobcard_button}          ${imagerootfolder}\\service_jc_btn.png
${transaction_button}        ${imagerootfolder}\\menu_transaction_btn.png
${testdb_choose_button}    ${imagerootfolder}\\testdb_choose_btn.png
${Service_Jobcard_Heading}    ${imagerootfolder}\\Service_Jobcard_Heading.png
${Bodyshop Jobcard Title}    ${imagerootfolder}\\Bodyshop Jobcard Title.png
${Testdb_box}    ${imagerootfolder}\\Testdb_box.png
${prod_choose_btn}             ${imagerootfolder}\\prod_choose_btn.png
${tools}    2|6
${>testdb}    1|1|1|2|2|1|1|2|1|1
${voucher_type}    3|1|1|1|1|1|1|1|1|2|1|1
${ERP_Password_Path}       ${CURDIR}${/}..\\Config\\Popular_Credentials.xlsx
# ${log_folder}     ${CURDIR}${/}..\\Log
# ${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
${log_folder}     C:\\JobcardOpeningIntegrated\\Screenshot
${wings_title_img}             ${imagerootfolder}\\wings_title_img.png
${uname_input}    1|1|1|1|1|2|2|6|1 
${pword_input}    1|1|1|1|1|2|2|8|1
${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx

*** Keywords ***

ERP Credentials 
    # [Arguments]    ${Input_Sheet_Path}
    TRY
        # ${file_status}    ERP_methods.Check File Exists    ${Input_Sheet_Path}
        # IF    '${file_status}' == '${False}'
        #     ERP_methods.Show Message Box    Alert    DMS Output Sheet Doesnt Exist.Update Sheet and Rerun the Bot after Closing Windows 
        #     Fail
        # ELSE
        # IF    '${Input_Sheet_Path}' != '${EMPTY}'
        #     ERP_methods.Extract And Correct Registration    ${Input_Sheet_Path}
        # END        

        # END
        Log    ${ERP_Password_Path}
        ${file_status}    ERP_methods.Check File Exists    ${ERP_Password_Path}
        
        IF    '${file_status}' == '${True}'
            Open Workbook   ${ERP_Password_Path}
            ${credential_data}=    Read Worksheet As Table   header=${True} 
            Sleep    ${Min_time}
            
            FOR    ${row}    IN    @{credential_data}
                ${uname}    Set Variable    ${row}[ERP Username]
                ${pword}    Set Variable    ${row}[ERP Password] 
                # ${uname}    Set Variable    ${row}[Testdb Username]
                # ${pword}    Set Variable    ${row}[Testdb Password]  
                ${profile}    Set Variable    ${row}[ERP Profile Name]

            END

            IF    '${pword}' != '${None}' and '${uname}' != '${None}' and '${profile}' != '${None}'
                RETURN    ${pword}    ${uname}    ${profile}
                Close Workbook
            ELSE
                ERP_methods.Show Message Box    Alert    Password Error.Update and Rerun the Bot after Closing Windows 
                Fail          
            END
            
        ELSE
            ERP_methods.Show Message Box    Alert    Credential Excel Doesnt Exist.Update and Rerun the Bot after Closing Windows 
            Fail
        END
            
        
    EXCEPT  AS   ${error_message}
        Log    ${error_message}
        Fail       
    END
        
ERP Login 
    # [Arguments]     ${Input_Sheet_Path}
    
    TRY
        ${pword1}  ${uname1}  ${profile}    ERP Credentials    
        IF    '${pword1}' != '${None}' and '${uname1}' != '${None}' and '${profile}' != '${None}'
            ERP_methods.Open Application
            Sleep    ${average_sleep}
            # Click Action    name:"${DB_Test}"
            Click Action    name:"${profile}"

            # Window Navigation    Wings ERP 23E - Web Client
            # Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${testdb_choose_button}    ${sik_time}
            # ${testdb_button_exists}=    SikuliLibrary.Exists    ${testdb_choose_button}
            # IF    ${testdb_button_exists}==${True}           
            #     Press Keys Action    enter 
            # ELSE 
            #     Press Keys Action    enter
            # END 
            
            # Sleep    ${average_sleep} 
            # #uname
            # Click Action    path:1|1|1|1|1|2|2|6|1
            # Press Keys  ctrl  a
            # Type Text Action    ${uname1}
            # #pword  
            # Click Action    path:1|1|1|1|1|2|2|8|1
            # Type Text Action    ${pword1}
            # Press Keys Action    enter

            
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${wings_title_img}    ${sik_time}
            ${wings_title_img_exists}=    SikuliLibrary.Exists    ${wings_title_img}
            IF    ${wings_title_img_exists}==${True} 
                Window Navigation    Wings ERP 23E - Web Client          
                Press Keys Action    enter            
            ELSE 
                Window Navigation    Wings ERP 23E - Web Client
                Press Keys Action    enter
            END 
            
            Sleep    ${average_sleep} 
            #uname
            Double Click Action    path:${uname_input}
            Type Text Action    ${uname1}
            #pword  
            Click Action    path:${pword_input}
            Type Text Action    ${pword1}
            Press Keys Action    enter

            # ${pw_status}    Password Expiry
            # # Show Message Box    title    ${used_jc_msg1}
            # Log    ${pw_status}
            # IF  '${pw_status}' != '${None}'
            #     IF  '${pw_status}' != 'Dashboard'
            #         RETURN  ${pw_status}
            #     END
            # END


            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
            ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
            IF    ${trans_button_exists}==${True}     
                # Window Navigation    Change Password
                # Run Keyword And Ignore Error    Click Action    path:2|1      
                RETURN    ${True}            
            ELSE
                RETURN    ${False}
            END  
        ELSE 
            ERP_methods.Show Message Box    Alert    Password/Username Error.Update and Rerun the Bot after Closing Windows 
            Fail
        END
        
    EXCEPT  AS  ${error_message}  
           
        Log    ${error_message} 
        Capture Screenshot 
        ERP_methods.Show Message Box    Message    Login Error.Retry Login after Closing Windows   
        Fail
           
    END


Jobcard Data Entry
    [Arguments]    ${Input_Sheet_Path}    ${jc_no}
    
    TRY

        Capture Screenshot
        ${row}=    Set Variable    None
        ${data}    Read Consolidated Report    ${Input_Sheet_Path}

        FOR    ${each_row}    IN    @{data}
            IF    '${each_row}[Job Card No]' == '${jc_no}'
                ${row}=    Set Variable    ${each_row}
                Exit For Loop
            END
        END
        Log     ${row}

        IF  '${row}[DMS Execution Status]' == '${DMS_Status1}' and '${row}[Recall Status]' == '${status_no}' and '${row}[ERP Execution Status]' != 'Fail' 
                    
            IF    """${row}[Service Type Code]""" == """PMS""" and """${row}[Sub Service Type]""" == """${None}"""
                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name} 
                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Blank Sub Service Type for PMS.    ${exception_reason_column_name}
            
            ELSE                 

                ${body_repair_value}    Clean String    ${body_repair}
                ${service_desc _value}    Clean String    ${row}[Service Type Description]
                #Choosing Bodyshop or Service Jobcard
                IF    '${service_desc _value}' == '${body_repair_value}' or '${service_desc _value}' == 'bandp'
                    Capture Screenshot             
                    ${nav_status1}    BodyShop Job Card Navigation   
                    Capture Screenshot
                    # ${nav_status1}    Production Bodyshop Jobcard Navigation
                    IF    '${nav_status1}' == '${True}'
                                        
                        ${status_val}    Bodyshop Branch and Location path    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Branch]
                        Capture Screenshot
                        IF    '${status_val}' == 'branch code not available'
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Branch not found in Branch Mapping File   ${exception_reason_column_name}
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                Press Keys Action    esc 
                                Press Keys Action    enter 
                                RETURN     branch code not available
                                # Continue For Loop
                        END
                        # Log    hi
                    ELSE       
                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Issue in Bodyshop Service Navigation   ${exception_reason_column_name}
                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                        Press Keys Action    esc 
                        Press Keys Action    enter  
                        RETURN      Issue in Bodyshop Service Navigation 
                    END

                ELSE                
                    ${nav_status2}    Service Job Card Navigation
                    Capture Screenshot
                    # ${nav_status2}    Production Service Jobcard Navigation
                    IF    '${nav_status2}' == '${True}'
                        
                        ${status_val}    Service Jobcard Branch and Location path    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Branch]
                        Capture Screenshot
                        IF    '${status_val}' == 'branch code not available'
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Branch not found in Branch Mapping File   ${exception_reason_column_name}
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                Press Keys Action    esc 
                                Press Keys Action    enter  
                                RETURN     branch code not available
                                # Continue For Loop
                        END
                    ELSE       
                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Issue in Regular Service Navigation   ${exception_reason_column_name}
                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                        Press Keys Action    esc 
                        Press Keys Action    enter  
                        RETURN     Issue in Regular Service Navigation
                    END
                END   

                #Data search with registration no.    
                TRY 
                    
                    ${alert_box_visible_1}    Vehicle Details Search with Registration Number    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}
                    Capture Screenshot
                    # ${alert_box_visible_1_lower}    Convert To Lower Case    ${alert_box_visible_1}
                
                EXCEPT  AS  ${jobcard_error}
                    RETURN     Vehicle Details Search with Registration Number Failed
                    # Continue For Loop
                END

                # IF  '${alert_box_visible_1_lower}' == 'message'
                IF  '${alert_box_visible_1}' != ''   # or '${alert_box_visible_1}' == 'Message'
                    Click Action  path:${no_rec_ok_button}
                    Click Action  path:${Chassis_No}            
                    Capture Screenshot
                #If data exist in registration no. check
                ELSE 

                    #Validation in Register no search

                    ${validation_status}    Validation in Registration Number Search    ${row}[Vehicle ID]    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description]
                    Capture Screenshot
                    # ${validation_status_stripped}=    Strip String    ${validation_status}
                    # ${vehicle_id_stripped}=          Strip String    ${row}[Vehicle ID]
                    # IF    '${validation_status_stripped}' == '${vehicle_id_stripped}'
                    IF    '${validation_status}' == '${row}[Vehicle ID]'

                        ${Vehicle_data_exist}    Run Keyword And Return Status    Click Action Maximum Retry  path:${ok_to_tabs}
                        Capture Screenshot
                    
                    ELSE 
                        #------------------------------------Retry for Vehicle Search----------------------------------------------------------#
                        ${validation_status}    Validation in Registration Number Search    ${row}[Vehicle ID]    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description]
                        Capture Screenshot
                        # ${validation_status_stripped}=    Strip String    ${validation_status}
                        # ${vehicle_id_stripped}=          Strip String    ${row}[Vehicle ID]
                        # IF    '${validation_status_stripped}' == '${vehicle_id_stripped}'
                        IF    '${validation_status}' == '${row}[Vehicle ID]'
                            ${Vehicle_data_exist}    Run Keyword And Return Status    Click Action Maximum Retry  path:${ok_to_tabs}
                            Capture Screenshot
                        
                        ELSE
                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Chassis Number mismatch in ERP Vehicle Search    ${exception_reason_column_name}
                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                            Press Keys Action    esc     #to close search window
                            Press Keys Action    esc 
                            Press Keys Action    enter 
                            RETURN    Chassis Number mismatch in ERP Vehicle Search
                            # Continue For Loop 
                        END
                        #------------------------------------Retry for Vehicle Search----------------------------------------------------------#             
                    END
                    IF    '${Vehicle_data_exist}' == '${True}'
                        TRY

                            ${used_jc_msgs}    JobCard Tab Entry Activation  ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 
                            # Show Message Box    title   ${used_jc_msgs}
                            Capture Screenshot
                            IF  '${used_jc_msgs}' != '${None}'
                                
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${used_jc_msgs}    ${exception_reason_column_name}
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                Press Keys Action    esc 
                                Press Keys Action    enter    
                                
                            ELSE

                                ${pickup_status}    Pickup Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 

                                IF  """${pickup_status}""" == """Some mandatory values except Pickup Driver is missing"""
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Some mandatory values except Pickup Driver is missing    ${exception_reason_column_name}
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                    Press Keys Action    esc 
                                    Press Keys Action    enter
                                # ELSE IF  """${pickup_status}""" == """Pick Up Details not updated in excel sheet"""
                                #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Pick Up Details not updated in excel sheet    ${exception_reason_column_name}
                                #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                #     Press Keys Action    esc 
                                #     Press Keys Action    enter
                                ELSE                               
                                        ${service_tab_exceptions}    Service Details Tab    ${row}[Job Card No]    ${row}[Promised vehicle delivery date and time]    ${row}[Service Type Description]      
                                        ...    ${row}[Name of Service Advisor]    ${row}[OMR]    ${row}[Demand Code]    ${row}[Service Type Code]    ${Input_Sheet_Path}    ${row}[registration no]
                                        ...    ${row}[Vehicle ID]    ${row}[Mobile]    ${row}[Service Type Description]    ${row}[Sub Service Type]    ${row}[Branch]  
                                        Capture Screenshot   
                                        IF    """${service_tab_exceptions}""" == """Odometer Exception Occured:DMS reading is higher than ERP reading"""                        
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Odometer Exception Occured:DMS reading is higher than ERP reading    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter
                                        
                                        ELSE IF  '${service_tab_exceptions}' == 'branch code not available'

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Branch not found in Branch Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  
                                        
                                        ELSE IF  '${service_tab_exceptions}' == 'odometer updation failed'

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    'odometer updation failed'   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  
                                        

                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for type of service'

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Type of Service not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  

                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for service type'  

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Type not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter 

                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for service advisor'  

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Advisor not found in Service Advisor Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter 


                                        ELSE IF  '${service_tab_exceptions}' == 'Jobcard already opened '

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                            
                                                                            
                                        ELSE IF  """${service_tab_exceptions}""" == """This Service Type already used on this chassis number"""
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    This Service Type already used on this chassis number    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                        ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 2nd FREE SERVICE as per the schedule."""

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                        ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 1st FREE SERVICE as per the schedule."""

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                        ELSE
                                            Suggested Jobs Tab    ${row}[Job Card No]   ${row}[Demand Code]     ${Input_Sheet_Path}    ${row}[Service Type Description]     
                                            Step Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 
                                             #just commented--------------------------------#
                                            # ${save_status}    ${save_status_message}    Final Save    ${row}[Job Card No]    ${Input_Sheet_Path}
                                            # Log    ${save_status_message}
                                            # IF    ${save_status} == ${True}
                                            # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                            # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Success    ${execution_status_column_name} 
                                            # Press Keys Action    esc 
                                            # Press Keys Action    enter
                                            # ELSE

                                            #     update_execution_status_in_summary_report    ${Input_Sheet_Path}     ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                            #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            #     Press Keys Action    esc 
                                            #     Press Keys Action    enter
                                            # END
                                             #just commented--------------------------------#
                                
                                        END 

                                END                        
                            END                                
                                    
                        EXCEPT  AS  ${jobcard_error}
                            Log    ${jobcard_error}
                            Press Keys Action    esc 
                            Sleep    ${Min_time}
                            Press Keys Action    enter
                            RETURN    ${False}
                            # Continue For Loop
                        END
                    END
                    RETURN    ${False}
                    # Continue For Loop
                    
                END  

                #Data search with chassis no.
                TRY

                    ${alert_box_visible_2}    Vehicle Details Search with Chassis Number    ${row}[Vehicle ID]    ${row}[Job Card No]    ${Input_Sheet_Path}
                    # ${alert_box_visible_2_lower}    Convert To Lower Case    ${alert_box_visible_2}
                
                EXCEPT  AS  ${jobcard_error}
                    # Continue For Loop
                    RETURN    Vehicle Details Search with Chassis Number Failed
                END


                # IF  '${alert_box_visible_2_lower}' == 'message' 
                IF  '${alert_box_visible_2}' != ''  #   or '${alert_box_visible_2}' == 'Message' 
                    Click Action  path:${no_rec_ok_button}   
                    Click Action  path:${Mobile_no}           
                
                #Data exist with chassis number check
                ELSE
                    #Validation in Chassis number

                    ${validation_status}    Validation in Chassis Number Search    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description]
                    # ${validation_status_stripped}=    Strip String    ${validation_status}
                    # ${reg_no_stripped}=          Strip String    ${row}[registration no]
                    # IF    '${validation_status_stripped}' == '${reg_no_stripped}'    
                    IF    '${validation_status}' == '${row}[registration no]'

                        ${Vehicle_data_exist}    Run Keyword And Return Status    Click Action Maximum Retry  path:${ok_to_tabs}
                    
                    ELSE 
                        #------------------------------------Retry for Vehicle Search----------------------------------------------------------#

                        ${validation_status}    Validation in Chassis Number Search    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description]
                        # ${validation_status_stripped}=    Strip String    ${validation_status}
                        # ${reg_no_stripped}=          Strip String    ${row}[registration no]
                        # IF    '${validation_status_stripped}' == '${reg_no_stripped}'
                        IF    '${validation_status}' == '${row}[registration no]'

                            ${Vehicle_data_exist}    Run Keyword And Return Status    Click Action Maximum Retry  path:${ok_to_tabs}
                        
                        ELSE
                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Registration Number mismatch in ERP Vehicle Search    ${exception_reason_column_name}
                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                            Press Keys Action    esc     #to close search window
                            Press Keys Action    esc 
                            Press Keys Action    enter 
                            RETURN     Registration Number mismatch in ERP Vehicle Search 
                            # Continue For Loop 
                        END
                        #------------------------------------Retry for Vehicle Search----------------------------------------------------------#             
                    END                 
                    
                    IF    '${Vehicle_data_exist}' == '${True}'
                        TRY
                
                            ${used_jc_msgs}    JobCard Tab Entry Activation  ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 
                            # Show Message Box    title   ${used_jc_msgs}
                            IF  '${used_jc_msgs}' != '${None}'
                                
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${used_jc_msgs}    ${exception_reason_column_name}
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                Press Keys Action    esc 
                                Press Keys Action    enter    
                                
                            ELSE

                                ${pickup_status}    Pickup Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 

                                IF  """${pickup_status}""" == """Some mandatory values except Pickup Driver is missing"""
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Some mandatory values except Pickup Driver is missing    ${exception_reason_column_name}
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                    Press Keys Action    esc 
                                    Press Keys Action    enter
                                # ELSE IF  """${pickup_status}""" == """Pick Up Details not updated in excel sheet"""
                                #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Pick Up Details not updated in excel sheet    ${exception_reason_column_name}
                                #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                #     Press Keys Action    esc 
                                #     Press Keys Action    enter
                                ELSE                  
                                        ${service_tab_exceptions}    Service Details Tab    ${row}[Job Card No]    ${row}[Promised vehicle delivery date and time]    ${row}[Service Type Description]      
                                        ...    ${row}[Name of Service Advisor]    ${row}[OMR]    ${row}[Demand Code]    ${row}[Service Type Code]    ${Input_Sheet_Path}    ${row}[registration no]
                                        ...    ${row}[Vehicle ID]    ${row}[Mobile]    ${row}[Service Type Description]    ${row}[Sub Service Type]    ${row}[Branch]   
                                        IF    """${service_tab_exceptions}""" == """Odometer Exception Occured:DMS reading is higher than ERP reading"""                        
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Odometer Exception Occured:DMS reading is higher than ERP reading    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter
                            
                                        ELSE IF  '${service_tab_exceptions}' == 'branch code not available'

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Branch not found in Branch Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  
                                        
                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for type of service'

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Type of Service not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  

                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for service type'  

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Type not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter 
                                        
                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for service advisor'  

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Advisor not found in Service Advisor Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter 


                                        ELSE IF  '${service_tab_exceptions}' == 'Jobcard already opened '

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter
                                        
                                        
                                        ELSE IF  """${service_tab_exceptions}""" == """This Service Type already used on this chassis number"""
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    This Service Type already used on this chassis number    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                        ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 2nd FREE SERVICE as per the schedule."""

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                        ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 1st FREE SERVICE as per the schedule."""

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                        ELSE
                                            Suggested Jobs Tab    ${row}[Job Card No]   ${row}[Demand Code]     ${Input_Sheet_Path}    ${row}[Service Type Description]     
                                            Step Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description]  
                                             #just commented--------------------------------#
                                            # ${save_status}    ${save_status_message}    Final Save    ${row}[Job Card No]    ${Input_Sheet_Path}
                                            # Log    ${save_status_message}
                                            # IF    ${save_status} == ${True}
                                            # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                            # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Success    ${execution_status_column_name}    
                                            # Press Keys Action    esc 
                                            # Press Keys Action    enter
                                            # ELSE

                                            #     update_execution_status_in_summary_report    ${Input_Sheet_Path}     ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                            #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            #     Press Keys Action    esc 
                                            #     Press Keys Action    enter
                                            # END 
                                             #just commented--------------------------------#  
                                        END

                                END
                            END

                        EXCEPT  AS  ${jobcard_error}
                            Press Keys Action    esc 
                            Press Keys Action    enter
                            RETURN     ${False} 
                            # Continue For Loop
                        END
                    END
                    # Continue For Loop
                    RETURN     ${False}
                    
                END  

                #Data search with mobile no.
                TRY

                    ${alert_box_visible_3}    Vehicle Details Search with Mobile Number    ${row}[Mobile]    ${row}[Job Card No]    ${Input_Sheet_Path}
                    # ${alert_box_visible_3_lower}    Convert To Lower Case    ${alert_box_visible_3}

                EXCEPT  AS  ${jobcard_error}
                    RETURN     Vehicle Details Search with Mobile Number Failed
                END

                #Create Chassis and Customer details if no data in Register number,chassis and mobile number check
                # IF  '${alert_box_visible_3_lower}' == 'message' 
                IF  '${alert_box_visible_3}' != ''     #or '${alert_box_visible_3}' == 'Message' 
                    Click Action  path:${no_rec_ok_button} 

                    IF  '${row}[Vehicle Sales Date]' != '${None}'

                        TRY
                            ${init_status}    New Chassis Creation Initiaion  ${row}[Vehicle ID]   ${row}[Vehicle Model]   ${row}[Vehicle Variant]   ${row}[ENGINE NUM]   
                            ...    ${row}[Color]   ${row}[Vehicle Sales Date]   ${row}[Customer Name]   ${row}[State]   ${row}[Address 1]   ${row}[City]   ${row}[Mobile]
                            ...    ${row}[Job Card No]    ${Input_Sheet_Path}
                            IF    '${init_status}' == '${True}'
                                New Chassis Creation Checkbox Window  ${row}[Vehicle ID]   ${row}[Vehicle Model]   ${row}[Vehicle Variant]   ${row}[ENGINE NUM]   
                                ...    ${row}[Color]   ${row}[Vehicle Sales Date]   ${row}[Customer Name]   ${row}[State]   ${row}[Address 1]   ${row}[City]   ${row}[Mobile]
                                ...    ${row}[Job Card No]    ${Input_Sheet_Path} 
                            END                                
                            ${chassis_status}  ${chassis_status_message}  New Chassis Creation Data Window  ${row}[Vehicle ID]   ${row}[Vehicle Model]   ${row}[Vehicle Variant]   ${row}[ENGINE NUM]   
                            ...    ${row}[Color]   ${row}[Vehicle Sales Date]   ${row}[Customer Name]   ${row}[State]   ${row}[Address 1]   ${row}[City]   ${row}[Mobile]
                            ...    ${row}[Job Card No]    ${Input_Sheet_Path}
                            
                            IF    ${chassis_status} == ${True}
                                Customer Creation Checkbox Window    ${row}[Customer Name]    ${row}[State]    ${row}[Address 1]    ${row}[City]   
                                ...    ${row}[Mobile]    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}
                                Customer Data Entry Window    ${row}[Customer Name]    ${row}[State]    ${row}[Address 1]    ${row}[City]   
                                ...    ${row}[Mobile]    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}
                                Search Customer Window    ${row}[Customer Name]    ${row}[State]    ${row}[Address 1]    ${row}[City]   
                                ...    ${row}[Mobile]    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}
                                ${customer_status}  ${customer_status_message}  New Customer Data Entry    ${row}[Customer Name]    ${row}[State]    ${row}[Address 1]    ${row}[City]   
                                ...    ${row}[Mobile]    ${row}[registration no]    ${row}[Job Card No]    ${row}[Pin]    ${Input_Sheet_Path}
                                IF    ${customer_status} == ${True}

                                    ${Update_status}    Update Registraion    ${row}[Customer Name]    ${row}[State]    ${row}[Address 1]    ${row}[City]   
                                    ...    ${row}[Mobile]    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}
                                ELSE
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${customer_status_message}    ${exception_reason_column_name}                                
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                    Press Keys Action    esc 
                                    Press Keys Action    enter
                                    Continue For Loop
                                END
                                
                            ELSE IF  ${chassis_status} == ${None}

                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    No match found for vehicle service model    ${exception_reason_column_name}
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                    Press Keys Action    esc 
                                    Press Keys Action    esc 
                                    Press Keys Action    esc 
                                    Press Keys Action    enter
                                    # Continue For Loop
                                    RETURN    No match found for vehicle service model
                                                    
                            ELSE

                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${chassis_status_message}    ${exception_reason_column_name}
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                Press Keys Action    esc 
                                Press Keys Action    enter
                                Continue For Loop
                                
                            END
                                                    
                            IF    '${Update_status}' == '${True}'

                                ${used_jc_msgs}    JobCard Tab Entry Activation  ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 
                                # Show Message Box    title   ${used_jc_msgs}
                                IF  '${used_jc_msgs}' != '${None}'
                                    
                                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${used_jc_msgs}    ${exception_reason_column_name}
                                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                        Press Keys Action    esc 
                                        Press Keys Action    enter    
                                    
                                ELSE

                                    ${pickup_status}    Pickup Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 

                                    IF  """${pickup_status}""" == """Some mandatory values except Pickup Driver is missing"""
                                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Some mandatory values except Pickup Driver is missing    ${exception_reason_column_name}
                                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                        Press Keys Action    esc 
                                        Press Keys Action    enter
                                    # ELSE IF  """${pickup_status}""" == """Pick Up Details not updated in excel sheet"""
                                    #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Pick Up Details not updated in excel sheet    ${exception_reason_column_name}
                                    #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                    #     Press Keys Action    esc 
                                    #     Press Keys Action    enter
                                    ELSE  

                                            ${service_tab_exceptions}    Service Details Tab    ${row}[Job Card No]    ${row}[Promised vehicle delivery date and time]    ${row}[Service Type Description]      
                                            ...    ${row}[Name of Service Advisor]    ${row}[OMR]    ${row}[Demand Code]    ${row}[Service Type Code]    ${Input_Sheet_Path}    ${row}[registration no]
                                            ...    ${row}[Vehicle ID]    ${row}[Mobile]    ${row}[Service Type Description]    ${row}[Sub Service Type]    ${row}[Branch]   
                                            IF    """${service_tab_exceptions}""" == """Odometer Exception Occured:DMS reading is higher than ERP reading"""                        
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Odometer Exception Occured:DMS reading is higher than ERP reading    ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter

                                            ELSE IF  '${service_tab_exceptions}' == 'branch code not available'

                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Branch not found in Branch Mapping File   ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter  
                                            
                                            ELSE IF  '${service_tab_exceptions}' == 'no match found for type of service'

                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Type of Service not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter  

                                            ELSE IF  '${service_tab_exceptions}' == 'no match found for service type'  

                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Type not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter 

                                            ELSE IF  '${service_tab_exceptions}' == 'no match found for service advisor'  

                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Advisor not found in Service Advisor Mapping File   ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter 

                                            ELSE IF  '${service_tab_exceptions}' == 'Jobcard already opened '

                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter
                                            
                                            
                                            
                                            ELSE IF  """${service_tab_exceptions}""" == """This Service Type already used on this chassis number"""
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    This Service Type already used on this chassis number    ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter
                                            
                                            ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 2nd FREE SERVICE as per the schedule."""

                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter

                                            ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 1st FREE SERVICE as per the schedule."""

                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                Press Keys Action    esc 
                                                Press Keys Action    enter
                                            
                                            ELSE
                                                Suggested Jobs Tab    ${row}[Job Card No]   ${row}[Demand Code]     ${Input_Sheet_Path}    ${row}[Service Type Description]     
                                                Step Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 
                                                 #just commented--------------------------------#
                                                ${save_status}    ${save_status_message}        ${row}[Job Card No]    ${Input_Sheet_Path}
                                                Log    ${save_status_message}
                                                IF    ${save_status} == ${True}
                                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Success    ${execution_status_column_name} 
                                                    Press Keys Action    esc 
                                                    Press Keys Action    enter
                                                ELSE

                                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}     ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                                    Press Keys Action    esc 
                                                    Press Keys Action    enter
                                                END
                                                 #just commented--------------------------------#
                                            END
                                    END        
                                    
                                END
                                    
                                
                                
                            ELSE  

                                update_execution_status_in_summary_report    ${Input_Sheet_Path}     ${row}[Job Card No]    ${Update_status}    ${exception_reason_column_name}
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                Press Keys Action    esc
                                Press Keys Action    enter
                                Continue For Loop
                            END
                        
                        EXCEPT  AS  ${chassis_customer_creation_error}                       
                            # Continue For Loop
                            RETURN    ${False}
                        END
                    ELSE
                        update_execution_status_in_summary_report    ${Input_Sheet_Path}     ${row}[Job Card No]    DMS Vehicle Sales Date is Empty    ${exception_reason_column_name}
                        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                        Press Keys Action    esc
                        Press Keys Action    enter
                        Continue For Loop
                    END
                    # Continue For Loop
                    RETURN    ${False}

                #If data exist in the mobile number check
                ELSE    #''value

                    #data validation in mobile number search
                    ${validation_chassis}    ${validation_reg}    Validation in Mobile Number Search    ${row}[Vehicle ID]    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description]
                    # ${validation_chassis_stripped}=    Strip String    ${validation_chassis}
                    # ${validation_reg_stripped}=    Strip String    ${validation_reg}
                    # ${reg_stripped}=    Strip String    ${row}[registration no]
                    # ${chassis_stripped}=    Strip String    ${row}[Vehicle ID]
                    

                    # IF    '${validation_chassis_stripped}' == '${chassis_stripped}' and '${validation_reg_stripped}' == '${reg_stripped}'
                    IF    '${validation_chassis}' == '${row}[Vehicle ID]' and '${validation_reg}' == '${row}[registration no]'

                        ${Vehicle_data_exist}    Run Keyword And Return Status    Click Action Maximum Retry  path:${ok_to_tabs}
                    
                    ELSE 
                        #------------------------------------Retry for Vehicle Search----------------------------------------------------------#
                            ${validation_chassis}    ${validation_reg}    Validation in Mobile Number Search    ${row}[Vehicle ID]    ${row}[registration no]    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description]
                            # ${validation_chassis_stripped}=    Strip String    ${validation_chassis}
                            # ${validation_reg_stripped}=    Strip String    ${validation_reg}
                            # ${reg_stripped}=    Strip String    ${row}[registration no]
                            # ${chassis_stripped}=    Strip String    ${row}[Vehicle ID]
                    

                            # IF    '${validation_chassis_stripped}' == '${chassis_stripped}' and '${validation_reg_stripped}' == '${reg_stripped}'
                            IF    '${validation_chassis}' == '${row}[Vehicle ID]' and '${validation_reg}' == '${row}[registration no]'
                        
                                ${Vehicle_data_exist}    Run Keyword And Return Status    Click Action Maximum Retry  path:${ok_to_tabs}   
                            
                            ELSE
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Registration Number/Chassis Number mismatch in ERP Vehicle Search    ${exception_reason_column_name}
                                update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                Press Keys Action    esc     #to close search window
                                Press Keys Action    esc 
                                Press Keys Action    enter 
                                # Continue For Loop      
                                RETURN    Registration Number/Chassis Number mismatch in ERP Vehicle Search   
                            END 
                        #------------------------------------Retry for Vehicle Search----------------------------------------------------------#    
                    END      
                        
                    
                    IF    '${Vehicle_data_exist}' == '${True}'
                        TRY
                            
                            ${used_jc_msgs}    JobCard Tab Entry Activation  ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 
                            # Show Message Box    title   ${used_jc_msgs}
                            IF  '${used_jc_msgs}' != '${None}'
                                
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${used_jc_msgs}    ${exception_reason_column_name}
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                    Press Keys Action    esc 
                                    Press Keys Action    enter    
                                
                            ELSE
                                ${pickup_status}    Pickup Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 

                                IF  """${pickup_status}""" == """Some mandatory values except Pickup Driver is missing"""
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Some mandatory values except Pickup Driver is missing    ${exception_reason_column_name}
                                    update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                    Press Keys Action    esc 
                                    Press Keys Action    enter
                                # ELSE IF  """${pickup_status}""" == """Pick Up Details not updated in excel sheet"""
                                #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Pick Up Details not updated in excel sheet    ${exception_reason_column_name}
                                #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                #     Press Keys Action    esc 
                                #     Press Keys Action    enter
                                ELSE                              
                                        ${service_tab_exceptions}    Service Details Tab    ${row}[Job Card No]    ${row}[Promised vehicle delivery date and time]    ${row}[Service Type Description]      
                                        ...    ${row}[Name of Service Advisor]    ${row}[OMR]    ${row}[Demand Code]    ${row}[Service Type Code]    ${Input_Sheet_Path}    ${row}[registration no]
                                        ...    ${row}[Vehicle ID]    ${row}[Mobile]    ${row}[Service Type Description]    ${row}[Sub Service Type]    ${row}[Branch]   
                                        IF    """${service_tab_exceptions}""" == """Odometer Exception Occured:DMS reading is higher than ERP reading"""                        
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Odometer Exception Occured:DMS reading is higher than ERP reading    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter

                                        ELSE IF  '${service_tab_exceptions}' == 'branch code not available'

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Branch not found in Branch Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  

                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for type of service'

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For DMS Type of Service not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  

                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for service type'  

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Type not found in Service Type Cumulative Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter 
                                        
                                        ELSE IF  '${service_tab_exceptions}' == 'no match found for service advisor'  

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Mapping For Service Advisor not found in Service Advisor Mapping File   ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter 

                                        ELSE IF  '${service_tab_exceptions}' == 'Jobcard already opened '

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter  


                                        ELSE IF  """${service_tab_exceptions}""" == """This Service Type already used on this chassis number"""
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    This Service Type already used on this chassis number    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter
                                        
                                        ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 2nd FREE SERVICE as per the schedule."""

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter    

                                        ELSE IF  """${service_tab_exceptions}""" == """Kilometers not valid for 1st FREE SERVICE as per the schedule."""

                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Jobcard already opened    ${exception_reason_column_name}
                                            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            Press Keys Action    esc 
                                            Press Keys Action    enter                            
                                        
                                        ELSE
                                            Suggested Jobs Tab    ${row}[Job Card No]   ${row}[Demand Code]     ${Input_Sheet_Path}    ${row}[Service Type Description]     
                                            Step Details Tab    ${row}[Job Card No]    ${Input_Sheet_Path}    ${row}[Service Type Description] 
                                             #just commented--------------------------------#
                                            # ${save_status}    ${save_status_message}    Final Save    ${row}[Job Card No]    ${Input_Sheet_Path}
                                            # Log    ${save_status_message}
                                            # IF    ${save_status} == ${True}
                                            # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                            # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Success    ${execution_status_column_name}  
                                            # Press Keys Action    esc 
                                            # Press Keys Action    enter
                                            # ELSE

                                            #     update_execution_status_in_summary_report    ${Input_Sheet_Path}     ${row}[Job Card No]    ${save_status_message}    ${exception_reason_column_name}
                                            #     update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name}
                                            #     Press Keys Action    esc 
                                            #     Press Keys Action    enter
                                            # END
                                             #just commented--------------------------------#
                                        END 

                                END
                            END
                        EXCEPT  AS  ${jobcard_error}
                            Press Keys Action    esc 
                            Press Keys Action    enter
                            # Continue For Loop
                            RETURN    ${False}
                        END
                    END
                    # Continue For Loop
                    RETURN    ${False}
                    
                END  
                Sleep    ${Min_time}
            END

        ELSE
            Log    No entries to proceed with ERP
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    Fail    ${execution_status_column_name} 
            update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${row}[Job Card No]    There is no matching details found to proceed with ERP.    ${exception_reason_column_name}

        END
            
        # END

        # Close Workbook
        # Closing ERP Window    Wings ERP 23E - Web Client
        # ERP_methods.Close Application
        # Close Erp
    
    EXCEPT  AS  ${error_message}       
        Log    ${error_message}
        Fail    ${error_message}
    END



