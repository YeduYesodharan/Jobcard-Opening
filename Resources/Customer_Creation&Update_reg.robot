*** Settings ***
Library             RPA.Desktop
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             OperatingSystem
Library             String
Library             Collections
Library             RPA.Windows
Library             Libraries/business_operations.py  
Variables           Variables/variables.py 
Resource            Wrappers.robot

*** Variables ***
${chassis_creation_btn}    1|2|1|7
${create_chssis_chckbox}    1|1|1|1|1|1
${chassis_no_input}    1|1|1|1|1|4|1
${service_model_code_input}    1|1|1|1|1|6|1
${vehicle_input}    1|1|1|1|1|8|1
${engine_number_input}    1|1|1|1|1|12|1
${colour_input}    1|1|1|1|1|14|1
${Date_of_sale_input}    1|1|1|1|1|18|1
${create_chassis_save}    1|1|1|2|2|1
${create/update_customer_checkbox}    1|1|1|2|1|1
${customer_type_input}    1|1|1|1|2|1
${customer_title_input}    1|1|1|1|10|1
${country_input}    1|1|1|1|30|1
${Customer_master}    1|1|1|1|18|1 
${assessee_type}    1|1|1|1|22|1 
${Customer_name_input}    1|1|1|1|12|1
${State_input}    1|1|1|1|28|1
${so/do/wo}    1|1|1|1|31|1|2|1
${Address_1_input}    1|1|1|1|31|1|4|1
${City_input}    1|1|1|1|31|1|8|1
${pin_input}    1|1|1|1|31|1|10|1
${phn&mob_input}    1|1|1|1|32|1|2|1
${create_new_cust}    1|1|1|2|2|1 
${cust_submit_closebtn}    1|1|1|2|2|3
${search_cust_btn}    1|1|1|2|1|1
${search_custname_input}    1|1|1|1|1|1|1|1
${search_custmob_radio}    1|1|1|1|1|1|1|2
${search_custwith_details}    1|1|1|1|1|1|1|5|1
${search_cust_go}    1|1|1|1|1|1|1|6
${search_cust_window_ok}    1|1|1|1|1|3|2
${search_cust_window_close}    1|1|1|1|1|3|1
${update_regis_checkbox}    1|1|1|3|1|1
${regis_no_input}    1|1|1|3|1|7|1
${cust_type_title}    1|1|1|1|1
${cust_title_title}    1|1|1|1|9
${country_title}    1|1|1|1|29
${cust_name_title}    1|1|1|1|11
${state_title}    1|1|1|1|27
${sodowo_title}    1|1|1|1|31|1|1
${address1_title}    1|1|1|1|31|1|3
${city_title}    1|1|1|1|31|1|7
${pin_title}    1|1|1|1|31|1|9
${mobile_title}    1|1|1|1|32|1|1
${cust_create_alert_box}    1|1|1|1|1    #|2
${no_cust_rec_ok_btn}    1|1|1|1|1|1
${create_cust_close_btn}    1|1|1|2|2|3
${create_chassis_overall_save}    1|1|1|5|1
${close_to_tabs}    1|3|1
${create_chassis_1st_window_close}    1|3|1
${create_chassis_checkbox_win_close}    1|1|1|5|2
${overall_chassis_crate_close_btn}    1|1|1|5|2
${cust_created_ok}    1|1|1|1|1
${regis_cust_updated_ok}    1|1|1|1
${create_new_cust_btn}    1|1|1|3|2|1
${chassis_customer_creation_msg_box}    1|1|1|1|3
${log_folder}     ${CURDIR}${/}..\\Screenshot
# ${branch_mapping}     ${CURDIR}${/}..\\Mapping\\Location Mapping DMS ERP.xlsx
${branch_mapping}    C:\\JobcardOpeningIntegrated\\Mapping\\Location Mapping DMS ERP.xlsx

*** Keywords ***

Customer Creation Checkbox Window
    [Arguments]    ${Customer Name}    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}

    TRY
        Click Action    path:${create/update_customer_checkbox}   
         
    EXCEPT  AS  ${error_message}
        Log  ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred during the Customer Creation Checkbox Window processing.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    path:${create_chassis_checkbox_win_close}
        Click Action    path:${create_chassis_1st_window_close}
        # Exit For Loop
    END   

Customer Data Entry Window
    [Arguments]    ${Customer Name}    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}
    
    TRY
        Capture Screenshot
        Click Action    path:${search_cust_btn} 
    
    EXCEPT  AS  ${error_message}
        Log  ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred during the Customer Data Entry Window processing.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    path:${create_cust_close_btn}
        Click Action    path:${create_chassis_checkbox_win_close}
        Click Action    path:${create_chassis_1st_window_close}
        # Exit For Loop
    END   

Search Customer Window
    [Arguments]    ${Customer Name}    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}
    
    TRY
        Capture Screenshot
        # Click Action    path:${search_custname_input}   
        Click Action    path:${search_custmob_radio}    
        Click Action    path:${search_custwith_details}    
        Type Text Action    ${Phone & Mobile No.}
        # Type Text Action    9946032700       #2
        
        Click Action    path:${search_cust_go} 
        Sleep    ${avg_time}  
        ${create_cust_alert_visible}    Get Text Action    path:${cust_create_alert_box}
        Log    ${create_cust_alert_visible}
        Set Global Variable    ${create_cust_alert_visible}
        ${create_cust_alert_visible_lower}    Convert To Lower Case    ${create_cust_alert_visible}
        Set Global Variable    ${create_cust_alert_visible_lower}
        IF  '${create_cust_alert_visible_lower}' == 'message'
            Click Action  path:${no_cust_rec_ok_btn}
            Click Action    path:${search_cust_window_close}

        ELSE   #customer existing
            Click Action    path:${search_cust_window_close}
            Click Action    path:${create_cust_close_btn}
            Update Registraion    ${Customer Name}    ${State}    ${Address}    ${City}   ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}


        END
    EXCEPT  AS  ${error_message}
        Log  ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while searching for the customer.    ${exception_reason_column_name}
        Click Action    path:${search_cust_window_close}
        Click Action    path:${create_cust_close_btn}
        Click Action    path:${create_chassis_checkbox_win_close}
        Click Action    path:${create_chassis_1st_window_close}   
        # Exit For Loop
    END  

New Customer Data Entry
    [Arguments]    ${Customer Name}    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Pin}    ${Input_Sheet_Path}

    TRY

        IF  '${create_cust_alert_visible_lower}' == 'message' 

            ${cust_type_titlevalue}    Get Text Action    path:${cust_type_title}
            ${cust_type_titlevalue_lower}    Convert To Lower Case    ${cust_type_titlevalue}
            IF    '${cust_type_titlevalue_lower}' == "customer type"
                # Capture Screenshot
                Click Action    path:${customer_type_input}
                Type Text Action    ${Customer_type}
                Press Keys Action    enter
                Capture Screenshot   

                
            END

            ${cust_title_titlevalue}    Get Text Action    path:${cust_title_title}
            ${cust_title_titlevalue_lower}    Convert To Lower Case    ${cust_title_titlevalue}
            IF    '${cust_title_titlevalue_lower}' == "title" 
                # Capture Screenshot              
                Click Action    path:${customer_title_input}   
                Type Text Action    ${customer_title}
                Press Keys Action    enter   
                Capture Screenshot   
                
            END

            ${country_titlevalue}    Get Text Action    path:${country_title}
            ${country_titlevalue_lower}    Convert To Lower Case    ${country_titlevalue}
            IF    '${country_titlevalue_lower}' == "country"     
                # Capture Screenshot          
                Click Action    path:${country_input}
                Type Text Action    ${Country}  
                Press Keys Action    enter
                 Capture Screenshot   
            END    

            ${cust_name_titlevalue}    Get Text Action    path:${cust_name_title}
            ${cust_name_titlevalue_lower}    Convert To Lower Case    ${cust_name_titlevalue}
            IF    '${cust_name_titlevalue_lower}' == "customer name"   
                # Capture Screenshot            
                Click Action    path:${Customer_name_input} 
                Type Text Action    ${Customer Name}
                Capture Screenshot   
            END 

            ${sodowo_titlevalue}    Get Text Action    path:${sodowo_title}
            ${sodowo_titlevalue_lower}    Convert To Lower Case    ${sodowo_titlevalue}
            IF    '${sodowo_titlevalue_lower}' == "s/o. d/o. w/o." 
                # Capture Screenshot              
                Click Action    path:${so/do/wo}
                Type Text Action    ${SO_DO_WO}
                Capture Screenshot   
            END     
        
            
            ${state_titlevalue}    Get Text Action    path:${state_title}
            ${state_titlevalue_lower}    Convert To Lower Case    ${state_titlevalue}
            IF    '${state_titlevalue}' == "state"    
                # Capture Screenshot           
                Click Action    path:${State_input}
                Type Text Action    ${State}
                Press Keys Action    enter
                Capture Screenshot   
            END           
        
            ${city_titlevalue}    Get Text Action    path:${city_title}
            ${city_titlevalue_lower}    Convert To Lower Case    ${city_titlevalue}
            IF    '${city_titlevalue_lower}' == "city"   
                Capture Screenshot            
                Click Action    path:${City_input}
                Type Text Action    ${City}
            END
            
            ${Address1_titlevalue}    Get Text Action    path:${address1_title}
            ${Address1_titlevalue_lower}    Convert To Lower Case    ${Address1_titlevalue}
            IF    '${Address1_titlevalue_lower}' == "address 1" or '${Address1_titlevalue_lower}' == "address 2"  
                ${Address_data}    ERP_methods.Extract And Combine Addresses    ${Input_Sheet_Path} 
                Log    ${Address_data}
                Capture Screenshot          
                Click Action    path:${Address_1_input}
                Type Text Action    ${Address_data}
                Capture Screenshot
            END
    
            
            ${pin_titlevalue}    Get Text Action    path:${pin_title}
            ${pin_titlevalue_lower}    Convert To Lower Case    ${pin_titlevalue}
            IF    '${pin_titlevalue_lower}' == "pin"               
                Click Action   path:${pin_input}
                Type Text Action    ${Pin}
                 Capture Screenshot   

            END
            
            ${mobno_titlevalue}    Get Text Action    path:${mobile_title}
            ${mobno_titlevalue_lower}    Convert To Lower Case    ${mobno_titlevalue}
            IF    '${mobno_titlevalue_lower}' == "mobile no"  
                # Capture Screenshot             
                Click Action    path:${phn&mob_input} 
                Type Text Action    ${Phone & Mobile No.}  
                Capture Screenshot   

            END   
            
            Capture Screenshot
            Click Action    path:${create_new_cust}

            Sleep    ${avg_time}                          
            ${cust_creation_msg_content}    Get Attribute Action    path:${chassis_customer_creation_msg_box}    Name
            Log    ${cust_creation_msg_content}

            IF    """${cust_creation_msg_content}""" == """Customer created."""              
                
                # Click Action    path:${cust_created_ok}  
                Press Keys Action    enter  
                Capture Screenshot 
                RETURN    ${True}    ${cust_creation_msg_content}          
            
            ELSE   #error or exception

                ${cust_creation_msg_content}    Get Attribute Action    path:${chassis_customer_creation_msg_box}    Name
                Log    ${cust_creation_msg_content}
                Click Action    path:${cust_created_ok}
                Click Action    path:${create_cust_close_btn}
                Click Action    path:${create_chassis_checkbox_win_close}
                Click Action    path:${create_chassis_1st_window_close} 
                RETURN    ${False}    ${cust_creation_msg_content}

            END                                               
            # Click Action    path:${create_cust_close_btn}            
            
        END  

    EXCEPT  AS  ${error_message}
        Log  ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred during the New Customer Data Entry process.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    path:${create_cust_close_btn}
        Click Action    path:${create_chassis_checkbox_win_close}
        Click Action    path:${create_chassis_1st_window_close} 
        # Exit For Loop
    END      


Update Registraion
    [Arguments]    ${Customer Name}    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}
    TRY
            #update Registration details    
            Capture Screenshot        
            Click Action    path:${update_regis_checkbox}        
            Click Action    path:${regis_no_input}        
            Type Text Action    ${Registration No.}

            # Click Action    path:${create_chassis_checkbox_win_close}

            Click Action    path:${create_chassis_overall_save}
            Capture Screenshot
            Sleep    ${avg_time}

            ${regist_popup_message_content}    Get Attribute Action    path:1|1|1|3    Name    
            log    ${regist_popup_message_content}

            IF    """${regist_popup_message_content}""" == """Registration And Customer Details Updated"""
                Capture Screenshot
                Click Action    path:1|1|1|1    #ok will leads to tab entry page in db
                RETURN    ${True}
            
            ELSE
                ${regist_popup_message_content}    Get Attribute Action    path:1|1|1|3    Name    
                log    ${regist_popup_message_content}
                Click Action    path:1|1|1|1
                Click Action    path:${create_chassis_checkbox_win_close}
                Click Action    path:${create_chassis_1st_window_close}
                RETURN  ${regist_popup_message_content}

            END
            

    EXCEPT  AS  ${error_message}
        Log  ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while updating the registration.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    path:${create_chassis_checkbox_win_close}
        Click Action    path:${create_chassis_1st_window_close} 
        # Exit For Loop
    END          


 