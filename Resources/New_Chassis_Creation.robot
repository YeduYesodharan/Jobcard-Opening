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
${create_cust_submit}    1|1|1|2|2|1 
${cust_submit_closebtn}    1|1|1|2|2|3
${search_cust_btn}    1|1|1|2|1|1
${search_custname_input}    1|1|1|1|1|1|1|1
${search_custmob_input}    1|1|1|1|1|1|1|2
${search_custwith_details}    1|1|1|1|1|1|1|5|1
${search_cust_go}    1|1|1|1|1|1|1|6
${search_cust_window_ok}    1|1|1|1|1|3|2
${search_cust_window_close}    1|1|1|1|1|3|1
${update_regis_checkbox}    1|1|1|3|1|1
${regis_no_input}    1|1|1|3|1|7|1
${overall_chassis_crate_close_btn}    1|1|1|5|2
${chsis_no_title}    1|1|1|1|1|3
${vehicle_sermc_title}    1|1|1|1|1|5
${vehicle_title}    1|1|1|1|1|7
${engine_no_title}    1|1|1|1|1|11
${color_title}    1|1|1|1|1|13
${F_Date_ofsale_title}    1|1|1|1|1|17
${create_chassis_datapage_close}    1|1|1|2|2|2
${create_chassis_1st_window_close}    1|3|1                       
${create_chassis_checkbox_win_close}    1|1|1|5|2
${chassis_final_ok}    1|1|1|1|1
${chassis_customer_creation_msg_box}    1|1|1|1|3
${service_model_code_sheet}       ${CURDIR}${/}..\\Mapping\\Service Model Codes.xlsx
# ${log_folder}     ${CURDIR}${/}..\\Log
# ${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
${log_folder}     C:\\JobcardOpeningIntegrated\\Screenshot
${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx
${chassis_checkbox_window_chassispath}    1|1|1|1|1|2

*** Keywords ***


New Chassis Creation Initiaion
    [Arguments]   ${Chassis No.}  ${Vehicle Service Model code}  ${Vehicle}  ${Engine Number}  ${Colour}  ${Sale Date}    ${Customer Name}    
    ...    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Job Card No.}    ${Input_Sheet_Path}
    
    TRY
      
        ${creation_chassis_click}    Run Keyword And Return Status    Double Click Action    path:${chassis_creation_btn}  
        IF    ${creation_chassis_click} == ${True}
            RETURN    ${True}
        ELSE
            RETURN    ${False}
        END
        Capture Screenshot
            
    EXCEPT  AS  ${Chassis_creation_btn_error}
        Log  ${Chassis_creation_btn_error}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred during the initiation of New Chassis Creation.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    path:${create_chassis_1st_window_close}
        # Exit For Loop
    END   

New Chassis Creation Checkbox Window
    [Arguments]   ${Chassis No.}  ${Vehicle Service Model code}  ${Vehicle}  ${Engine Number}  ${Colour}  ${Sale Date}    ${Customer Name}    
    ...    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Job Card No.}    ${Input_Sheet_Path}

    TRY
        Capture Screenshot

        Click Action    path:${create_chssis_chckbox}

    EXCEPT  AS  ${Chassis_checkbox_windowerror}
        Log  ${Chassis_checkbox_windowerror}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while interacting with the New Chassis Creation Checkbox Window.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    path:${create_chassis_checkbox_win_close}
        Click Action    path:${create_chassis_1st_window_close}
        # Exit For Loop
    END    

New Chassis Creation Data Window
    [Arguments]   ${Chassis No.}  ${Vehicle Service Model code}  ${Vehicle}  ${Engine Number}  ${Colour}  ${Sale Date}    ${Customer Name}    
    ...    ${State}    ${Address}    ${City}    ${Phone & Mobile No.}    ${Job Card No.}     ${Input_Sheet_Path}   

    TRY

            ${chassis_no_title_value}    Get Text Action    path:${chsis_no_title}
            ${chassis_no_title_value_lower}    Convert To Lower Case    ${chassis_no_title_value}
            IF    '${chassis_no_title_value_lower}' == "chassis number"
                Capture Screenshot
                Click Action    path:${chassis_no_input}
                Type Text Action    ${Chassis No.}  
                Capture Screenshot
            END
            
            ${service_modelcode_code_op}    ERP_methods.Validate Service Model Code    ${service_model_code_sheet}    ${Vehicle Service Model code}
            Log    ${service_modelcode_code_op}
            ${vsmc_title_value}    Get Text Action    path:${vehicle_sermc_title}
            ${vsmc_title_value_lower}    Convert To Lower Case    ${vsmc_title_value}
            IF    '${vsmc_title_value_lower}' == "vehicle service model code"
                
                IF    '${service_modelcode_code_op}' != 'no match found for vehicle service model'
                    # Capture Screenshot
                    Click Action    path:${service_model_code_input}   
                    #to handle an exception popup
                    Run Keyword And Ignore Error    Press Keys Action    enter
                    # Sleep    5      #for manual handling chassis extra popup   
                    Type Text Action    ${service_modelcode_code_op}
                    Press Keys Action    enter
                    Capture Screenshot
                ELSE
                    RETURN   ${None}    ${service_modelcode_code_op} 
                END            
            
            END


            ${vehicle_title_value}    Get Text Action    path:${vehicle_title}
            ${vehicle_title_value_lower}    Convert To Lower Case    ${vehicle_title_value}
            IF    '${vehicle_title_value_lower}' == "vehicle"
                Click Action    path:${vehicle_input}
                Type Text Action    ${Vehicle}
                Press Keys Action    enter
                Capture Screenshot
            END

            ${engine_no_title_value}    Get Text Action    path:${engine_no_title}
            ${engine_no_title_value_lower}    Convert To Lower Case    ${engine_no_title_value}
            IF    '${engine_no_title_value_lower}' == "engine number"
                # Capture Screenshot
                Click Action    path:${engine_number_input}
                Type Text Action    ${Engine Number}
                Press Keys Action    enter
                Capture Screenshot
            END
            
            ${color_title_value}    Get Text Action    path:${color_title}
            ${color_title_value_lower}    Convert To Lower Case    ${color_title_value}
            IF  '${color_title_value_lower}' == "colour"
                Click Action    path:${colour_input}
                Type Text Action    ${Colour}
                # Sleep    10
                Press Keys Action    enter
                Capture Screenshot
                
            END
            ${final_date of sale}    ERP_methods.Convert Date Format    ${Sale Date}
            ${f_dof_sale_value}    Get Text Action    path:${F_Date_ofsale_title}
            ${f_dof_sale_value_lower}    Convert To Lower Case    ${f_dof_sale_value}
            IF  '${f_dof_sale_value_lower}' == "date of sale"
                 Click Action    path:${Date_of_sale_input}
                 Type Text Action    ${final_date of sale}
                 Capture Screenshot
            
            END

            # #----- Save Code Need to disable while testing. start-----            
            Click Action    path:${create_chassis_save}
            Sleep    ${Max_Time}                  
 
            ${Save_chassis_content}    Get Attribute Action    path:${chassis_customer_creation_msg_box}    Name
            IF    """${Save_chassis_content}""" == """Chassis Number With Details Created"""
 
                Click Action    path:${chassis_final_ok}
                Capture Screenshot
                # Press Keys Action    enter
                RETURN    ${True}    ${Save_chassis_content}
 
            ELSE
 
                ${Save_chassis_content}    Get Attribute Action    path:${chassis_customer_creation_msg_box}    Name
               
                # Click Action    path:${chassis_final_ok}
                Press Keys Action    enter
                Capture Screenshot
                Click Action    path:${create_chassis_datapage_close}
                Click Action    path:${create_chassis_checkbox_win_close}
                Click Action    path:${create_chassis_1st_window_close}
                Capture Screenshot
                RETURN    ${False}    ${Save_chassis_content}
 
            END
            # #----- Save Code Need to disable while testing. start-----

            # # ----- Save Code Need to disable while production run. start-----
            # Click Action    path:${create_chassis_datapage_close}
            # # ----- Save Code Need to disable while production run. end-----

    EXCEPT  AS  ${Chassis_creation_data_error}
        Log  ${Chassis_creation_data_error}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred during the New Chassis Creation Data Window processing.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Click Action    path:${create_chassis_datapage_close}
        Click Action    path:${create_chassis_checkbox_win_close}
        Click Action    path:${create_chassis_1st_window_close}
        # Exit For Loop
    END      