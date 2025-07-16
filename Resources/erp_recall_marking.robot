*** Settings ***
Library             RPA.Desktop
Library             RPA.FileSystem
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             RPA.Images
Library             OperatingSystem
Library             String
Library             Collections
Library             RPA.Windows
Library             Libraries/ERP_methods.py   
Library             Libraries/business_operations.py       
Variables           Variables/variables.py 
Resource            Wrappers.robot
Library             SikuliLibrary  mode=OLD


*** Variables ***
${recall_marking_for_chassis_title}    3|1|1|3|1|1|1|1
${recall_report_btn}    3|1|1|2|1|1|8
${current_month_3dot}    1|1|1|2|1|1|1|1|1|5
${submit_to recall_report}    1|1|2|2
${circ_no_column}    3|1|1|1|1|1|1|1|1|1|2|6
${designer_btn}    3|1|1|2|1|1|1|1|4
${filter_btn}    1|1|1|1|1|1
${circular_filter}    1|3|5|1
${reg_filter}    1|2|1
${operator_choose}    1|4|5|2 
${equal_value}    1|3|2|1|1
${cir_filter_valueinput}    1|4|5|3
${filter_apply_btn}    3|2
${recall_row}    3|1|1|1|1|1|1|1|1|1|3|1|1
${first_cell_filter}    1|3|1|1
${selected_cell_filter}    1|2|1
${imagerootfolder}            ${CURDIR}${/}..\\Locators
# ${log_folder}                 ${CURDIR}${/}..\\Log
# ${log_folder}     C:\\Users\popular\Desktop\JobcardOpeningIntegrated\Jobcard-Opening\Screenshot
${current_month_button}          ${imagerootfolder}\\current_month_button.png
${transaction_button}        ${imagerootfolder}\\menu_transaction_btn.png
${rate_header}          ${imagerootfolder}\\Rate_header.png
${voucher_recallvalue_cell}        3|1|1|1|1|1|1|1|1|1|4|1|3
${edit_transaction}    1|1
${chassis_header}    3|1|1|1|2|1|1|1|1|1|1|1|1|1|2|2
${chassis_row_prefix}    3|1|1|1|2|1|1|1|1|1|1|1|1|1                
${first_comment_row_value}    3|1|1|1|2|1|1|1|1|1|1|1|1|1|2|8
${rate_first_cell}    3|1|1|1|2|1|1|1|1|1|1|1|1|1|2|6
${final_save}    3|1|1|3|1|2|1|1|2
${rate_path_prefix}    3|1|1|1|2|1|1|1|1|1|1|1|1|1
${k}    3
${recall_close tab}    1|1
${master_name_title}    ${imagerootfolder}\\master_name_title.png
${Recall_Marking_Chassis_Title}    ${imagerootfolder}\\Recall_Marking_Chassis_Title.png    
${Designer_btn}    ${imagerootfolder}\\Designer_btn.png
${recall_report_img}    ${imagerootfolder}\\recall_report_img.png
${xlsx_extn_img}    ${imagerootfolder}\\xlsx_img.png
# ${Recall_reports_static_path}    E:\JobcardOpeningIntegrated\Recall Reports\Report
${filter_editor_img}    ${imagerootfolder}\\filter_editor_img.png  
${refresh_btn_img}    ${imagerootfolder}\\refresh_btn.png
${excel_download_confirm}    1|1|1|1|2|4
${filter_choose_value}    1|1|1|2|1|1|3  
${input_box_filter}    1|1|1|1  
${filter_ok}    1|1|2
${filter_apply}    1|2|1|2  
${filter_confirm}    1|2|1|3
${Voucher_no_header}    3|1|1|1|1|1|1|1|1|1|3|3
${Report_first_sl_no_path_prod}    3|1|1|1|1|1|1|1|1|1|4|1|1   
${Report_first_sl_no_path_test}    3|1|1|1|1|1|1|1|1|1|3|1|1 
${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx
*** Keywords ***

ERP Recall Marking Navigation
    
    TRY
        Window Navigation    Wings ERP 23E - Web Client
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
            ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
            IF    ${trans_button_exists}==True            
                SikuliLibrary.Click    ${transaction_button}            
            END
        Click Action   name:${menu_transactions} > ${menu_AutoDms}
        Click Action   name:${menu_transactions} > ${menu_service}    
        Click Action   name:${menu_transactions} > ${menu_value_added_services}
        Click Action   name:${menu_transactions} > ${menu_recall_marking}
        Sleep    ${Min_time}

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${recall_report_img}    ${sik_max_time}
        ${recall_report_img_exists}=    SikuliLibrary.Exists    ${recall_report_img}
        IF    ${recall_report_img_exists}==${True}            
            RETURN  ${True}       
        END
            
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        # Exit For Loop

    END  

Production ERP Recall Marking Navigation
    
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
        Click Action   name:${menu_transactions} > ${menu_recall_marking}
        Sleep    ${Min_time}

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${recall_report_img}    ${sik_max_time}
        ${recall_report_img_exists}=    SikuliLibrary.Exists    ${recall_report_img}
        IF    ${recall_report_img_exists}==${True}            
            RETURN  ${True}       
        END
            
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        # Exit For Loop

    END  

Recall Report Button
    [Arguments]    ${Job Card No.}    ${Input_Sheet_Path}    

    TRY
        Window Navigation    Wings ERP 23E - Web Client
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${recall_report_img}    ${sik_time}
        ${recall_report_img_exists}=    SikuliLibrary.Exists    ${recall_report_img}
        IF    ${recall_report_img_exists}==${True}            
            SikuliLibrary.Click    ${recall_report_img}      
        END
        Sleep    ${Min_time}
        Click Action    name:${reportoptn_choose_btn}
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${recall_report_img}    ${sik_max_time}
        ${recall_report_img_exists}=    SikuliLibrary.Exists    ${recall_report_img}
        IF    ${recall_report_img_exists}==${True}            
            RETURN  ${True}       
        END  


    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while clicking the Recall Report Button.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Press Keys Action    esc 
        # Exit For Loop

    END    

Report Period Window
    [Arguments]    ${Job Card No.}    ${Input_Sheet_Path}
    TRY

        Sleep    ${avg_time}
        Click Action    path:${current_month_3dot}
        
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while accessing the Report Period Window.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Press Keys Action    esc 
        # Exit For Loop   
    END   

Choosing Current Month
    [Arguments]    ${Job Card No.}    ${Input_Sheet_Path}
    TRY     
        Window Navigation    Wings ERP 23E - Web Client
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${current_month_button}     ${sik_time}
        ${current_month_button_exists}=    SikuliLibrary.Exists    ${current_month_button}
        IF    ${current_month_button_exists}==${True}
            SikuliLibrary.Click    ${current_month_button}
        END

        Click Action    path:${submit_to recall_report}
        Press Keys Action    enter

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${refresh_btn_img}     ${sik_max_time}
        ${refresh_btn_img_exists}=    SikuliLibrary.Exists    ${refresh_btn_img}
        IF    ${refresh_btn_img_exists}==${True}
            RETURN  ${True}
        END

    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while choosing the current month.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Press Keys Action    esc
        Press Keys Action    esc 
        # Exit For Loop
        
    END    


Recall Report Window
    [Arguments]    ${Branch}     ${Recall Code}    ${Chassis No.}    ${Registration No.}    ${Job Card No.}     ${Input_Sheet_Path}   
    TRY 

        ${first_sl_no_value}    Get Value Action    path:${Report_first_sl_no_path_test}
        # ${first_sl_no_value}    Get Value Action    path:${Report_first_sl_no_path_prod}
        IF    ${first_sl_no_value} == ${Report_first_sl_no}  

            ${recall_report_path_name}    ${download_status}     Report Download Status
            IF    '${download_status}' == '${True}'

                    ${tag_result}    ERP_methods.Recall Tagged Or Not    ${recall_report_path_name}    ${Branch}    ${Recall Code}    ${Chassis No.}    ${Registration No.}
                    Log    ${tag_result}
                    RETURN    ${tag_result}   

            ELSE

                    ${recall_report_path_name_2}    ${download_status_2}     Report Download Status
                    IF    """${download_status_2}""" == """${False}"""                           
                            RETURN    ${download_fail}   
                    END

            END
           

        ELSE 
           
            RPA.Desktop.Press Keys  Alt  C  
            RETURN    ${False}            

        END
            
    
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while exporting the Recall Report.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        RPA.Desktop.Press Keys  Alt  C
        # Exit For Loop
    END        

Report Download Status  

    RPA.Desktop.Press Keys    alt  x
    Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${xlsx_extn_img}     ${sik_time}
    ${xlsx_extn_img_exists}=    SikuliLibrary.Exists    ${xlsx_extn_img}
    IF    ${xlsx_extn_img_exists}==${True}
        SikuliLibrary.Click    ${xlsx_extn_img}
    END

    ${recall_report_path_name}     Save Recall Report  
    Log    ${recall_report_path_name}
    Sleep    ${Min_sleep}
    ${download_status}    Run Keyword And Return Status    Click Action    path:${excel_download_confirm}

    RETURN    ${recall_report_path_name}    ${download_status}    


Recall Editing
    [Arguments]    ${Job Card No.}    ${tag_value}    ${Input_Sheet_Path}   
    TRY
        Window Navigation    Wings ERP 23E - Web Client

        Right Click Action    path:${Voucher_no_header}    
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${filter_editor_img}     ${sik_time}
        ${filter_editor_img_exists}=    SikuliLibrary.Exists    ${filter_editor_img}
        IF    ${filter_editor_img_exists}==${True}
            SikuliLibrary.Click    ${filter_editor_img}
        END

        Click Action    path:${filter_choose_value} 
        Click Action    path:${input_box_filter}         
        RPA.Desktop.Press Keys     ctrl  a           
        Type Text Action    ${tag_value}
        Click Action    path:${filter_ok}    
        Click Action    path:${filter_apply}     
        Click Action    path:${filter_confirm}      

        Right Click Action    path:${voucher_recallvalue_cell}
        Click Action    path:${edit_transaction}  


    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while editing the recall.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        RPA.Desktop.Press Keys  Alt  C
        # Exit For Loop
    END      

Recall Chassis Adding
    [Arguments]    ${Chassis No.}   ${Job Card No.}    ${Input_Sheet_Path}  
 
    
    TRY
        Wait Action    Window Navigation    Wings ERP 23E - Web Client
        Click Action    path:${chassis_header}
        
        ${chassis_1st_val}=    Get Value Action    path:${chassis_row_prefix}|${k}|2
        
        WHILE    '${chassis_1st_val}' != '(null)'

            Log  ${chassis_1st_val}
            IF    '${chassis_1st_val}' == '${Chassis no.}'
                Recall Not Existing    ${Job Card No.}    ${Input_Sheet_Path}    
                BREAK
            END 
            ${k}=    Evaluate    ${k} + 1
            Press Keys Action    down
            ${chassis_1st_val}=    Get Value Action    path:${chassis_row_prefix}|${k}|2
        
        END
        Type Text Action    ${Chassis No.}
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${master_name_title}    ${sik_max_time}
            ${master_title_exists}=    SikuliLibrary.Exists    ${master_name_title}
            IF    ${master_title_exists}==${True}            
                    Press Keys Action    tab      
            END

        Recall Claim Labour Details tab    ${Job Card No.}    ${Input_Sheet_Path} 
    EXCEPT  AS   ${error_message}
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while adding a chassis to the recall.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Press Keys Action    esc
        RPA.Desktop.Press Keys  enter
        Press Keys Action    esc
        RPA.Desktop.Press Keys  enter
        # Exit For Loop
    END      
Recall Not Existing
    [Arguments]    ${Job Card No.}    ${Input_Sheet_Path}
    TRY    
        Window Navigation    Wings ERP 23E - Web Client
        Press Keys Action    esc
        Click Action    path:${recall_close tab}
    
    EXCEPT  AS   ${error_message}       
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred when handling a non-existing recall.    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Press Keys Action    esc
        # Exit For Loop
    END          

Recall Claim Labour Details tab  
    [Arguments]    ${Job Card No.}    ${Input_Sheet_Path}
    TRY
        Window Navigation    Wings ERP 23E - Web Client  
        Press Keys Action    f4
        RPA.Windows.Double Click    path:${rate_first_cell}

        # Initialize the row index
        ${j}    Set Variable    2
        
        # Start a while loop that continues until rate_value becomes (null)
        WHILE    True
            ${rate_value}    Get Value Action    path:${rate_path_prefix}|${j}|6

            # Exit the loop if rate_value is (null)
            IF    '${rate_value}' == '(null)'
                Exit For Loop
            END

            IF    '${rate_value}' != '0.00'
                Press Keys Action    down
            ELSE
                RPA.Desktop.Press Keys    tab
                ${comment_value}=    Get Value Action    path:${rate_path_prefix}|${j}|8
                RPA.Desktop.Press Keys    shift  tab
                IF    '${comment_value}' != '(null)'
                    Type Text Action   ${comment_value}
                END
                RPA.Desktop.Press Keys    down
            END

            # Increment row index for next iteration
            ${j}    Evaluate    ${j} + 1
        END


        #Transaction Saving
        Double Click Action    path:${final_save} 
        Sleep    ${avg_time}

        ${save_confirm_message}    Get Text Action    path:1  
        
        #closing window popup
        IF    "${save_confirm_message}" == "Close"
            
            Click Action    path:1|2
            Sleep    ${Min_time}
            ${save_confirm_message}    Get Text Action    path:1 
            
        END
        Sleep    ${Min_time}
        
        #Transacion Saving Options
        IF   """${save_confirm_message}""" == """Transaction saved."""  
         
            
            ${confirm_message_content}    Get Attribute    path:1|3    Name    
            log    ${confirm_message_content}
            Click Action    path:1|1 
            RETURN    ${confirm_message_content}

            # RETURN  ${True}

        ELSE IF  """${save_confirm_message}""" == """Error"""
            
            ${confirm_message_content}    Get Attribute    path:1|3    Name    
            log    ${confirm_message_content}
            Click Action    path:1|1
            # RETURN  ${Error}
            RETURN    ${confirm_message_content}

        ELSE IF  """${save_confirm_message}""" == """Exception"""             
              
            ${confirm_message_content}    Get Attribute    path:1|3    Name    
            log    ${confirm_message_content}
            Click Action    path:1|1|1|1|2|4
            RETURN    ${confirm_message_content}
            # RETURN  ${Exception_msg} 
        END
    
        
        
    EXCEPT  AS   ${error_message}    
        
        Log    ${error_message}
        Capture Screenshot
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred in Recall Claim Labour Details tab or Transaction Saving    ${exception_reason_column_name}
        update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        Press Keys Action    esc
        RPA.Desktop.Press Keys  enter  
        Press Keys Action    esc
        RPA.Desktop.Press Keys  enter
        # Exit For Loop
        
    END  
   
   

Save Recall Report 
    ${project_root}     ERP_methods.get_process_root_directory
    Log    ${project_root}
    ${recall_report_path}    Set Variable    ${project_root}\\${recall_report_folder_name}\\${curr_date}
    ${file_name}    prepare_file_name
    ${recall_report_path_name}    Set Variable    ${recall_report_path}\\${file_name}.xlsx
    Log    ${recall_report_path_name}
    Sleep    ${Min_sleep}
    Type Text    ${recall_report_path_name}     enter=True  
    RETURN    ${recall_report_path_name}
                 






