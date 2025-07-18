*** Settings ***
Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             SikuliLibrary    mode=OLD
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

*** Variables ***

# ${ERP_Password_Path}       ${CURDIR}${/}..\\Config\\Popular_Credentials.xlsx
${ERP_Password_Path}    C:\\JobcardOpeningIntegrated\\Config\\Popular_Credentials.xlsx
${keep_button}    1|3|1|1|2|1|1|1|4|1|3
${dms_link}    1|3|1|1|2|1|1|1|4|1 
${keep_btn_img}                        ${imagerootfolder}\\keep_btn_img.png
${open_file_img}                       ${imagerootfolder}\\open_file_img.png
${time_out}    80  
${recall_bill_timeout}    120   
${session_time_out}   5
${dont_show1_popup}                    ${imagerootfolder}\\dont_show1_popup.png    
${popup1_run_btn}                      ${imagerootfolder}\\popup1_run_btn.png
${Dms_run_popup_2}                     ${imagerootfolder}\\Dms_run_popup_2.png
${popup2_run_btn}                      ${imagerootfolder}\\popup2_run_btn.png
${login_title}    //*[@id="title"]
${DMS_blue_icon}    //*[@class="dms1IconDiv"]            
${username_field}    //*[@id='username']
${password_field}    //*[@id='password']
${submit_field}    //*[@type="submit"]
${login_failed}    //*[@class="loginFailed"]
*** Keywords ***
DMS_V3_Login

    TRY
        Log    ${ERP_Password_Path}
        Log    ${CURDIR}
        Log    ${CURDIR}${/}..
        ${file_status}    ERP_methods.Check File Exists    ${ERP_Password_Path}
        IF    '${file_status}' == '${True}'
            Open Workbook    ${ERP_Password_Path}
            ${credential_data}=    Read Worksheet As Table    header=${True}
            Sleep    ${Min_time}
            FOR    ${row}    IN    @{credential_data}
                ${v3uname}=    Set Variable    ${row}[DMS V3 UserName]
                ${v3pword}=    Set Variable    ${row}[DMS V3 Password]
                ${v3url}=    Set Variable    ${row}[DMS V3 URL]
            END

            IF    '${v3uname}' != '${None}' and '${v3pword}' != '${None}' and '${v3url}' != '${None}'
                ${login_success}=    Set Variable    ${False}
                FOR    ${i}    IN RANGE    3
                    Log    Attempt ${i+1} to login
                    ${edge_open}=    Run Keyword And Return Status    Open Browser    browser=edge    url=${v3url}
                    IF    ${edge_open} == ${True}
                        Maximize Browser Window
                        ${dms_icon_visible}=    Credentials Feed    ${v3uname}    ${v3pword}
                        IF    ${dms_icon_visible} == ${True}
                            ${login_success}=    Set Variable    ${True}
                            Exit For Loop
                        ELSE
                            Close Browser Processes
                        END
                    END
                END

                IF    ${login_success} == ${True}
                    ${Oracle_form_download_status}    Download Login Link
                    IF    ${Oracle_form_download_status} == ${False}
                        RETURN  ${False}
                    ELSE   
                        RETURN  ${True}
                    END
                ELSE
                    # Close Browser
                    RETURN  ${False}
                    Close Browser Processes
                    # ERP_methods.Show Message Box    Alert    Login failed after 3 attempts. Update credentials or check access.
                    Fail
                END
            ELSE
                RETURN  ${False}
                ERP_methods.Show Message Box    Alert    DMS Login Error. Update and rerun the bot after closing windows
                Fail
            END
        ELSE
            RETURN  ${False}
            ERP_methods.Show Message Box    Alert    Credential Excel doesn't exist. Update and rerun the bot after closing windows
            Fail
        END
    EXCEPT    AS    ${error_message}
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message} 
    END


Credentials Feed
    [Arguments]    ${v3uname}    ${v3pword}            
    TRY 

        ${login_title_status}    Run Keyword And Return Status    Wait Until Element Is Visible    ${login_title}    timeout=${DMS_window_time}
        # Wait Until Element Is Visible    ${login_title}    timeout=${DMS_window_time}
        IF    ${login_title_status} == ${False}

            RETURN    ${False}
        
        ELSE
                ${uname_value1}    Selenium Get Text    ${login_title} 

            Log    ${uname_value1}
            IF    '${uname_value1}' == 'Login'
                
                ${authentication_failed}    Credential Data    ${v3uname}    ${v3pword}
                
                IF    ${authentication_failed} == ${True}
                    
                    ${authentication_failed}    Credential Data    ${v3uname}    ${v3pword}

                    IF    ${authentication_failed} == ${True}

                        ${authentication_failed}    Credential Data    ${v3uname}    ${v3pword}
                    
                        IF    ${authentication_failed} == ${True}

                            Close Browser Processes
                            RETURN    ${False}

                        END
                    
                    ELSE
                        
                        Run Keyword And Ignore Error    Wait Until Element Is Visible    ${DMS_blue_icon}    timeout=${DMS_window_time}
                        ${dms_icon_visible}    Is Element Visible    ${DMS_blue_icon}
                        RETURN    ${dms_icon_visible}
                    
                    END

                ELSE
                    Run Keyword And Ignore Error    Wait Until Element Is Visible    ${DMS_blue_icon}    timeout=${DMS_window_time}
                    ${dms_icon_visible}    Is Element Visible    ${DMS_blue_icon}
                    RETURN    ${dms_icon_visible}
                END


            END
        END
                 
    EXCEPT  AS   ${error_message}
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message}     
    END

Credential Data
    [Arguments]    ${v3uname}    ${v3pword}
    TRY
        Selenium Click Element    ${username_field}
        Selenium Input Text    ${username_field}   ${v3uname}

        Selenium Click Element    ${password_field}
        Selenium Input Text    ${password_field}    ${v3pword}
    
        Selenium Click Button    ${submit_field}

        Sleep    1

        ${authentication_failed}    Is Element Visible    ${login_failed}
        RETURN    ${authentication_failed}

    EXCEPT  AS   ${error_message}
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message}    
    END

Download Login Link  
    TRY     


        Selenium Click Element    ${DMS_blue_icon}

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${keep_btn_img}     ${time_out}
        ${keep_btn_img_exists}=    SikuliLibrary.Exists    ${keep_btn_img}   
        IF    ${keep_btn_img_exists}==True   

            SikuliLibrary.Click    ${keep_btn_img} 

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${open_file_img}     ${time_out}
            ${open_file_img_exists}=    SikuliLibrary.Exists    ${open_file_img}   
            IF    ${open_file_img_exists}==True  

                SikuliLibrary.Click    ${open_file_img} 

            END

            #Additional Popups
            ${popup_handle_status}    Run Keyword And Return Status    DMS_Run_Popups
            

        ELSE

            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${open_file_img}     ${time_out}
            ${open_file_img_exists}=    SikuliLibrary.Exists    ${open_file_img}   
            IF    ${open_file_img_exists}==True  

                SikuliLibrary.Click    ${open_file_img} 

            END
        END
        RETURN    ${True}
    

    EXCEPT  AS   ${error_message}
         log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message}  
    END

DMS_Run_Popups
    TRY
        
        #DMS run first popup
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${dont_show1_popup}     ${time_out}
        ${dont_show1_popup_exists}=    SikuliLibrary.Exists    ${dont_show1_popup}   
        IF    ${dont_show1_popup_exists}==True  

            ${dont_show1_popup_status}    Run Keyword And Return Status    SikuliLibrary.Click    ${dont_show1_popup} 
            IF    ${dont_show1_popup_status} == ${True}
                SikuliLibrary.Click    ${popup1_run_btn}
            END


        END

        #DMS run second popup
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${Dms_run_popup_2}     ${time_out}
        ${Dms_run_popup_2_exists}=    SikuliLibrary.Exists    ${Dms_run_popup_2}   
        IF    ${Dms_run_popup_2_exists}==True  

            ${Dms_run_popup_2_status}    Run Keyword And Return Status    SikuliLibrary.Click    ${Dms_run_popup_2} 
            IF    ${Dms_run_popup_2_status} == ${True}
                SikuliLibrary.Click    ${popup2_run_btn}
            END


        END
        


    EXCEPT  AS   ${error_message}
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message}     
    END


