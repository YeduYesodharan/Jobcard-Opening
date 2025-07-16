*** Settings ***
Library   SikuliLibrary  mode=OLD
Library   RPA.Desktop
Library   RPA.Tables
Library   RPA.Excel.Files
library   RPA.Windows
Library   Dialogs
Library    Collections
Library    String
Variables  Variables/variables.py
Library    Libraries/business_operations.py
Resource   Resources/dms_login.robot
Resource   Resources/dms_jobcard_extraction.robot
Library    RPA.Browser.Selenium

*** Variables ***
${print_btn_image}                                ${imagerootfolder}\\print_btn.png
${print_preview_popup_image}                      ${imagerootfolder}\\print_preview_popup.png
${view_btn_image}                                 ${imagerootfolder}\\view_btn.png
${veh_user_name_recall_image}                     ${imagerootfolder}\\veh_user_name_recall.png
${recall_pdf_close_image}                         ${imagerootfolder}\\recall_pdf_close.png
${recall_pdf_jobcard_title_image}                 ${imagerootfolder}\\recall_pdf_jobcard_title.png
# ${log_folder}                                     ${CURDIR}${/}..\\Log
${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
# ${log_folder}     C:\\Users\popular\Desktop\JobcardOpeningIntegrated\Jobcard-Opening\Screenshot


*** Keywords ***
Check Recall Exist
    [Arguments]    ${jc_no}
    TRY
        
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${print_btn_image}    ${time_out}
        ${print_exist}    SikuliLibrary.Exists    ${print_btn_image}
        IF    ${print_exist} == ${False}
            Return From Keyword    ${False}
        ELSE
            SikuliLibrary.Click    ${print_btn_image}     
        END

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${print_preview_popup_image}    ${time_out}
        ${print_preview_popup_image_exist}    SikuliLibrary.Exists    ${print_preview_popup_image}
        
        IF    ${print_preview_popup_image_exist} == ${True}
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${view_btn_image}    ${time_out}
            ${view_btn_image_exist}    SikuliLibrary.Exists    ${view_btn_image}
            
            Capture Screenshot

            IF    ${view_btn_image_exist} == ${True}
                SikuliLibrary.Click    ${view_btn_image}  

                Capture Screenshot 

                ${view_btn_image_exist}    SikuliLibrary.Exists    ${view_btn_image}
                IF    ${view_btn_image_exist} == ${True}  
                       SikuliLibrary.Click    ${view_btn_image}
                END
            ELSE
                Return From Keyword    ${False}
            END
        ELSE
            Return From Keyword    ${False}     
        END

        Run Keyword And Ignore Error        RPA.Browser.Selenium.Maximize Browser Window
        Run Keyword And Ignore Error     RPA.Windows.Maximize Window
        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${recall_pdf_jobcard_title_image}    ${recall_bill_timeout}
        Run Keyword And Ignore Error        RPA.Browser.Selenium.Maximize Browser Window  
        Run Keyword And Ignore Error     RPA.Windows.Maximize Window     
        ${recall_pdf_jobcard_title_image_exist}    SikuliLibrary.Exists    ${recall_pdf_jobcard_title_image}
        Capture Screenshot
        
        ${recall_status}    Set Variable    No
        IF    ${recall_pdf_jobcard_title_image_exist} == ${True}
            
            ${running_status}    ${extracted_value}    Custom Get Text From Image    ${veh_user_name_recall_image}    "0"    "80"    "80"    "5"
           
            IF    ${running_status} == ${False}
                # Close Recall Pdf 
                Close Browser Processes
                Return From Keyword    ${False}    ${extracted_value}    ${recall_status}
            ELSE
                IF    '${extracted_value}' != '${EMPTY}'
                    ${recall_status}    Set Variable    Yes
                END
            END 
        ELSE
            ${extracted_value}    Set Variable    ${EMPTY}
            # Close Recall Pdf
            Close Browser Processes
            Return From Keyword    ${False}    ${extracted_value}    ${recall_status}
        END
    EXCEPT    AS    ${error_message}
        Log    ${error_message}
        # Close Recall Pdf
        Close Browser Processes
        Capture Screenshot
        Fail   ${error_message} 
    END
    # Close Recall Pdf  
    Close Browser Processes
    Return From Keyword    ${True}    ${extracted_value}    ${recall_status}













