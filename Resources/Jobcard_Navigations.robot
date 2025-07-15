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
Library             SikuliLibrary  mode=OLD

*** Variables ***
${testdb_Submit}    1|1|1|1|1|2|3|2
${testdb_hyperlink}    2|1|1|2|2|1|1|2|1|1
${Branch_path}    3|1|1|1|1|1|1|1|3|2|1|1
${Location_path}   3|1|1|1|1|1|1|1|4|2|1|1
${close_erp_window_approval}   1|1
${imagerootfolder}            ${CURDIR}${/}..\\Locators
# ${log_folder}                 ${CURDIR}${/}..\\Log
# ${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
${log_folder}     C:\\JobcardOpeningIntegrated\\Screenshot
${current_month_button}          ${imagerootfolder}\\current_month_button.png
${jobcard_button}          ${imagerootfolder}\\jobcard_btn.png
${service_jobcard_button}          ${imagerootfolder}\\service_jc_btn.png
${transaction_button}        ${imagerootfolder}\\menu_transaction_btn.png
${testdb_choose_button}    ${imagerootfolder}\\testdb_choose_btn.png
${Service_Jobcard_Heading}    ${imagerootfolder}\\Service_Jobcard_Heading.png
${Bodyshop Jobcard Title}    ${imagerootfolder}\\Bodyshop Jobcard Title.png
${bodyshop_jobcard_menu_img}    ${imagerootfolder}\\bodyshop_jobcard_menu_img.png
${bodyshop_jobcard_btn}    ${imagerootfolder}\\bodyshop_jobcard_btn.png
${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx

*** Keywords ***

Service Jobcard Navigation

    TRY

        # Window Navigation    Wings ERP 23E - Web Client

        # Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
        # ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
        # IF    ${trans_button_exists}==${True}            
        #     SikuliLibrary.Click    ${transaction_button}            
        # END
        # Click Action   name:${menu_transactions} > ${menu_AutoDms}
        # Click Action   name:${menu_transactions} > ${menu_service}    
        # Click Action   name:${menu_transactions} > ${menu_regular_service} 
        # Capture Screenshot

        # Press Keys Action    tab
        # Press Keys Action    enter  
        # Press Keys Action    enter

        # Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${Service_Jobcard_Heading}     ${sik_max_time}
        # ${service_jobcard_heading_exists}=    SikuliLibrary.Exists    ${Service_Jobcard_Heading}
        # IF    ${service_jobcard_heading_exists}==${True}
        #     RETURN    ${True}
        # END
        ${retry_count}=    Set Variable    0
        FOR    ${i}    IN RANGE    3
           Log    Attempt ${i+1} to navigate AutoDMS
           Run Keyword And Ignore Error    Window Navigation    Wings ERP 23E - Web Client
           Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
           ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
           IF    ${trans_button_exists}==${True}            
               SikuliLibrary.Click    ${transaction_button}            
           END
           ${click1}=    Run Keyword And Return Status    Click Action   name:${menu_transactions} > ${menu_AutoDms}
           ${click2}=    Run Keyword And Return Status    Click Action   name:${menu_transactions} > ${menu_service}    
           ${click3}=    Run Keyword And Return Status    Click Action   name:${menu_transactions} > ${menu_regular_service}   
           IF    '${click1}'=='False' or '${click2}'=='False' or '${click3}'=='False'
        
                Click Action    path:2|1    #file btn
                Continue For Loop   
           END 
           Capture Screenshot
           Press Keys Action    tab
           Press Keys Action    enter  
           Press Keys Action    enter
           Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${Service_Jobcard_Heading}     ${sik_max_time}
           ${service_jobcard_heading_exists}=    SikuliLibrary.Exists    ${Service_Jobcard_Heading}
           IF    ${service_jobcard_heading_exists}==${True}
               RETURN    ${True}
           END
          
        END
        
        
    EXCEPT  AS   ${error_message}
        Log    ${error_message} 
        Capture Screenshot
        Click Action   path:2|1    #file btn
        
    END    


BodyShop Jobcard Navigation
    TRY

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
        ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
        IF    ${trans_button_exists}==True            
            SikuliLibrary.Click    ${transaction_button}            
        END

        Click Action   name:${menu_transactions} > ${menu_AutoDms}
        Click Action   name:${menu_transactions} > ${menu_service}
        Click Action   name:${menu_transactions} > ${menu_bodyshop_service}
        # Click Action   name:${menu_transactions} > ${menu_job_card}
        # Click Action   name:${menu_transactions} > ${menu_bodyshop_jobcard}

        #  Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${bodyshop_jobcard_btn}    ${sik_time}
        # ${body_jobcard_button_exists}=    SikuliLibrary.Exists    ${bodyshop_jobcard_btn}
        # IF    ${body_jobcard_button_exists}==${True}
        #     SikuliLibrary.Click    ${bodyshop_jobcard_btn}
        # END

        # Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${bodyshop_jobcard_menu_img}     ${sik_time}
        # ${bodyshop_jobcard_menu_img_exists}=    SikuliLibrary.Exists    ${bodyshop_jobcard_menu_img}
        # IF    ${bodyshop_jobcard_menu_img_exists}==${True}
        #     SikuliLibrary.Click    ${bodyshop_jobcard_menu_img}
        # END
        Press Keys Action    tab
        Press Keys Action    enter  
        Press Keys Action    enter

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${Bodyshop Jobcard Title}     ${sik_max_time}
        ${bodyshop_jobcard_heading_exists}=    SikuliLibrary.Exists    ${Bodyshop Jobcard Title}
        IF    ${bodyshop_jobcard_heading_exists}==${True}
            RETURN    ${True}
        END 

    EXCEPT  AS   ${error_message}
        Log  ${error_message} 
        Capture Screenshot
        Click Action   name:${menu_transactions}  
            

    END    



Production Service Jobcard Navigation

    TRY

        Window Navigation    Wings ERP 23E - Web Client

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
        ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
        IF    ${trans_button_exists}==${True}            
            SikuliLibrary.Click    ${transaction_button}            
        END
        Click Action   name:${menu_transactions} > ${menu_AutoDms}
        Click Action   name:${menu_transactions} > ${menu_service} 
        Click Action   name:${menu_transactions} > ${menu_services}   
        Click Action   name:${menu_transactions} > ${menu_regular_service_jobcard} 

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${Service_Jobcard_Heading}     ${sik_max_time}
        ${service_jobcard_heading_exists}=    SikuliLibrary.Exists    ${Service_Jobcard_Heading}
        IF    ${service_jobcard_heading_exists}==${True}
            RETURN    ${True}
        END
                
        
    EXCEPT  AS   ${error_message}
        Log    ${error_message} 
        Capture Screenshot
        Click Action   name:${menu_transactions}      
        
    END 

Production Bodyshop Jobcard Navigation

    TRY

        Window Navigation    Wings ERP 23E - Web Client

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${transaction_button}    ${sik_max_time}
        ${trans_button_exists}=    SikuliLibrary.Exists    ${transaction_button}
        IF    ${trans_button_exists}==${True}            
            SikuliLibrary.Click    ${transaction_button}            
        END
        Click Action   name:${menu_transactions} > ${menu_AutoDms}
        Click Action   name:${menu_transactions} > ${menu_service} 
        Click Action   name:${menu_transactions} > ${menu_bodyshop_service}   
        Click Action   name:${menu_transactions} > ${menu_job_card}
        Click Action   name:${menu_transactions} > ${menu_bodyshopjobcard}
         

        Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${Service_Jobcard_Heading}     ${sik_max_time}
        ${service_jobcard_heading_exists}=    SikuliLibrary.Exists    ${Service_Jobcard_Heading}
        IF    ${service_jobcard_heading_exists}==${True}
            RETURN    ${True}
        END
                
        
    EXCEPT  AS   ${error_message}
        Log    ${error_message} 
        Capture Screenshot
        Click Action   name:${menu_transactions}      
        
    END 