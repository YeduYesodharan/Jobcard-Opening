*** Settings ***
Library   SikuliLibrary  mode=OLD
Library   RPA.Desktop
Library   RPA.Tables
Library   RPA.Excel.Files
library   RPA.Windows
Library   Dialogs
Library    Collections
Library    String
Library     RPA.Windows
Variables   Variables/variables.py
Library    Libraries/utility.py
Library    RPA.FileSystem
Library    DateTime
Resource    Resources/Wrappers.robot
Resource    DMS_V3_Mainlogin.robot



*** Variables ***
@{reverse_list}
@{status_updation}
${password}    16510M86J20
${jobcard}
${function_sheetname}   sheet1
${row_count}    0
@{dms_list}
@{jobcard_empty_list}
#${sheet_name}   sheet1
${imagerootfolder}            ${CURDIR}${/}..\\Locators
# ${log_folder}                 ${CURDIR}${/}..\\Log
${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
# ${log_folder}     C:\\Users\popular\Desktop\JobcardOpeningIntegrated\Jobcard-Opening\Screenshot
${login_image}                ${imagerootfolder}\\login.png
${password_image}             ${imagerootfolder}\\password_input.png
${login_ok_button}            ${imagerootfolder}\\login_ok.png
${homeScreen_image}           ${imagerootfolder}\\home_screen_2.png
${screen_blue_image}          ${imagerootfolder}\\blue.png 
${part_image}                 ${imagerootfolder}\\part.png
${screen_image}               ${imagerootfolder}\\screen.png
${dropDown_image}             ${imagerootfolder}\\drop.png
${jobcard_image}              ${imagerootfolder}\\jobcard.png
${jobCardNotFound}               ${imagerootfolder}\\jobcardnotfound.png
${jobCardNotFound_ok_button}     ${imagerootfolder}\\jobcardnotfound_ok.png
${Invalid_Part}           ${imagerootfolder}\\not_found_part.png
${partNumberNotFound_ok_button}    ${imagerootfolder}\\partnumber_ok.png 
${mrp_three_dots}                   ${imagerootfolder}\\dot.png
${bathDetailsScreen}                ${imagerootfolder}\\batch_details_screen.png
${entry_delete_button}              ${imagerootfolder}\\entry_delete_button.png
${find_mrp_inputbox}                ${imagerootfolder}\\find_mrp_inputbox.png
${mrp_find_button}                  ${imagerootfolder}\\mrp_find_button.png 
${mrp_empty_screen_image}                 ${imagerootfolder}\\mrp_empty_screen.png
${mrp_cancel_button_image}                ${imagerootfolder}\\mrp_cancel_button.png 
${entry_delete_button_image}              ${imagerootfolder}\\entry_delete_button.png
${mrp_ok_button_image}                    ${imagerootfolder}\\mrp_ok_button.png
${closeBrowser}                           ${imagerootfolder}\\closeBrowser.png
${dms_exit_image}                          ${imagerootfolder}\\exit_dms.png
${main_screen_image}                        ${imagerootfolder}\\main_screen.png
${partnumber_anotherwindow}                  ${imagerootfolder}\\partnumber_window.png
${issue_type_1}                               ${imagerootfolder}\\issue_type_1.png
${selected_customer}                          ${imagerootfolder}\\issue_type_customer.png
${invalid_quantity}                            ${imagerootfolder}\\stock_issue.png
${invalid_quantity_ok}                         ${imagerootfolder}\\stock_issue_ok.png
${part_desc_image}                                   ${imagerootfolder}\\ew.png
${issue_type_ok}                                   ${imagerootfolder}\\issue_type_ok.png
${total_dms_qunatity}                              ${imagerootfolder}\\total_quantity.png
${multiple_mrp_blue}                                ${imagerootfolder}\\mrp_blue_image.png
${dummy_image}                                      ${imagerootfolder}\\dummy_image.png
${counter}
${sheet_path}                                      Config//Popular_Credentials.xlsx
${edge_run_this_time_button}                       ${imagerootfolder}\\run_this_time.png
${edge_refresh_button}                             ${imagerootfolder}\\edge_refresh_image.png
${sheet_name}                                      Sheet1
${incorrect_password_image}                        ${imagerootfolder}\\incorrect_password.png
${incorrect_password_ok_image}                      ${imagerootfolder}\\incorrect_password_ok.png
${Execution_Status}                                 ${True}
${status_counter}                                   1
${path_location_to_branch}                           Config//Popular_Credentials.xlsx
${another_edge_run_this_button}                      ${imagerootfolder}\\another_run_this_time.png
${restore_button_image}                              ${imagerootfolder}\\restore_button.png
${restore_close_image}                                 ${imagerootfolder}\\restore_close.png 


*** Keywords ***
Read_credentials
    TRY
          #Read DMS credentials from Popular_Credentials
          ${sheet_path}   Set Variable    Config//Popular_Credentials.xlsx
          ${sheet_name}    Set Variable    Sheet1
          
          #To check credential sheet exists or not in config
          ${credential_sheet_exist}    Check Sheet Exists    ${sheet_path}   ${sheet_name}
          IF    ${credential_sheet_exist}==False          
              Fail    Config sheet is missing
          END

          ${status_counter}    Set Variable   1
          Set Global Variable    ${status_counter}
          #To read credentials from credentials sheet in config 
    EXCEPT    AS   ${error_message}
        Capture Screenshot  
        Fail    ${error_message}        
    END
   


dms_login 

  TRY

    ${login_status}    DMS_V3_Login
    IF    ${login_status} == ${False}
        RETURN    ${False}
    END
    #-------------------------commented for DMS V3 integration----------------------------------------#
    ${user_id}   ${password}  ${base_url}   Login Read Credentials From Excel    ${sheet_path}
#     #To dynmaically generate sid
#     ${sid}    Generate Current Date String
#     #this will dynamically generate url

#     ${url}   Generate Dynamic Url    ${user_id}    ${sid}  ${base_url}
#     Log    ${url}
    log    ${password}
#     ${branch_location}      Get Location Value    ${path_location_to_branch}    
    
#     #To open edge with dynamic url
#     Open Edge    ${url}
#     sleep  2
    #-------------------------commented for DMS V3 integration----------------------------------------#
    # ----- Keyword used to switch to the DMS application------
#     Bring Window To Front    ${dms_title}
    
    ${log_status}    Check Restore Button Exist
#     IF    "${log_status}" == "${True}"
#         Open Edge    ${url}
#         sleep  2
#     END
    Capture Screenshot
    #-------------------------commented for DMS V3----------------------------------------#
#     IF    '${branch_location}'=='SRM_KMG'
#            Run Keyword And Ignore Error   SikuliLibrary.Wait Until Screen Contain    ${another_edge_run_this_button}    ${time_out}
#             ${button_exist}=    SikuliLibrary.Exists    ${another_edge_run_this_button}
#     ELSE
#              Run Keyword And Ignore Error   SikuliLibrary.Wait Until Screen Contain    ${edge_run_this_time_button}    ${time_out}
#               ${button_exist}=    SikuliLibrary.Exists    ${edge_run_this_time_button}
    
#     END
    #-------------------------commented for DMS V3----------------------------------------#
    #To check the run this time button exists
    #Run Keyword And Ignore Error   SikuliLibrary.Wait Until Screen Contain    ${edge_run_this_time_button}    ${time_out}
    #${button_exist}=    SikuliLibrary.Exists    ${edge_run_this_time_button}
    #If the button doesnot exist it will retry 5 times to check it by refreshing the page
    ${counter}    Set Variable   1
    WHILE    ${counter}<=4
         
          #-------------------------commented for DMS V3 integration----------------------------------------#
          # IF    ${button_exist}==True
               #-------------------------commented for DMS V3 integration----------------------------------------#
               # IF    '${branch_location}'=='SRM_KMG'

               #      SikuliLibrary.Click   ${another_edge_run_this_button}
               # ELSE
               #      SikuliLibrary.Click     ${edge_run_this_time_button}
               # END   
               #-------------------------commented for DMS V3 integration----------------------------------------#        
               Run Keyword And Ignore Error   SikuliLibrary.Wait Until Screen Contain      ${password_image}  90
               ${password_image_exists}=    SikuliLibrary.Exists  ${password_image} 
               IF    ${password_image_exists}==False
                         
                    Fail
                    
                    
               END 
               Run Keyword And Ignore Error   SikuliLibrary.Click     ${password_image}
               Run Keyword And Ignore Error  SikuliLibrary.Click     ${password_image}
               Sleep    1
               SikuliLibrary.Input Text    ${EMPTY}    ${password}
               Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${login_ok_button}     ${time_out}
               ${login_ok_button_exists}=    SikuliLibrary.Exists    ${login_ok_button}   
               IF    ${login_ok_button_exists}==False
                    
                    Fail

                    
               END
               SikuliLibrary.Click    ${login_ok_button}
               Sleep    2
               ${incorrect_password_exists}=   SikuliLibrary.Exists    ${incorrect_password_image}
               IF    ${incorrect_password_exists}==True
               
                    SikuliLibrary.Click    ${incorrect_password_ok_image} 
                #-------------------------commented for DMS V3 integration----------------------------------------#    
                    # Fail
                #-------------------------commented for DMS V3 integration----------------------------------------#    
               END
               #----------------------------Extra added for incorrect password------------------------------------------#
               ${incorrect_password_exists}=   SikuliLibrary.Exists    ${incorrect_password_image}
               IF    ${incorrect_password_exists}==True

                    SikuliLibrary.Click    ${incorrect_password_ok_image} 
                #-------------------------commented for DMS V3 integration----------------------------------------#    
                    # Fail
                #-------------------------commented for DMS V3 integration----------------------------------------#    
                    RETURN    ${False}
               END
               #----------------------------Extra added for incorrect password------------------------------------------#

               Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${homeScreen_image}  ${time_out}
               #Sleep    ${Max_sleep}
               BREAK
          #-------------------------commented for DMS V3 integration----------------------------------------#
          # ELSE
          #-------------------------commented for DMS V3 integration----------------------------------------#
               #${refresh_exist}=    SikuliLibrary.Exists    ${edge_refresh_button} 
           
                #SikuliLibrary.Click    ${edge_refresh_button}
                #Sleep    2
          #-------------------------commented for DMS V3 integration----------------------------------------#
     #            RPA.Desktop.Press Keys   ctrl  r
     #            RPA.Desktop.Press Keys   alt   r 
     #            RPA.Desktop.Press Keys   ctrl  r
     #            RPA.Desktop.Press Keys   alt   r 
     #            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain      ${password_image}  ${time_out}
     #            ${password_image_exists}=    SikuliLibrary.Exists  ${password_image} 
     #        IF    ${password_image_exists}==False
                    
     #              Fail
     #        END 
     #          SikuliLibrary.Click     ${password_image}
     #          SikuliLibrary.Click     ${password_image}
     #          Sleep    1
     #         SikuliLibrary.Input Text    ${EMPTY}    ${password}

     #         Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${login_ok_button}     ${time_out}
     #         ${login_ok_button_exists}=    SikuliLibrary.Exists    ${login_ok_button}   
     #         IF    ${login_ok_button_exists}==False
                  
     #              Fail
     #         END
     #         SikuliLibrary.Click    ${login_ok_button}
     #         Sleep    2
     #         ${incorrect_password_exists}=   SikuliLibrary.Exists    ${incorrect_password_image}
     #         IF    ${incorrect_password_exists}==True
     #            SikuliLibrary.Click    ${incorrect_password_ok_image} 
                
     #            Fail
                 
     #         END
     #         Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${homeScreen_image}    ${time_out}
             
     #        BREAK
     #   END 
       #-------------------------commented for DMS V3 integration----------------------------------------# 
       ${counter}   Evaluate    ${counter} + 1
    END
    
    ${exist_home_screen}=   SikuliLibrary.Exists    ${homeScreen_image}  
    IF    ${exist_home_screen}==False
      Fail
    ELSE       
        RETURN    ${True} 
    END
    
  EXCEPT  AS   ${message}
        #DMS_String_Manipulation.Close Edge Browser
        #DMS_String_Manipulation.Close Edge Browser Tab
       
       ${status_counter}   Evaluate    ${status_counter} + 1
       Set Global Variable    ${status_counter}
       log   ${message}
       Capture Screenshot 
       Fail    ${message}
      
  END 
  
DMS_Login_With_Alternative_URL
  TRY
     
    ${user_id}   ${password}   ${alternative_url}   Login With Alternative Url    ${sheet_path}
    #To dynmaically generate sid
    ${sid}    Generate Current Date String
    #this will dynamically generate url

    ${url}   Generate Dynamic Url    ${user_id}    ${sid}  ${alternative_url} 
    Log    ${url}
    log    ${password}
    #To open edge with dynamic url
    Open Edge    ${url}
    sleep  2
    ${log_status}    Check Restore Button Exist
#     IF    "${log_status}" == "${True}"
#         Open Edge    ${url}
#         sleep  2
#     END
    #To check the run this time button exists
    Run Keyword And Ignore Error   SikuliLibrary.Wait Until Screen Contain    ${edge_run_this_time_button}    10
    ${button_exist}=    SikuliLibrary.Exists    ${edge_run_this_time_button}
    log    ${edge_run_this_time_button}
    #If the button doesnot exist it will retry 5 times to check it by refreshing the page
    ${counter}    Set Variable   1
    WHILE    ${counter}<=2
         
          
        IF    ${button_exist}==True
            SikuliLibrary.Click    ${edge_run_this_time_button}
               
            Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain      ${password_image}  90
            ${password_image_exists}=    SikuliLibrary.Exists  ${password_image} 
            IF    ${password_image_exists}==False
                    
                  Fail
                  
                
            END 
              SikuliLibrary.Click     ${password_image}
              SikuliLibrary.Click     ${password_image}
              Sleep    ${Min_sleep}
             SikuliLibrary.Input Text    ${EMPTY}    ${password}
             Sleep    ${Min_sleep}
             ${login_ok_button_exists}=    SikuliLibrary.Exists    ${login_ok_button}   
             IF    ${login_ok_button_exists}==False
                  
                  Fail

                 
             END
             SikuliLibrary.Click    ${login_ok_button}
             Sleep    2
             ${incorrect_password_exists}=   SikuliLibrary.Exists    ${incorrect_password_image}
             IF    ${incorrect_password_exists}==True
                SikuliLibrary.Click    ${incorrect_password_ok_image} 
                
                Fail
                 
             END
             Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${homeScreen_image}    ${time_out}

            BREAK
        ELSE
           ${refresh_exist}=    SikuliLibrary.Exists    ${edge_refresh_button} 
          IF    ${refresh_exist}==True
               #SikuliLibrary.Click    ${edge_refresh_button}

               RPA.Desktop.Press Keys   ctrl  r
               RPA.Desktop.Press Keys   alt   r
               RPA.Desktop.Press Keys   ctrl  r
               RPA.Desktop.Press Keys   alt   r
               Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain      ${password_image}    ${time_out}
               ${password_image_exists}=    SikuliLibrary.Exists  ${password_image} 
               IF    ${password_image_exists}==False
                         
                    Fail
               END 
              SikuliLibrary.Click     ${password_image}
              SikuliLibrary.Click     ${password_image}
              Sleep    ${Min_sleep}
             SikuliLibrary.Input Text    ${EMPTY}    ${password}
             Sleep    ${Min_sleep}
             ${login_ok_button_exists}=    SikuliLibrary.Exists    ${login_ok_button}   
             IF    ${login_ok_button_exists}==False
                  
                  Fail
             END
             SikuliLibrary.Click    ${login_ok_button}
             Sleep    2
             ${incorrect_password_exists}=   SikuliLibrary.Exists    ${incorrect_password_image}
             IF    ${incorrect_password_exists}==True
                SikuliLibrary.Click    ${incorrect_password_ok_image} 
                
                Fail
                 
             END
             Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${homeScreen_image}    ${time_out}
            
            BREAK
                
           END
        
       END  
       ${counter}   Evaluate    ${counter} + 1
    END
    
    ${exist_home_screen}=   SikuliLibrary.Exists    ${homeScreen_image}  
    IF    ${exist_home_screen}==False
     Log    ${exist_home_screen}
      Fail
        
    END
  EXCEPT  AS   ${message}
     #    DMS_String_Manipulation.Close Edge Browser
        
       ${status_counter}   Evaluate    ${status_counter} + 1
       Set Global Variable    ${status_counter}
       log   ${message}
       Capture Screenshot 
       Fail
      
  END 
  

Login To DMS
    TRY
         Read_credentials
         #------------------------------------------DMS V3 code inegration---------------------------------#
        #  ${Execution_Status}=   Run Keyword And Return Status    Wait Until Keyword Succeeds    ${retry}    ${average_sleep}     dms_login
          # ${Execution_Status}=   Run Keyword And Return Status    dms_login
          ${Execution_Status}=   dms_login
          IF    ${Execution_Status} == ${False}
              RETURN    ${False}
          ELSE    
              RETURN    ${True}
          END

         #------------------------------------------DMS V3 code inegration---------------------------------#
          #------------------------------------------DMS V3 code inegration---------------------------------#
          # IF    ${Execution_Status} == False
          #      DMS Login with alternative url 
          #      ${read_credentials_status}=   Run Keyword And Return Status    Read_credentials  
          #      IF    ${read_credentials_status}==True
          #           ${alternative_login_execution_Status}=   Run Keyword And Return Status    Wait Until Keyword Succeeds    ${retry}    ${average_sleep}     DMS_Login_With_Alternative_URL 
          #                Sleep    1
          #      ELSE
          #           Fail    
          #      END
          #      IF   ${alternative_login_execution_Status}== False 
          #           Show Message Popup    Alert    DMS Login has been failed. Please Manually login to DMS
          #           Fail    DMS Login has been failed. Please Manually login to DMS
               
          #      END 
          # END
          #------------------------------------------DMS V3 code inegration---------------------------------#
     EXCEPT  AS   ${error_message}           
        log    ${error_message}
        Capture Screenshot 
        Fail    ${error_message} 
     END

Check Restore Button Exist  
     ${restore_button_image_exist}    SikuliLibrary.Exists  ${restore_button_image} 
     IF    ${restore_button_image_exist} == True
         Run Keyword And Ignore Error    SikuliLibrary.Wait Until Screen Contain    ${restore_close_image}     ${time_out}
         ${restore_close_image_exist}    SikuliLibrary.Exists  ${restore_close_image}
          IF    ${restore_close_image_exist} == True
              SikuliLibrary.Click    ${restore_close_image}
          END
     END
