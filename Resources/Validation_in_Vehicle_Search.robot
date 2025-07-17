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
${ok_to_tabs}    1|3|2
${log_folder}    ${CURDIR}${/}..\\Screenshot
${chassis_title}    1|1|1|2
${reg_title}    1|1|1|1 
${mob_title}    1|1|1|5        
# ${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx
${branch_mapping}    ${CURDIR}${/}..\\Mapping\\Location Mapping DMS ERP.xlsx

*** Keywords ***

Validation in Registration Number Search
    [Arguments]    ${Chassis_No_}    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}

    TRY
        ${index}=    Set Variable    0
        ${adjusted_index}=    Set Variable    1
        

        WHILE    True

            ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|2     Name  
            IF    ${attr_status} == ${True}
                ${chas_atr_name}    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|2     Name
            ELSE
                ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|2    Name
                IF    ${attr_status} == ${True}
                    ${chas_atr_name}    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|2     Name
                END
            END 

            IF    '${chas_atr_name}' != 'Chassis Number new item row'

                ${element_name}=    Set Variable    Chassis Number row${index}

                Click Action    name:"${element_name}"
                RPA.Desktop.Press Keys    ctrl  a

                ${chassis_search_value}=    Get Value Action    name:"${element_name}"
                Log    Row ${index} Value: ${chassis_search_value}


                # Check for end of list
                IF    '${chassis_search_value}' == '' or '${chassis_search_value}' == 'None' or '${chassis_search_value}' == '(null)'
                    Log    No more chassis numbers found. Stopping.
                    RETURN    ${chassis_search_value}
                END
                
                IF    $Chassis_No_ in $chassis_search_value
                    Log    Match found at row ${index}
                    RETURN    ${Chassis_No_}              
                END

                # Check for match
                IF    '${chassis_search_value}' == '${Chassis_No_}'
                    Log    Match found at row ${index}
                    RETURN    ${chassis_search_value}
                END

                Press Keys Action    down
                ${index}=    Evaluate    ${index} + 1
                ${adjusted_index}=    Evaluate    ${adjusted_index} + 1

            ELSE

                RETURN  ${None}

            END
                
        
        END

    EXCEPT  AS   ${message}
        Log    Error encountered: ${message}
        Capture Screenshot

    END


Validation in Chassis Number Search
    [Arguments]    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}

    TRY
        ${index}=    Set Variable    0
        ${adjusted_index}=    Set Variable    1

        WHILE    True

            ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|1     Name   
            IF    ${attr_status} == ${True}
                ${reg_atr_name}    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|1     Name
            ELSE
                ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|1     Name
                IF    ${attr_status} == ${True}
                    ${reg_atr_name}    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|1     Name
                END
            END 

            IF    '${reg_atr_name}' != 'Registration No new item row'

                ${element_name}=    Set Variable    Registration No row${index}

                Click Action    name:"${element_name}"
                RPA.Desktop.Press Keys    ctrl  a

                ${reg_search_value}=    Get Value Action    name:"${element_name}"
                Log    Row ${index} Value: ${reg_search_value}

                # Check for end of list
                IF    '${reg_search_value}' == '' or '${reg_search_value}' == 'None' or '${reg_search_value}' == '(null)'
                    Log    No more registration numbers found. Stopping.
                    RETURN    ${reg_search_value}
                END

                # Check for match
                IF    '${reg_search_value}' == '${Registration No.}'
                    Log    Match found at row ${index}
                    RETURN    ${reg_search_value}
                # END
                ELSE          
                    # ${updated_erp_reg_number}    ERP_methods.Correct Registration Number From Var    ${reg_search_value}
                    ${match_staus}    ${updated_erp_reg_number}    ERP_methods.vehicle_number_found    ${Registration No.}    ${reg_search_value}
                    
                    IF    '${updated_erp_reg_number}' == '${Registration No.}'
                        Log    Match found at row ${index}
                        RETURN    ${updated_erp_reg_number}
                    END
                END

                Press Keys Action    down
                ${index}=    Evaluate    ${index} + 1
                ${adjusted_index}=    Evaluate    ${adjusted_index} + 1
           
            ELSE

                RETURN  ${None}

            END

        END

    EXCEPT  AS   ${message}
        Log    Error encountered: ${message}
        Capture Screenshot

    END
Validation in Mobile Number Search
    [Arguments]    ${Chassis No.}    ${Registration No.}    ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description}

    TRY
        ${index}=    Set Variable    0
        ${adjusted_index}=    Set Variable    1

        WHILE    True

            # IF    "${Service Type Description}" == "BODY REPAIR" or "${Service Type Description}" == "BANDP"
            #     ${chas_attr_name}    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|2     Name
            # ELSE
            #     ${chas_attr_name}    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|2     Name
            # END

            # IF    "${Service Type Description}" == "BODY REPAIR" or "${Service Type Description}" == "BANDP"
            #     ${reg_attr_name}    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|1     Name
            # ELSE
            #     ${reg_attr_name}    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|1     Name
            # END
            ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|2     Name  
            IF    ${attr_status} == ${True}
                ${chas_attr_name}    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|2     Name
            ELSE
                ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|2    Name
                IF    ${attr_status} == ${True}
                    ${chas_attr_name}    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|2     Name
                END
            END 

            ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|1     Name   
            IF    ${attr_status} == ${True}
                ${reg_attr_name}    Get Attribute Action Validation    path:1|1|1|2|${adjusted_index}|1     Name
            ELSE
                ${attr_status}  Run Keyword And Return Status    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|1     Name
                IF    ${attr_status} == ${True}
                    ${reg_attr_name}    Get Attribute Action Validation    path:1|1|2|${adjusted_index}|1     Name
                END
            END 

            # Check if reached new item row
            IF    '${chas_attr_name}' == 'Chassis Number new item row' or '${reg_attr_name}' == 'Registration No new item row'
                Log    Reached new item row, stopping loop.
                RETURN    ${None}
            END

            ${chassis_element}=    Set Variable    Chassis Number row${index}
            ${reg_element}=        Set Variable    Registration No row${index}

            # Get and check Chassis Number
            Click Action    name:"${chassis_element}"
            RPA.Desktop.Press Keys    ctrl  a
            ${chassis_search_value}=    Get Value Action    name:"${chassis_element}"
            Log    Row ${index} Chassis Value: ${chassis_search_value}

            # Get and check Registration Number
            Click Action    name:"${reg_element}"
            RPA.Desktop.Press Keys    ctrl  a
            ${reg_search_value}=    Get Value Action    name:"${reg_element}"
            Log    Row ${index} Registration Value: ${reg_search_value}

            # Check if either value is empty/null
            IF    '${chassis_search_value}' == '' or '${chassis_search_value}' == 'None' or '${chassis_search_value}' == '(null)' or '${reg_search_value}' == '' or '${reg_search_value}' == 'None' or '${reg_search_value}' == '(null)'
                Log    End of data reached. No match found.
                RETURN    ${None}
            END

            # Check for match
            IF    '${chassis_search_value}' == '${Chassis No.}' and '${reg_search_value}' == '${Registration No.}'
                Log    Match found at row ${index}
                RETURN    ${chassis_search_value}    ${reg_search_value}
            END

            # Move to next row
            Press Keys Action    down
            ${index}=    Evaluate    ${index} + 1
            ${adjusted_index}=    Evaluate    ${adjusted_index} + 1

        END

    EXCEPT  AS   ${message}
        Log    Error encountered: ${message}
        Capture Screenshot

    END




























