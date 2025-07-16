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
${Service_Type_Sheet}       ${CURDIR}${/}..\\Mapping\\Service Type Cumulative.xlsx
${advisor_list_sheet}       ${CURDIR}${/}..\\Mapping\\Service Advisor List ERP SRM_ELM.xlsx
# ${log_folder}     ${CURDIR}${/}..\\Log
${log_folder}     C:${/}JobcardOpeningIntegrated\\Screenshot
# ${log_folder}     C:\\Users\popular\Desktop\JobcardOpeningIntegrated\Jobcard-Opening\Screenshot
${prefix}    Jobcard already opened for
${body_check_list_type}
${substring}    Jobcard already opened 
${branch_mapping}    Mapping//Location Mapping DMS ERP.xlsx
${pickup_driver_title}    3|1|1|1|2|1|1|1|1|1|1|1|1
${pickup_driver_inp_field}    3|1|1|1|2|1|1|1|1|1|1|1|2|1|1  
${pickupdate_title}    3|1|1|1|2|1|1|1|1|1|1|2|1
${pickupdate_inp_field}    3|1|1|1|2|1|1|1|1|1|1|2|2|1  
${pickuptime_title}    3|1|1|1|2|1|1|1|1|1|1|3|1  
${disancekm_title}    3|1|1|1|2|1|1|1|1|1|1|5|1
${distancekm_inp_field}    3|1|1|1|2|1|1|1|1|1|1|5|2|1
${pickupaddress_title}    3|1|1|1|2|1|1|1|1|1|1|4|1
${comments_title}    3|1|1|1|2|1|1|1|1|1|1|6|1
${comments_inp_field}    3|1|1|1|2|1|1|1|1|1|1|6|2|1       
${pickup_address}    3|1|1|1|2|1|1|1|1|1|1|4|2|1                 
${picker_details_mapping}    C:\\JobcardOpeningIntegrated\\Mapping\\PickUp_Details.xlsx

*** Keywords ***

Pickup Details Tab
    [Arguments]  ${Job Card No.}    ${Input_Sheet_Path}    ${Service Type Description} 
    # [Arguments]  ${Job Card No.}   ${Service Type Description}  
    
    TRY
        ${Service Type Description_value}    Clean String    ${Service Type Description}
        IF    "${Service Type Description_value}" == "bodyrepair" or "${Service Type Description_value}" == "bandp"

            Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f3

            #extra added to handle present odometer popup
            Run Keyword And Ignore Error    Additional Windows  

        ELSE
            # Window Navigation    Wings ERP 23E - Web Client
            Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ctrl  f8
        
        END
        # Details Extracion from Excel row Mapping
        ${row_data}    ${row_status}    ERP_methods.Get Pickup Details    ${picker_details_mapping}    ${Job Card No.}
        Log    ${row_data}
        Log    ${row_status}



        #if there is no particular row for that jobcard, give default no pickup data
        IF    "${row_status}" == "${None}"   

            # RETURN    Pick Up Details not updated in excel sheet

            #pickup driver name/No pickup/Employee pickup
            ${pickupdriver_title_value}    Get Text Action    path:${pickup_driver_title}
            Log    ${pickupdriver_title_value}
            ${pickupdriver_title_value_lower}    Convert To Lower Case    ${pickupdriver_title_value}
            IF    "${pickupdriver_title_value_lower}" == "pick up driver *" or "${pickupdriver_title_value_lower}" == "pick up driver"
        
                Click Action    path:${pickup_driver_inp_field} 
                RPA.Desktop.Press Keys    ctrl  a  
                Type Text Action    NO PICKUP
                Capture Screenshot
                Press Keys Action    enter

            END    
            Capture Screenshot

            #pickup date
            ${pickupdate_title_value}    Get Text Action    path:${pickupdate_title}
            Log    ${pickupdate_title_value}
            ${pickupdate_title_value_lower}    Convert To Lower Case    ${pickupdate_title_value}
            IF    '${pickupdate_title_value_lower}' == 'pick up date *' or '${pickupdate_title_value_lower}' == 'pick up date'

                Click Action    path:${pickupdate_inp_field}
                RPA.Desktop.Press Keys    ctrl  a    
                ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                Type Text Action    ${cur_date}          
                Press Keys Action    tab
                Capture Screenshot
            END          
            Capture Screenshot

            #pickup time
            ${pickuptime_title_value}    Get Text Action    path:${pickuptime_title}
            Log    ${pickuptime_title_value}
            ${pickuptime_title_value_lower}    Convert To Lower Case    ${pickuptime_title_value}
            IF    '${pickuptime_title_value_lower}' == 'pick up time *' or '${pickuptime_title_value_lower}' == 'pick up time'


                ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                Type Text Action    ${cur_time}
                Capture Screenshot
                Press Keys Action    tab   #will come to address field

            END 
            Capture Screenshot

            #pickup address
            ${pickupaddress_title_value}    Get Text Action    path:${pickupaddress_title}
            Log    ${pickupaddress_title_value}
            ${pickupaddress_title_value_lower}    Convert To Lower Case    ${pickupaddress_title_value}
            IF    '${pickupaddress_title_value_lower}' == 'pick up address'

                Click Action    path:${pickup_address}
                Press Keys Action    tab    #come to distance field
                Capture Screenshot

            END 
            Capture Screenshot

            #Entering distance value
            ${km_title_value}    Get Text Action    path:${disancekm_title}
            Log    ${km_title_value}
            ${km_title_value_lower}    Convert To Lower Case    ${km_title_value}
            IF    '${km_title_value_lower}' == 'distance in km *' or '${km_title_value_lower}' == 'distance in km'
                
                RPA.Desktop.Press Keys    ctrl  a 
                Press Keys Action    delete
                Click Action    path:${distancekm_inp_field}
                # Log    ${row_data}[Distance in KM]
                Type Text Action    1
                Capture Screenshot
                Press Keys Action    tab  #come to comments 

            END 
            Capture Screenshot

            #Entering comments
            ${comments_title_value}    Get Text Action    path:${comments_title}
            Log    ${comments_title_value}
            ${comments_title_value_lower}    Convert To Lower Case    ${comments_title_value}
            IF    '${comments_title_value_lower}' == 'comments'
                
                Click Action    path:${comments_inp_field}
                Press Keys Action    enter
                Capture Screenshot
            END 
            Capture Screenshot


            RETURN    ${True}
            

        #if any mandatory value is missing
        ELSE IF  """${row_status}""" == """employee pickup""" or """${row_status}""" == """employee pickup with some mandatory data is missing"""

            #pickup driver name/No pickup/Employee pickup
            ${pickupdriver_title_value}    Get Text Action    path:${pickup_driver_title}
            Log    ${pickupdriver_title_value}
            ${pickupdriver_title_value_lower}    Convert To Lower Case    ${pickupdriver_title_value}
            IF    "${pickupdriver_title_value_lower}" == "pick up driver *" or "${pickupdriver_title_value_lower}" == "pick up driver"
        
                Click Action    path:${pickup_driver_inp_field} 
                RPA.Desktop.Press Keys    ctrl  a  
                Log    ${row_data}[Pick Up Driver]
                Type Text Action    ${row_data}[Pick Up Driver]
                Capture Screenshot
                Press Keys Action    enter

            END    
            Capture Screenshot

            #pickup date
            ${pickupdate_title_value}    Get Text Action    path:${pickupdate_title}
            Log    ${pickupdate_title_value}
            ${pickupdate_title_value_lower}    Convert To Lower Case    ${pickupdate_title_value}
            IF    '${pickupdate_title_value_lower}' == 'pick up date *' or '${pickupdate_title_value_lower}' == 'pick up date'

                Click Action    path:${pickupdate_inp_field}
                RPA.Desktop.Press Keys    ctrl  a 
                Log    ${row_data}[Pick Up Date]  
                IF    '${row_data}[Pick Up Date]' != '${None}'

                    Type Text Action    ${row_data}[Pick Up Date]    
                    Press Keys Action    tab
                    Capture Screenshot
                ELSE
                    ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                    Type Text Action    ${cur_date}          
                    Press Keys Action    tab
                    Capture Screenshot
                END  
                
            END          
            Capture Screenshot

            #pickup time
            ${pickuptime_title_value}    Get Text Action    path:${pickuptime_title}
            Log    ${pickuptime_title_value}
            ${pickuptime_title_value_lower}    Convert To Lower Case    ${pickuptime_title_value}
            IF    '${pickuptime_title_value_lower}' == 'pick up time *' or '${pickuptime_title_value_lower}' == 'pick up time'

                Log    ${row_data}[Pick Up Time]
                IF    '${row_data}[Pick Up Time]' != '${None}'
                    # Click Action    path:3|1|1|1|2|1|1|1|1|1|1|3|2|1
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    Type Text Action    ${row_data}[Pick Up Time]
                    Capture Screenshot
                    Press Keys Action    tab   #will come to address field
                ELSE
                    # Click Action    path:3|1|1|1|2|1|1|1|1|1|1|3|2|1
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    # Press Keys Action    backspace
                    ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                    Type Text Action    ${cur_time}
                    Capture Screenshot
                    Press Keys Action    tab   #will come to address field
                END
                
            END 
            Capture Screenshot

            #pickup address
            ${pickupaddress_title_value}    Get Text Action    path:${pickupaddress_title}
            Log    ${pickupaddress_title_value}
            ${pickupaddress_title_value_lower}    Convert To Lower Case    ${pickupaddress_title_value}
            IF    '${pickupaddress_title_value_lower}' == 'pick up address'

                Click Action    path:${pickup_address}
                # RPA.Desktop.Press Keys    ctrl  a   
                Log    ${row_data}[Pick Up Address]
                IF    '${row_data}[Pick Up Address]' == 'None'
                    Press Keys Action    tab    #come to distance field
                    Capture Screenshot
                ELSE
                    Type Text Action    ${row_data}[Pick Up Address]
                    Capture Screenshot
                    Press Keys Action    tab  #come to distance field
                END

            END 
            Capture Screenshot

            #Entering distance value
            ${km_title_value}    Get Text Action    path:${disancekm_title}
            Log    ${km_title_value}
            ${km_title_value_lower}    Convert To Lower Case    ${km_title_value}
            IF    '${km_title_value_lower}' == 'distance in km *' or '${km_title_value_lower}' == 'distance in km'
                
                RPA.Desktop.Press Keys    ctrl  a 
                Press Keys Action    delete
                Click Action    path:${distancekm_inp_field}
                Log    ${row_data}[Distance in KM]
                IF    '${row_data}[Distance in KM]' != '${None}'
                    Type Text Action    ${row_data}[Distance in KM]
                    Capture Screenshot
                    Press Keys Action    tab  #come to comments 
                ELSE
                    Type Text Action    1
                    Capture Screenshot
                    Press Keys Action    tab  #come to comments 
                END
                
            END 
            Capture Screenshot

            #Entering comments
            ${comments_title_value}    Get Text Action    path:${comments_title}
            Log    ${comments_title_value}
            ${comments_title_value_lower}    Convert To Lower Case    ${comments_title_value}
            IF    '${comments_title_value_lower}' == 'comments'
                
                Click Action    path:${comments_inp_field}
                Log    ${row_data}[Comments]
                IF    '${row_data}[Comments]' == 'None'
                    Press Keys Action    enter
                    Capture Screenshot
                ELSE
                   Type Text Action    ${row_data}[Comments]
                    Press Keys Action    enter
                    Capture Screenshot
                END
                
                
            END 
            Capture Screenshot

            RETURN    ${True}

        ELSE IF  """${row_status}""" == """no pickup"""

            #pickup driver name/No pickup/Employee pickup
            ${pickupdriver_title_value}    Get Text Action    path:${pickup_driver_title}
            Log    ${pickupdriver_title_value}
            ${pickupdriver_title_value_lower}    Convert To Lower Case    ${pickupdriver_title_value}
            IF    "${pickupdriver_title_value_lower}" == "pick up driver *" or "${pickupdriver_title_value_lower}" == "pick up driver"
        
                Click Action    path:${pickup_driver_inp_field} 
                RPA.Desktop.Press Keys    ctrl  a  
                Log    ${row_data}[Pick Up Driver]
                Type Text Action    ${row_data}[Pick Up Driver]
                Capture Screenshot
                Press Keys Action    enter

            END    
            Capture Screenshot

            #pickup date
            ${pickupdate_title_value}    Get Text Action    path:${pickupdate_title}
            Log    ${pickupdate_title_value}
            ${pickupdate_title_value_lower}    Convert To Lower Case    ${pickupdate_title_value}
            IF    '${pickupdate_title_value_lower}' == 'pick up date *' or '${pickupdate_title_value_lower}' == 'pick up date'

                Click Action    path:${pickupdate_inp_field}
                RPA.Desktop.Press Keys    ctrl  a  
                # Log    ${row_data}[Pick Up Date]    
                ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                Type Text Action    ${cur_date}          
                Press Keys Action    tab
                Capture Screenshot
            END          
            Capture Screenshot

            #pickup time
            ${pickuptime_title_value}    Get Text Action    path:${pickuptime_title}
            Log    ${pickuptime_title_value}
            ${pickuptime_title_value_lower}    Convert To Lower Case    ${pickuptime_title_value}
            IF    '${pickuptime_title_value_lower}' == 'pick up time *' or '${pickuptime_title_value_lower}' == 'pick up time'

                ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                Type Text Action    ${cur_time}
                Capture Screenshot
                Press Keys Action    tab   #will come to address field

            END 
            Capture Screenshot

            #pickup address
            ${pickupaddress_title_value}    Get Text Action    path:${pickupaddress_title}
            Log    ${pickupaddress_title_value}
            ${pickupaddress_title_value_lower}    Convert To Lower Case    ${pickupaddress_title_value}
            IF    '${pickupaddress_title_value_lower}' == 'pick up address'

                Click Action    path:${pickup_address}
                # RPA.Desktop.Press Keys    ctrl  a   
                Log    ${row_data}[Pick Up Address]
                IF    '${row_data}[Pick Up Address]' == 'None'
                    Press Keys Action    tab    #come to distance field
                    Capture Screenshot
                ELSE
                    Type Text Action    ${row_data}[Pick Up Address]
                    Capture Screenshot
                    Press Keys Action    tab  #come to distance field
                END

            END 
            Capture Screenshot

            #Entering distance value
            ${km_title_value}    Get Text Action    path:${disancekm_title}
            Log    ${km_title_value}
            ${km_title_value_lower}    Convert To Lower Case    ${km_title_value}
            IF    '${km_title_value_lower}' == 'distance in km *' or '${km_title_value_lower}' == 'distance in km'
                
                RPA.Desktop.Press Keys    ctrl  a 
                Press Keys Action    delete
                Click Action    path:${distancekm_inp_field}
                # Log    ${row_data}[Distance in KM]
                Type Text Action    1
                Capture Screenshot
                Press Keys Action    tab  #come to comments 

            END 
            Capture Screenshot

            #Entering comments
            ${comments_title_value}    Get Text Action    path:${comments_title}
            Log    ${comments_title_value}
            ${comments_title_value_lower}    Convert To Lower Case    ${comments_title_value}
            IF    '${comments_title_value_lower}' == 'comments'
                
                Click Action    path:${comments_inp_field}
                Log    ${row_data}[Comments]
                IF    '${row_data}[Comments]' == 'None'
                    Press Keys Action    enter
                    Capture Screenshot
                ELSE
                   Type Text Action    ${row_data}[Comments]
                    Press Keys Action    enter
                    Capture Screenshot
                END
                
                
            END 
            Capture Screenshot


            RETURN    ${True}


        ELSE IF  """${row_status}""" == """Some mandatory values except Pickup Driver is missing"""    


            RETURN    Some mandatory values except Pickup Driver is missing    


        ELSE IF  """${row_status}""" == """Pickup Driver is missing"""

            #pickup driver name/No pickup/Employee pickup
            ${pickupdriver_title_value}    Get Text Action    path:${pickup_driver_title}
            Log    ${pickupdriver_title_value}
            ${pickupdriver_title_value_lower}    Convert To Lower Case    ${pickupdriver_title_value}
            IF    "${pickupdriver_title_value_lower}" == "pick up driver *" or "${pickupdriver_title_value_lower}" == "pick up driver"
        
                Click Action    path:${pickup_driver_inp_field} 
                RPA.Desktop.Press Keys    ctrl  a  
                Log    ${row_data}[Pick Up Driver]
                Type Text Action    NO PICKUP
                Capture Screenshot
                Press Keys Action    enter

            END    
            Capture Screenshot

            #pickup date
            ${pickupdate_title_value}    Get Text Action    path:${pickupdate_title}
            Log    ${pickupdate_title_value}
            ${pickupdate_title_value_lower}    Convert To Lower Case    ${pickupdate_title_value}
            IF    '${pickupdate_title_value_lower}' == 'pick up date *' or '${pickupdate_title_value_lower}' == 'pick up date'

                Click Action    path:${pickupdate_inp_field}
                RPA.Desktop.Press Keys    ctrl  a  
                # Log    ${row_data}[Pick Up Date]    
                ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                Type Text Action    ${cur_date}          
                Press Keys Action    tab
                Capture Screenshot
            END          
            Capture Screenshot

            #pickup time
            ${pickuptime_title_value}    Get Text Action    path:${pickuptime_title}
            Log    ${pickuptime_title_value}
            ${pickuptime_title_value_lower}    Convert To Lower Case    ${pickuptime_title_value}
            IF    '${pickuptime_title_value_lower}' == 'pick up time *' or '${pickuptime_title_value_lower}' == 'pick up time'

                # Click Action    path:3|1|1|1|2|1|1|1|1|1|1|3|2|1
                # RPA.Desktop.Press Keys    ctrl  a  
                # Press Keys Action    left  
                # Log    ${row_data}[Pick Up Time]
                ${cur_date}    ${cur_time}    ERP_methods.Get Current Date And Time Pickup
                Type Text Action    ${cur_time}
                Capture Screenshot
                Press Keys Action    tab   #will come to address field

            END 
            Capture Screenshot

            #pickup address
            ${pickupaddress_title_value}    Get Text Action    path:${pickupaddress_title}
            Log    ${pickupaddress_title_value}
            ${pickupaddress_title_value_lower}    Convert To Lower Case    ${pickupaddress_title_value}
            IF    '${pickupaddress_title_value_lower}' == 'pick up address'

                Click Action    path:${pickup_address}
                # RPA.Desktop.Press Keys    ctrl  a   
                Log    ${row_data}[Pick Up Address]
                IF    '${row_data}[Pick Up Address]' == 'None'
                    Press Keys Action    tab    #come to distance field
                    Capture Screenshot
                ELSE
                    Type Text Action    ${row_data}[Pick Up Address]
                    Capture Screenshot
                    Press Keys Action    tab  #come to distance field
                END

            END 
            Capture Screenshot

            #Entering distance value
            ${km_title_value}    Get Text Action    path:${disancekm_title}
            Log    ${km_title_value}
            ${km_title_value_lower}    Convert To Lower Case    ${km_title_value}
            IF    '${km_title_value_lower}' == 'distance in km *' or '${km_title_value_lower}' == 'distance in km'
                
                RPA.Desktop.Press Keys    ctrl  a 
                Press Keys Action    delete
                Click Action    path:${distancekm_inp_field}
                # Log    ${row_data}[Distance in KM]
                Type Text Action    1
                Capture Screenshot
                Press Keys Action    tab  #come to comments 

            END 
            Capture Screenshot

            #Entering comments
            ${comments_title_value}    Get Text Action    path:${comments_title}
            Log    ${comments_title_value}
            ${comments_title_value_lower}    Convert To Lower Case    ${comments_title_value}
            IF    '${comments_title_value_lower}' == 'comments'
                
                Click Action    path:${comments_inp_field}
                Log    ${row_data}[Comments]
                IF    '${row_data}[Comments]' == 'None'
                    Press Keys Action    enter
                    Capture Screenshot
                ELSE
                   Type Text Action    ${row_data}[Comments]
                    Press Keys Action    enter
                    Capture Screenshot
                END
                
            END 
            Capture Screenshot


            RETURN    ${True}


        ELSE IF  """${row_status}""" == """fields extracted"""

            #pickup driver name/No pickup/Employee pickup
            ${pickupdriver_title_value}    Get Text Action    path:${pickup_driver_title}
            Log    ${pickupdriver_title_value}
            ${pickupdriver_title_value_lower}    Convert To Lower Case    ${pickupdriver_title_value}
            IF    "${pickupdriver_title_value_lower}" == "pick up driver *" or "${pickupdriver_title_value_lower}" == "pick up driver"
        
                Click Action    path:${pickup_driver_inp_field} 
                RPA.Desktop.Press Keys    ctrl  a  
                Log    ${row_data}[Pick Up Driver]
                Type Text Action    ${row_data}[Pick Up Driver]
                Capture Screenshot
                Press Keys Action    enter

            END    
            Capture Screenshot

            #pickup date
            ${pickupdate_title_value}    Get Text Action    path:${pickupdate_title}
            Log    ${pickupdate_title_value}
            ${pickupdate_title_value_lower}    Convert To Lower Case    ${pickupdate_title_value}
            IF    '${pickupdate_title_value_lower}' == 'pick up date *' or '${pickupdate_title_value_lower}' == 'pick up date'

                Click Action    path:${pickupdate_inp_field}
                RPA.Desktop.Press Keys    ctrl  a  
                Log    ${row_data}[Pick Up Date]    
                Type Text Action    ${row_data}[Pick Up Date]           
                Press Keys Action    tab
                Capture Screenshot
            END          
            Capture Screenshot

            #pickup time
            ${pickuptime_title_value}    Get Text Action    path:${pickuptime_title}
            Log    ${pickuptime_title_value}
            ${pickuptime_title_value_lower}    Convert To Lower Case    ${pickuptime_title_value}
            IF    '${pickuptime_title_value_lower}' == 'pick up time *' or '${pickuptime_title_value_lower}' == 'pick up time'

                # Click Action    path:3|1|1|1|2|1|1|1|1|1|1|3|2|1
                # RPA.Desktop.Press Keys    ctrl  a  
                # Press Keys Action    left  
                Log    ${row_data}[Pick Up Time]
                Type Text Action    ${row_data}[Pick Up Time]
                Capture Screenshot
                Press Keys Action    tab   #will come to address field

            END 
            Capture Screenshot

            #pickup address
            ${pickupaddress_title_value}    Get Text Action    path:${pickupaddress_title}
            Log    ${pickupaddress_title_value}
            ${pickupaddress_title_value_lower}    Convert To Lower Case    ${pickupaddress_title_value}
            IF    '${pickupaddress_title_value_lower}' == 'pick up address'

                Click Action    path:${pickup_address}
                # RPA.Desktop.Press Keys    ctrl  a   
                Log    ${row_data}[Pick Up Address]
                IF    '${row_data}[Pick Up Address]' == 'None'
                    Press Keys Action    tab    #come to distance field
                    Capture Screenshot
                ELSE
                    Type Text Action    ${row_data}[Pick Up Address]
                    Capture Screenshot
                    Press Keys Action    tab  #come to distance field
                END
            END 
            Capture Screenshot

            #Entering distance value
            ${km_title_value}    Get Text Action    path:${disancekm_title}
            Log    ${km_title_value}
            ${km_title_value_lower}    Convert To Lower Case    ${km_title_value}
            IF    '${km_title_value_lower}' == 'distance in km *' or '${km_title_value_lower}' == 'distance in km'
                
                RPA.Desktop.Press Keys    ctrl  a 
                Press Keys Action    delete
                Click Action    path:${distancekm_inp_field}
                Log    ${row_data}[Distance in KM]
                Type Text Action    ${row_data}[Distance in KM]
                Capture Screenshot
                Press Keys Action    tab  #come to comments 

            END 
            Capture Screenshot

            #Entering comments
            ${comments_title_value}    Get Text Action    path:${comments_title}
            Log    ${comments_title_value}
            ${comments_title_value_lower}    Convert To Lower Case    ${comments_title_value}
            IF    '${comments_title_value_lower}' == 'comments'
                
                Click Action    path:${comments_inp_field}
                Log    ${row_data}[Comments]
                IF    '${row_data}[Comments]' == 'None'
                    Press Keys Action    enter
                    Capture Screenshot
                ELSE
                   Type Text Action    ${row_data}[Comments]
                    Press Keys Action    enter
                    Capture Screenshot
                END
                
            END 
            Capture Screenshot

            RETURN    ${True}

        
        END     

        

    EXCEPT  AS  ${Pickup_Details_Tab_error}
        Log  ${Pickup_Details_Tab_error}
        Capture Screenshot
        # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Error occurred while interacting with the Pickup Details Tab.    ${exception_reason_column_name}
        # update_execution_status_in_summary_report    ${Input_Sheet_Path}    ${Job Card No.}    Fail    ${execution_status_column_name}
        # Exit For Loop
        
    END   

# *** Tasks ***  
# Demo  
#     Pickup Details Tab    JC25001866    RR
 