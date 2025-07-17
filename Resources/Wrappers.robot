*** Settings ***
Library             RPA.Browser.Selenium    auto_close=${FALSE}
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
Resource            Jobcard_Navigations.robot
Resource            Main_Flow.robot
Resource            JobCard_Tab_Data_entry.robot


*** Variables ***
${log_folder}     ${CURDIR}${/}..\\Screenshot



*** Keywords ***

Click Action
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Windows.Click    ${locator}

Selenium Click Element 
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Browser.Selenium.Click Element    ${locator}
    
Selenium Click Button
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Browser.Selenium.Click Button    ${locator}

Selenium Input Text
     [Arguments]    ${locator}    ${value}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Browser.Selenium.Input Text    ${locator}    ${value}


Click Action Maximum Retry
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${max_retry_interval}    ${Window_time}    RPA.Windows.Click    ${locator}

# Control Window Action
#     [Arguments]    ${locator}
#     Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    Control Window    ${locator}

Close Window title
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${min_retry_inerval}    ${Min_time}    Close Window    ${locator}

Closing Window Action
    [Arguments]    ${window_details}
        Sleep    ${normal_sleep}
        ${windows}=  List Windows
        FOR  ${window}  IN  @{windows}
        ${title}=   ERP_methods.Get Title Starting With    ${window}    ${window_details}
        log  ${title}
        Exit For Loop If    '${title}'!='None'
        END
        Close Window title    name:"${title}"

Type Text Action
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Type Text    ${locator}

# Get Text Action Maximum Retry
#     [Arguments]    ${locator}
#     ${text}=    Wait Until Keyword Succeeds   ${max_retry_interval}    ${text_Max_Time}    RPA.Windows.Get Text    ${locator}
#     RETURN    ${text}

Get Text Action Maximum Retry
    [Arguments]    ${locator}
    ${text}=    Set Variable    Start
    ${retry}=    Set Variable    0
    WHILE    '${text}' == 'Start' and ${retry} < 20
        ${text}=    Wait Until Keyword Succeeds   ${max_retry_interval}    ${text_Max_Time}    RPA.Windows.Get Text    ${locator}
        ${retry}=    Evaluate    ${retry} + 1
        Sleep    1
    END
    RETURN    ${text}

Right Click Action
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Windows.Right Click    ${locator}

Get Text Action
    [Arguments]    ${locator}
    ${text}=    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Windows.Get Text    ${locator}
    RETURN    ${text}

Selenium Get Text
    [Arguments]    ${locator}
    ${text}=    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Browser.Selenium.Get Text    ${locator}
    RETURN    ${text}

Get Value Action
    [Arguments]    ${locator}
    ${text}=    Wait Until Keyword Succeeds    ${max_retry_interval}    ${avg_time}    RPA.Windows.Get Value    ${locator}
    # Log To Console    ${text}
    RETURN    ${text}



Press Keys Action
    [Arguments]    ${keyword}


    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Desktop.Press Keys    ${keyword}

    
Filter Value Entering
    [Arguments]    ${value}

    Press Keys Action    tab
    Type Text Action    '='
    Press Keys Action    tab
    Type Text Action    ${value} 
    Sleep    1
    Press Keys Action    tab
 

Window Navigation
    [Arguments]    ${window_details}
        ${windows}=  List Windows
        FOR  ${window}  IN  @{windows}
        ${title}=   ERP_methods.Get Title Starting With    ${window}    ${window_details}
        log  ${title}
        Exit For Loop If    '${title}'!='None'
        END
        Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    Control Window    name:"${title}"

Control Window Action
    [Arguments]    ${window_details}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    Control Window    name:"${window_details}"

Closing ERP Window
        [Arguments]    ${window_details}
        ${windows}=  List Windows
        FOR  ${window}  IN  @{windows}
        ${title}=   ERP_methods.Get Title Starting With    ${window}    ${window_details}
        log  ${title}
        Exit For Loop If    '${title}'!='None'
        END
        Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Windows.Close Window    ${title}


Wait Action  
    [Arguments]    ${action}    ${value}
    Wait Until Keyword Succeeds    ${avg_time}    ${avg_retry_interval}    ${action}    ${value}

Capture Screenshot
    ${timestamp}    Get Current Date    result_format=%Y%m%d_%H%M%S
    ${screenshot}=   RPA.Desktop.Take Screenshot    path=popular_screenshot_${timestamp}.png
    move_screenshot    ${screenshot}    ${log_folder}

Wait For Window
    [Arguments]    ${keyword}
    Wait Until Keyword Succeeds    ${min_retry_inerval}    ${Window_time}    RPA.Windows.Click    ${keyword}

Double Click Action
    [Arguments]    ${locator}
    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    RPA.Windows.Double Click    ${locator}

Get Attribute Action  
    [Arguments]    ${locator}    ${type}
    ${text}=    Wait Until Keyword Succeeds    ${avg_retry_interval}    ${avg_time}    Get Attribute    ${locator}    ${type}
    RETURN    ${text}

Get Attribute Action Validation
    [Arguments]    ${locator}    ${type}
    ${text}=    Wait Until Keyword Succeeds    2x    3    Get Attribute    ${locator}    ${type}
    RETURN    ${text}
    