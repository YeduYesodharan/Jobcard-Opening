*** Settings ***
Documentation       Template robot main suite.
Library             OperatingSystem
Resource   Resources/dms_login.robot
Resource   Resources/dms_generate_jobcard_report.robot
Resource   Resources/dms_process_start.robot
Library    Libraries/business_operations.py
Resource   Resources/erp_recall_marking.robot
Resource   Resources/erp_initial_report_download.robot
Library    Libraries/mail_send.py
Library    Libraries/Post_request.py
Resource   Resources/Wrappers.robot
Variables  Variables/variables.py

*** Variables ***
${report_name}   Job Card Creation Process
${path}    ${EMPTY}
# ${sheet_path}            Config//Popular_Credentials.xlsx
${sheet_path}            C:\\JobcardOpeningIntegrated\\Config\\Popular_Credentials.xlsx
# ${region_mapping_sheet}            Mapping//Location Mapping DMS ERP.xlsxc
${region_mapping_sheet}            C:\\JobcardOpeningIntegrated\\Mapping\\Location Mapping DMS ERP.xlsx
# ${log_folder}                ${CURDIR}${/}${/}Log
${log_folder}     C:\\JobcardOpeningIntegrated\\Screenshot   

*** Tasks ***
#Main process start from here
Popular Jobcard Extraction Process
    TRY
        Set Global Variable    ${report_name}
        ${consolid_report_path_in_processed_erp_dir}    Set Variable    ${EMPTY}
        ${erp_combined_report_path}    Set Variable    ${EMPTY}

        ${location}    ${time_out}   read_dms_location_from_config    ${sheet_path}
        Set Global Variable    ${time_out}
        # ${erp_combined_report_path}    Set Variable   E:/JobcardOpeningIntegrated/ProcessRelatedFolders/2025-05-02/Downloads/ERP_Report_20250502124418.xlsx  
        
        # Current date generating function
        ${curr_date}    ${date_timestamp}    get_current_date_fun
        Set Global Variable    ${curr_date}  

        #-----Removing all the csv files from downloads folder -------
        remove_all_csv_files_from_downloads

        # ------ clear_log_folder -------
        Log    ${log_folder}
        # clear_log_folder     ${log_folder}
        
        #-----Keyword used to create the process related folders used for internal and external needs----
        #-----Argument: current_date-----
        Process Related Folder Creation    ${curr_date}
        
        #-----Keyword used to download Regular and Bodyshop reports from ERP----
        #-----Return Value: combined report path of both bodyshop job cards and regular jobcards-----
        ${erp_combined_report_path}    ERP Initial Report Download    ${path}    ${region_mapping_sheet}    ${location}  
        Log    ${erp_combined_report_path}
        
        #-----Keyword used extract jobcard details frm DMS----
        #-----Argument: current_date and combiob@4312ned report path of both bodyshop job cards and regular jobcards-----
        ${consolid_report_path_in_processed_erp_dir}    ${carry_over_ratio_report_path}    DMS Data Extraction Process   ${curr_date}    ${erp_combined_report_path}    ${date_timestamp}

        #${consolid_report_path_in_processed_dms_dir}    Set Variable    E:\\JobcardOpeningIntegrated\\ProcessRelatedFolders\\2025-03-28\\ProcessedDMS\\consolidated_jobcard_report20250326161915.xlsx
        #${carry_over_ratio_report_path}    Set Variable    E:\\JobcardOpeningIntegrated\\ProcessRelatedFolders\\2025-03-28\\Downloads\\ELM.FOB20250326041905.xlsx

               
        ${consolidated_report_exist}    utility.check_consolidated_report_exist_and_empty    ${consolid_report_path_in_processed_erp_dir}
        
        IF    '${consolidated_report_exist}' == '${True}'

            ${consolidated_report_path_in_results_dir}    Copy Consolidated Report To Results Folder    ${consolid_report_path_in_processed_erp_dir}    ${curr_date}

            Copy Dms Report To Results Downloads Folder    ${carry_over_ratio_report_path}    ${curr_date}

            ${message}    Set Variable    ${EMPTY}
            Send Email Output     ${curr_date}    ${report_name}    ${message}    ${consolidated_report_path_in_results_dir}  

        END        
    EXCEPT  AS   ${message}
        log   ${message}
        Capture Screenshot 

        IF    """${message}""" == ""
            Set Variable    ${message}    Unexpected Error Occurred
        END

        IF    """${message}""" == "Unable to login to the DMS"
            Log    hi
            Send Email Output     ${curr_date}    ${report_name}    ${message}    ${EMPTY}
        ELSE IF    """${message}""" == "There is no Jobcard details found in the DMS report to proceed."
            Close DMS
            Close Erp
        ELSE IF    """${message}""" == "Unable to download Jobcard Carry Over Ratio Report."
            Close DMS
            Close Erp
            Send Email Output     ${curr_date}    ${report_name}    ${message}    ${EMPTY}
        ELSE IF    "${consolid_report_path_in_processed_erp_dir}" == "${EMPTY}"
            Send Email Output     ${curr_date}    ${report_name}    ${message}    ${EMPTY}
            Log    hi
        ELSE
            Log    hi
            Send Email Output     ${curr_date}    ${report_name}    ${message}    ${consolid_report_path_in_processed_erp_dir}     
        END
        Fail      
    END 