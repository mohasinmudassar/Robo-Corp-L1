*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log in
    Download the Excel File
    Fill the form using the data from the Excel File
    Collect the results
    Export the table as a PDF
    Log out and Close the Browser


*** Keywords ***
Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/

Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the Excel File
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

#Fill and Submit the Form
#    Input Text    firstname    Mohasin
#    Input Text    lastname    Mudassar
#    Input Text    salesresult    500
#    Select From List By Value    salestarget    5000
#    Click Button    Submit

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Fill the form using the data from the Excel File
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet as Table    header=True
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export the table as a PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Log out and Close the Browser
    Click Button    Log out
    Close Browser
