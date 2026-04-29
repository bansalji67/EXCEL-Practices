'input excel file's full path
 ExcelFilePath="G:\Ansari\Practice of me on Excel\How to Open Exel even system locked\Open Excel File if System Lock.xlsx"

'input Module/Macro Name withing the excel file
 'MacroPath="extract_email_atmts.Extract"

'Create an instance of Excel
 Set ExcelApp=CreateObject("Excel.Application")

'Do you want this Excel instance to be visible?
 ExcelApp.Visible=True 'or "False"

'Prevennt any App Launch Alerts (ie Update External Links)
 ExcelApp.DisplayAlerts=False

'Open Excel File
 Set wb=ExcelApp.workbooks.Open(ExcelFilePath)

'Execute Macro Code
 'ExcelApp.Run MacroPath

'Save Excel File (if applicable)
 'wb.Save

'Reset Display Alerts Before Closing
' ExcelApp.DisplayAlerts=True

'Close Excel File
' wb.close

'End  Instance of Excel
 ExcelApp.Quit


'Leave an onscreen message
 msgbox  "Your Automated Task successfully"


