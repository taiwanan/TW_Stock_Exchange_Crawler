Attribute VB_Name = "Module1"
Sub UpdateTableValues()
Attribute UpdateTableValues.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UpdateTableValues Macro
'

    WorkingPath = ActiveWorkbook.Path
    'MsgBox ("Updating table values...")
    Application.DisplayStatusBar = True
    Dim mFile As String
       mFile = Dir(WorkingPath & "\tse_pressure_data\*.xlsx")
       Do While mFile <> ""
          Workbooks.Open Filename:=WorkingPath & "\tse_pressure_data\" & mFile
          Application.StatusBar = "Loading files:" & mFile & "..."
          mFile = Dir()
          ActiveWindow.Close (False)
       Loop
    Application.DisplayStatusBar = False

   
End Sub
