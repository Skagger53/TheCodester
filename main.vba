' This workbook and all VBA code was created by Matt Skaggs

Option Explicit
Global table_start_row As Integer
Global next_assign_row As Long
' Flag used in called subs for data validation and checked in calling sub
Global validation_failed As Boolean
' New semen/oocyte sheet name. Must be created in a called sub following earlier validation check but _
  then is used in calling sub (if validation check passed)
Global s_o_name As String
' New embryo sheet name. Must be created in a called sub following earlier validation check but _
  then is used in calling sub (if validation check passed)
Global e_name As String

Sub create_sheets()

On Error GoTo ErrorHandling
    
    ' Runs all validation checks.
    ' If a validation fails, it will inform the user and return here with validation_failed set to True
    validation_failed = False
    Call validate_data
    If validation_failed = True Then Exit Sub
    
    ' Obtaining confirmation from user
    Dim create_sheets_yn As VbMsgBoxResult
    
    create_sheets_yn = MsgBox( _
        "Do you want to set up the '" & Config.Range("sheet_name_1") & "' sheet and the '" & _
            Config.Range("sheet_name_2") & "' sheet for " & DATA_Accts.Range("month_name") & " " & _
            DATA_Accts.Range("year") & "?" & vbNewLine & vbNewLine & _
            "This will delete your undo history; you will not be able to undo this.", _
            vbYesNo + vbDefaultButton2, _
            "Set up monthly sheets?" _
            )
    If create_sheets_yn = vbNo Then Exit Sub
    
    ' Setting up new sheet names
    Call disable_updating
    Call unprotect_sheet(DATA_Accts, password_req:=False)
    
    ' Copying and renaming template
    Call unprotect_sheet(TEMPLATE, password_req:=False)
    
    Call copy_templ(e_name, Config.Range("sheet_name_2").Value)
    Call copy_templ(s_o_name, Config.Range("sheet_name_1").Value)
    
    ' Moving filtered data to relevant sheets
    Call move_data(DATA_Accts.Range("s_o_testrange"), DATA_Accts.Range("s_o_testrange").CurrentRegion, 1)
    Call move_data( _
        DATA_Accts.Range("e_testrange"), _
        DATA_Accts.Range("e_testrange").CurrentRegion, _
        2 _
        )
    
    ' If user has indicated they want to assign team members automatically, calls relevant sub
    If DATA_Accts.Range("assign_yn") = 1 Then Call assign_team_members
    
ErrorHandling:

    ' Protecting and hiding sheets and enabling Excel updating/calculations
    Call protect_sheet(TEMPLATE, hide:=True, password_req:=False)
    Call protect_sheet(DATA_Accts, hide:=True, password_req:=False)
    Dim ws As Worksheet
    
    ' Protects the newly created sheets
    ThisWorkbook.Sheets(1).Protect _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True
    ThisWorkbook.Sheets(2).Protect _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True
    
    Call enable_updating
    
    ' Ensuring user is back at OriginalData where they started
    OriginalData.Activate
    OriginalData.Range("A1").Select
    
    ' If no errors, inform user macro is complete
    If Err.Number = 0 Then
        MsgBox _
            "Sheets created.", _
            Title:="Macro complete"
    End If
    
    ' If any errors, informing user
    If Err.Number <> 0 Then
        MsgBox _
            "Error encountered: " & Err.Number & vbNewLine & vbNewLine & _
                "Check this workbook's sheets; setup may not have completed.", _
            Title:="Error encountered"
    End If

End Sub

Sub reset_import_table()

On Error GoTo ErrorHandling:

    ' Obtains user's confirmation
    Dim reset_table_yn As VbMsgBoxResult
    
    reset_table_yn = MsgBox( _
        "Do you want to reset the import table?" & vbNewLine & vbNewLine & _
        "This will remove ALL data from this table " & _
        "(though all data in the Semen/Oocyte sheets and Embryo sheets will remain)." & vbNewLine & vbNewLine & _
        "This will clear your undo history; you cannot undo this.", _
        vbYesNo + vbDefaultButton2, _
        "Reset import table" _
        )
    
    If reset_table_yn = vbNo Then Exit Sub
    
    ' Prepares workbook for editing
    Call disable_updating
    Call unprotect_sheet(OriginalData, password_req:=True)
    
    ' Gets start and end rows for data in Original Data sheet
    Dim table_data_range As Range
    Dim table_end_row As Integer
    
    Call assign_table_start_row(OriginalData.ListObjects("original_data")) ' Sets table_start_row
    table_end_row = OriginalData.Range("A" & OriginalData.Cells(OriginalData.Rows.Count, 1).Row).End(xlUp).Row
    
    Set table_data_range = OriginalData.Range("A" & table_start_row & ":D" & table_end_row)
    
    table_data_range.Value = ""
    
    OriginalData.ListObjects("original_data").Resize Range("$A$3:$D$13")
    
ErrorHandling:
    ' Prepares workbook for user access again
    Call protect_sheet(OriginalData, hide:=False, password_req:=True)
    Call enable_updating
    
    If Err.Number <> 0 Then ' Some error encountered
        MsgBox _
            "Error encountered: " & Err.Number & vbNewLine & vbNewLine & _
                "If table did not reset, clear the data out manually.", _
            Title:="Error encountered"
    End If
    
End Sub
