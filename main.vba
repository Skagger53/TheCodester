Option Explicit

' Disabling Excel screen updating, calculations
Sub disable_updating()

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .StatusBar = "Updating..."
    End With

End Sub

'Enabling Excel screen updating, calculations
Sub enable_updating()

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .StatusBar = ""
    End With

End Sub

' Checking to see if a sheet name already exists
Sub dup_name_msg(sheet_name As String)

    MsgBox _
        "The sheet name """ & sheet_name & """ already exists. " & _
            "Please generate a different month or delete the conflicting sheet(s)." & _
            vbNewLine & vbNewLine & _
            "Macro will end; no changes will be made.", _
        Title:="Duplicate sheet name"

End Sub

' Copying template sheet with designated sheet name and applying title in A1
Sub copy_templ(sheet_name As String, a1_title As String)
    
    TEMPLATE.Copy before:=ThisWorkbook.Sheets(1)
    ThisWorkbook.Sheets(1).Activate
    ActiveSheet.Name = sheet_name
    ActiveSheet.Range("A1").Value = a1_title

End Sub

Sub unprotect_sheet(code_name, password_req As Boolean)
    ' Unprotects sheet. password_req indicates if sheet is password-protected.
    
    code_name.Visible = xlSheetVisible
    code_name.Activate
    If password_req = True Then
        code_name.Unprotect password:="$z6jvJbm3vufn#NDLY"
    Else
        code_name.Unprotect
    End If

End Sub

Sub protect_sheet(code_name, hide As Boolean, password_req As Boolean)
    ' Protects a sheet.
    ' hide determines if sheet should be xlVeryHidden.
    ' password_req determines if sheet should be password-protected

    If password_req = True Then
        code_name.Protect _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            password:="$z6jvJbm3vufn#NDLY"
    Else
        code_name.Protect _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True
    End If
    If hide = True Then code_name.Visible = xlSheetVeryHidden
        
End Sub

Sub move_data(test_cell As Range, originating_data As Range, sheet_index As Byte)
    ' Moves data from DATA_

    If test_cell = "" Then Exit Sub ' If no filtered data was found
    
    ' The below is flexible and not hard-coded to allow for the TEMPLATE table to be changed (moved up or down rows)
    ' The below will fail if the TEMPLATE table were to originate in a column other than A
    
    ' Obtaining the first row of data in TEMPLATE (and thus newly created sheets) to begin placing data at
    ' table_start_row is global
    Call assign_table_start_row(TEMPLATE.ListObjects("monthly_template"))
    
    ' Gets row count of original data
    ' This must be added to the table start row to find the final row to copy data to (minus one)
    Dim originating_row_count As Integer
    originating_row_count = originating_data.Rows.Count
    
    ThisWorkbook.Sheets(sheet_index).Range( _
        "A" & table_start_row & ":" & _
        "D" & originating_row_count + table_start_row - 1 _
    ).Value = originating_data.Value
    
End Sub

' next_assign_row = row to start assigning to, assign_num = how many assignments to apply,
' team_memb_start_row = row to start on for team members' names (in DATA_Accts!W), and
' team_memb_end_row = row to end on for team members' names (in DATA_Accts!W).
Sub write_assignments(sheet_index, next_assign_row, assign_num, team_memb_start_row, team_memb_end_row)

    Dim i As Integer
    For i = team_memb_start_row To team_memb_end_row
        ' Starts at the next assignment row and continues on for however many assignments are passed in
        ' for this sub minus 1 (due to inclusive start range assignment)
        ' Assigns the next name from the team members list
        ThisWorkbook.Sheets(sheet_index).Range( _
            "F" & next_assign_row & ":F" & next_assign_row + assign_num - 1 _
            ).Value _
            = _
            DATA_Accts.Range("W" & i).Value
        ' Next row to assign is the previous starting row plus however many assignments are being applied
        next_assign_row = next_assign_row + assign_num
    Next i

End Sub

Sub assign_table_start_row(table_name As Object)
    
    ' Gets the range of the template's table, finds the start row for it, and adds 1 (skip over headings)
    table_start_row = _
        Replace( _
            Split( _
                table_name.Range.Address, "$" _
                )(2), ":", "" _
            ) + 1

End Sub

Sub validate_data()
    ' Conducts all data valition required.
    ' If any fails, notifies user, sets validation_failed = True, and exits sub

    ' Ensuring month and year are selected
        If DATA_Accts.Range("month_num").Value = "" Or DATA_Accts.Range("year").Value = "" Then
            MsgBox _
                "Please enter a month and year in the Config sheet before creating monthly sheets.", _
                Title:="Month and year required"
            validation_failed = True
            Exit Sub
        End If
    
    ' Ensuring names have been entered for both new sheets
        If Config.Range("sheet_name_1") = "" Or Config.Range("sheet_name_2") = "" Then
            MsgBox _
                "Please enter a name for both sheets (semen/oocyte and embryo) in the Config sheet.", _
                Title:="Sheet names required"
            validation_failed = True
            Exit Sub
        End If
    
    ' Ensuring data does not extend beyond the OriginalData table
        ' Obtaining last row in table
        Dim original_data_table_last_row As Long
        
        original_data_table_last_row = Split(OriginalData.ListObjects("original_data").Range.Address, "$")(4)
        
        ' Checking to see if data exists below the table in columns A:D
        Dim i As Long
        For i = 65 To 68
            Call data_below_table(Chr(i), original_data_table_last_row)
            If validation_failed = True Then Exit Sub
        Next i
        
    ' Checking that some data is present in column A
        Call assign_table_start_row(OriginalData.ListObjects("original_data")) ' Assigns table_start_row
        
        Dim rows_have_data As Boolean
        rows_have_data = False
        
        ' Iterates through all table rows to see if data is present in column A
        For i = table_start_row To original_data_table_last_row
            If OriginalData.Range("A" & i) <> "" Then
                rows_have_data = True
                Exit For
            End If
        Next i
        
        If rows_have_data = False Then ' No row was ever found with data
            MsgBox _
                "The Account Number column in the Original Data Import table appears to be empty. " & _
                    "Please enter data into the table before generating monthly sheets.", _
                Title:="Missing data"
            validation_failed = True
            Exit Sub
        End If
    
    ' Ensuring planned sheet names do not already exist
        ' Planning sheet names
        e_name = Format(DATA_Accts.Range("month_num"), "00") & "." & DATA_Accts.Range("year") & " E"
        s_o_name = Format(DATA_Accts.Range("month_num"), "00") & "." & DATA_Accts.Range("year")
        
        ' Checking if either sheet name already exists
        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Sheets
            If sh.Name = e_name Then
                Call dup_name_msg(e_name)
                validation_failed = True
                Exit Sub
            ElseIf sh.Name = s_o_name Then
                Call dup_name_msg(s_o_name)
                validation_failed = True
                Exit Sub
            End If
        Next sh

End Sub

Sub data_below_table(col_to_test, original_data_table_last_row)
    ' Obtaining last row of data in column passed in
    Dim original_data_last_row As Long
    
    original_data_last_row = OriginalData.Cells(OriginalData.Cells.Rows.Count, 1).Row
    original_data_last_row = OriginalData.Range(col_to_test & original_data_last_row).End(xlUp).Row
    
    ' Checking that last row of data is out beyond the last row of the table
    If original_data_last_row > original_data_table_last_row Then
        MsgBox _
            "Data in the Original Data sheet extends beyond the table. " _
                & vbNewLine & vbNewLine & "Check for data below the table. " & _
                "Either extend the table to include this data or " & _
                "remove the data below the table." & vbNewLine & vbNewLine & _
                "Macro will not execute; no action will be taken on this data.", _
            Title:="Data outside table bounds"
            validation_failed = True
            Exit Sub
    End If

End Sub

Sub assign_team_members()

    If DATA_Accts.Range("team_members_filtered") = "" Then Exit Sub
    
    ' Assigns how many team members there are and how many assignments (rows) there are for each sheet
    Dim num_team_members As Byte, num_s_o_rows As Long, num_e_rows As Long
    num_team_members = DATA_Accts.Range("team_members_filtered").CurrentRegion.Rows.Count
    num_s_o_rows = DATA_Accts.Range("s_o_testrange").CurrentRegion.Rows.Count
    num_e_rows = DATA_Accts.Range("e_testrange").CurrentRegion.Rows.Count
    
    ' Assigns how many team members will be at the higher number of assignments for each sheet and
    ' how many at the lower number (will only be a difference of one in how many they're assigned)
    Dim s_o_team_membs_high_assign As Integer ' How many semem/oocyte team members at the high assignment number
    Dim s_o_team_membs_low_assign As Integer ' How many semem/oocyte team members at the low assignment number
    Dim s_o_high_assign_num As Integer ' The high assignment number for semem/oocyte
    Dim s_o_low_assign_num As Integer ' The low assignment number for semem/oocyte
    Dim e_team_membs_high_assign As Integer ' How many embryo team members at the high assignment number
    Dim e_team_membs_low_assign As Integer ' How many embryo team members at the low assignment number
    Dim e_high_assign_num As Integer ' The high assignment number for embryo
    Dim e_low_assign_num As Integer ' The low assignment number for embro
    
    ' Lower assignment number is how many team members divided by possible assignments truncated
    s_o_low_assign_num = Int(num_s_o_rows / num_team_members)
    ' Higher assignment number is one above the lower
    s_o_high_assign_num = s_o_low_assign_num + 1
    ' How many team members at the higher assignment number is how many assignments modulus how many team members
    s_o_team_membs_high_assign = num_s_o_rows Mod num_team_members
    ' How many team members at the lower assignment number is however many team members are left
    s_o_team_membs_low_assign = num_team_members - s_o_team_membs_high_assign
    ' Lower assignment number is how many team members divided by possible assignments truncated
    e_low_assign_num = Int(num_e_rows / num_team_members)
    ' Higher assignment number is one above the lower
    e_high_assign_num = e_low_assign_num + 1
    ' How many team members at the higher assignment number is how many assignments modulus how many team members
    e_team_membs_high_assign = num_e_rows Mod num_team_members
    ' How many team members at the lower assignment number is however many team members are left
    e_team_membs_low_assign = num_team_members - e_team_membs_high_assign
    
    next_assign_row = table_start_row ' The "next" row to assign is the very first available row
    
    ' Writes the assignments to Sheet(1) (semen and oocyte) and Sheet(2) (embryo)
    ' Arguments are:
        ' sheet index,
        ' the next row to start on,
        ' higher assignment number,
        ' what row to start on for team members (in DATA_Accts!W),
        ' what row to end on for team members (in DATA_Accts!W).
    Call write_assignments(1, next_assign_row, s_o_high_assign_num, 1, s_o_team_membs_high_assign)
    Call write_assignments(1, next_assign_row, s_o_low_assign_num, s_o_team_membs_high_assign + 1, num_team_members)
        
    ' Writing to Sheet(2) (embryo)
    next_assign_row = table_start_row
    Call write_assignments(2, next_assign_row, e_high_assign_num, 1, e_team_membs_high_assign)
    Call write_assignments(2, next_assign_row, e_low_assign_num, e_team_membs_high_assign + 1, num_team_members)

End Sub

