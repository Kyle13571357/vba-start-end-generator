'==============================================================================
' Module: Transportation
' Author: Kyle Hsu
' Date: April 21, 2025
' Description: Functions to automate transportation allowance form completion
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Function: GenerateStartEndLocations
' Purpose: Automatically generates start and end locations for transportation 
'          reimbursement based on employee home address and schedule
' Parameters: None
' Returns: None
'------------------------------------------------------------------------------
Sub GenerateStartEndLocations()
    ' Variables declaration
    Dim iEmployee As Integer    ' Loop counter for employee rows
    Dim iSheet As Integer       ' Loop counter for worksheet index
    Dim strHomeLocation As String ' Employee's home location
    Dim objT1LookupRange As Range ' Range for T1 route lookup 
    Dim objT2LookupRange As Range ' Range for T2 route lookup
    Dim objCellValue As Range   ' Cell value for lookup
    Dim blnUserConfirmed As Boolean ' User confirmation flag
    
    On Error GoTo ErrorHandler
    
    ' Prompt user for confirmation
    blnUserConfirmed = (MsgBox("Generate start and end locations for all employees?", _
                             vbQuestion + vbYesNo, "Location Generator") = vbYes)
    
    If Not blnUserConfirmed Then
        Exit Sub
    End If
    
    ' Initialize lookup ranges from reference worksheet
    Set objT1LookupRange = Sheets("WorkSheet1").Range("A3:E100")
    Set objT2LookupRange = Sheets("WorkSheet1").Range("I3:M100")
    
    Application.ScreenUpdating = False
    
    ' Process Terminal 1 Sheets (sheets 6-49)
    For iSheet = 6 To 49
        Sheets(iSheet).Select
        
        ' Process each row in current sheet
        For iEmployee = 4 To 26
            Set objCellValue = ActiveSheet.Range("C" & iEmployee)
            
            ' Lookup start/end location from reference data
            If Not IsError(Application.WorksheetFunction.VLookup(objCellValue, objT1LookupRange, 5, False)) Then
                Range("F" & iEmployee) = Application.WorksheetFunction.VLookup(objCellValue, objT1LookupRange, 5, False)
            Else
                Range("F" & iEmployee) = ""
            End If
            
            ' Add home location to route direction if available
            If Range("L13").Value <> 0 Then
                If Range("F" & iEmployee).Value = " →Terminal1" Then
                    Range("F" & iEmployee) = Range("L13") & "→Terminal1"
                ElseIf Range("F" & iEmployee).Value = "Terminal1→ " Then
                    Range("F" & iEmployee) = "Terminal1→" & Range("L13")
                End If
            End If
        Next iEmployee
    Next iSheet
    
    ' Process Terminal 2 Sheets (sheets 50-95)
    For iSheet = 50 To 95
        Sheets(iSheet).Select
        
        ' Process each row in current sheet
        For iEmployee = 4 To 26
            Set objCellValue = ActiveSheet.Range("C" & iEmployee)
            
            ' Lookup start/end location from reference data
            If Not IsError(Application.WorksheetFunction.VLookup(objCellValue, objT2LookupRange, 5, False)) Then
                Range("F" & iEmployee) = Application.WorksheetFunction.VLookup(objCellValue, objT2LookupRange, 5, False)
            Else
                Range("F" & iEmployee) = ""
            End If
            
            ' Add home location to route direction if available
            If Range("L13").Value <> 0 Then
                If Range("F" & iEmployee).Value = " →Terminal2" Then
                    Range("F" & iEmployee) = Range("L13") & "→Terminal2"
                ElseIf Range("C" & iEmployee).Value = "C" Then
                    Range("F" & iEmployee) = "Terminal2Ext→T1→" & Range("L13")
                ElseIf Range("F" & iEmployee).Value = "Terminal2→ " Then
                    Range("F" & iEmployee) = "Terminal2→" & Range("L13")
                End If
            End If
        Next iEmployee
    Next iSheet
    
    Application.ScreenUpdating = True
    MsgBox "Location generation completed successfully!", vbInformation, "Process Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume ExitSub
    
ExitSub:
    ' Clean up
    Set objT1LookupRange = Nothing
    Set objT2LookupRange = Nothing
    Set objCellValue = Nothing
End Sub
