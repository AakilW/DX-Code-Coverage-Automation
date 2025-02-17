VBA for CheckDXCoverage 

Sub CheckDXCoverage()
    Dim wsMaster As Worksheet
    Dim wsCPT As Worksheet
    Dim cptCode As String
    Dim dxCode As String
    Dim dxCodeArray As Variant
    Dim foundMatch As Boolean
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long, k As Long, l As Long
    Dim dxLastRow As Long
    Dim dxLastCol As Long
    Dim dxCellValue As String

    ' Set the Master Tracker sheet
    Set wsMaster = ThisWorkbook.Sheets("Master Tracker")

    ' Get the last row of data in the Master Tracker
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in the Master Tracker
    For i = 2 To lastRow ' Assuming row 1 has headers
        cptCode = Trim(UCase(wsMaster.Cells(i, "A").Value)) ' CPT Code
        
        ' Check if DX column is not empty
        If Trim(wsMaster.Cells(i, "B").Value) <> "" Then
            ' Split multiple DX codes (comma-separated)
            dxCodeArray = Split(Trim(UCase(wsMaster.Cells(i, "B").Value)), ",")
        Else
            dxCodeArray = Array() ' Empty array
        End If
        
        foundMatch = False ' Reset flag for each row
        
        ' Check if the CPT sheet exists
        If SheetExists(cptCode) Then
            ' If the CPT sheet exists, assign it to wsCPT
            Set wsCPT = ThisWorkbook.Sheets(cptCode)
            wsMaster.Cells(i, "C").Value = "Sheet Exists"
            
            ' Get the last row and last column in the CPT sheet
            dxLastRow = wsCPT.Cells(Rows.Count, 1).End(xlUp).Row
            dxLastCol = wsCPT.Cells(1, Columns.Count).End(xlToLeft).Column

            ' Loop through each DX code in the Master Tracker
            For j = LBound(dxCodeArray) To UBound(dxCodeArray)
                dxCode = Trim(UCase(dxCodeArray(j))) ' Clean individual DX code
                
                ' Loop through the entire CPT sheet (All Rows & Columns)
                For k = 1 To dxLastRow ' Rows
                    For l = 1 To dxLastCol ' Columns
                        dxCellValue = Trim(UCase(Application.Trim(wsCPT.Cells(k, l).Value))) ' Get cell value

                        ' Handle merged cells (Extract from the first merged cell)
                        If wsCPT.Cells(k, l).MergeCells Then
                            dxCellValue = Trim(UCase(Application.Trim(wsCPT.Cells(k, l).MergeArea.Cells(1, 1).Value)))
                        End If

                        ' Compare values
                        If dxCellValue = dxCode Then
                            foundMatch = True
                            Exit For
                        End If
                    Next l
                    If foundMatch Then Exit For
                Next k
                
                ' Exit loop if any DX code is found
                If foundMatch Then Exit For
            Next j
            
            ' Output DX coverage result in column D
            If foundMatch Then
                wsMaster.Cells(i, "D").Value = "Covered"
            Else
                wsMaster.Cells(i, "D").Value = "Uncovered"
            End If
            
        Else
            ' Output if sheet doesn't exist in column C
            wsMaster.Cells(i, "C").Value = "Sheet Does Not Exist"
            
            ' Output "Check the AAPC" in column D if the sheet does not exist
            wsMaster.Cells(i, "D").Value = "Check the AAPC"
        End If
    Next i
End Sub

' Function to check if a sheet exists
Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
