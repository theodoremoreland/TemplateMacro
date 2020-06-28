VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SR_Template()

    'A variable that stores the number of columns that aren't empty (only counts on the current/"active" sheet).
    Dim last_col As Long: last_col = WorksheetFunction.CountA(ActiveSheet.Rows(1))
    
    'An array with each element being a letter in the alphabet (arranged/added in alphabetical order).
    Dim al(1 To 26) As String
    
    al(1) = "A"
    al(2) = "B"
    al(3) = "C"
    al(4) = "D"
    al(5) = "E"
    al(6) = "F"
    al(7) = "G"
    al(8) = "H"
    al(9) = "I"
    al(10) = "J"
    al(11) = "K"
    al(12) = "L"
    al(13) = "M"
    al(14) = "N"
    al(15) = "O"
    al(16) = "P"
    al(17) = "Q"
    al(18) = "R"
    al(19) = "S"
    al(20) = "T"
    al(21) = "U"
    al(22) = "V"
    al(23) = "W"
    al(24) = "X"
    al(25) = "Y"
    al(26) = "Z"
    
    
    'Loops through every Sheet & creates a varibale with the value "True" if the Sheet's name is equal to the given name.
    For i = 1 To Worksheets.Count
    
        If Worksheets(i).Name = "apollo info" Then
            apollo_info_exists = True
        End If
        
        If Worksheets(i).Name = "weight file STP" Then
            weight_file_STP_exists = True
        End If
        
        If Worksheets(i).Name = "seed qty update" Then
            seed_qty_update_exists = True
        End If
        
    Next i
    
    'If no Sheet has the same name as one of the three provided above, create a Sheet and give it one of the above names.
    If Not apollo_info_exists Then
        Sheets.Add
        ActiveSheet.Name = "apollo info"
    End If
    
    If Not weight_file_STP_exists Then
        Sheets.Add
        ActiveSheet.Name = "weight file STP"
    End If
    
    If Not seed_qty_update_exists Then
        Sheets.Add
        ActiveSheet.Name = "seed qty update"
    End If


    '---------------------------------------------------------------------------------------------------------------------------'


    'Select "weight file STP" Sheet. Code following this statement will be applied to this Sheet only, until selecting another Sheet.
    Sheets("weight file STP").Activate
    
    'Replaces old "last_col" variable with the count of columns in the newly activated sheet.
    last_col = WorksheetFunction.CountA(ActiveSheet.Rows(1))
    
    'Loops through every column header & creates a varibale with the value "True" if the header's name is equal to the given name.
    For i = 1 To last_col

        If Cells(1, i).Value = "SU Name" Then
            SU_Name_exists = True
        End If

        If Cells(1, i).Value = "Weight Needed" Then
            Weight_Needed_exists = True
        End If

        If Cells(1, i).Value = "#SRT1" Then
            SRT1_exists = True
        End If

    Next i

    'If no header has the same name as one of the three provided above, create a column and give it one of the above headers.
    If Not SU_Name_exists Then
        Range("B1").EntireColumn.Insert
        Range("B1").Value = "SU Name"
    End If

    If Not Weight_Needed_exists Then
        Range("D1").EntireColumn.Insert
        Range("D1").Value = "Weight Needed"
    End If

    If Not SRT1_exists Then
        Range("E1").EntireColumn.Insert
        Range("E1").Value = "#SRT1"
    End If
    
    
    'Creates a variable that stores the location of headers with the given name (as a one digit integer).
    Dim su_name As Integer: su_name = Application.WorksheetFunction.Match("SU Name", ActiveSheet.Rows(1), 0)
    Dim weight_needed As Integer: weight_needed = Application.WorksheetFunction.Match("Weight Needed", ActiveSheet.Rows(1), 0)
    Dim srt1 As Integer: srt1 = Application.WorksheetFunction.Match("#SRT1", ActiveSheet.Rows(1), 0)

    'Does the same as above, but because these columns must be added manually, there is a possibility that they are missing.
    'Because an error might occur, the following statement will allow the macro to ignore the error.
    On Error Resume Next
        Dim receiver_inv_bid As Integer: receiver_inv_bid = Application.WorksheetFunction.Match("Receiver Inv BID", ActiveSheet.Rows(1), 0)
        Dim sender_inv_bid As Integer: sender_inv_bid = Application.WorksheetFunction.Match("Sender Inv BID", ActiveSheet.Rows(1), 0)
        Dim seed_quantity_quantity As Integer: seed_quantity_quantity = Application.WorksheetFunction.Match("Sent Quantity Quantity", ActiveSheet.Rows(1), 0)
        
    'If there is an error code other than "0" (i.e. not zero errors), then show a message and stop the code.
    If Err <> 0 Then
        MsgBox "One of the following headers was not found on sheet 'weight file STP'" & vbNewLine & "'Receiver Inv BID', 'Sender Inv BID', and/or 'Sent Quantity Quantity'." & vbNewLine & "Resolve missing headers and restart macro."
        Exit Sub
    End If
    
    'Creates variable for a letter in the "al" array using the position of the given headers.
    sbid_letter_location = al(sender_inv_bid)
    sqq_letter_location = al(seed_quantity_quantity)
    srt1_letter_location = al(srt1)
    wn_letter_location = al(weight_needed)
    
    'Loops through every non-empty row then concats the letter from the above variables with the number of current row (starting @ 2).
    'After creating variables for the concatenation, add formulas to each Cell as the loops continues.
    For i = 2 To WorksheetFunction.CountA(ActiveSheet.Columns(1))
        
        row_num = CStr(i)

        sbid_rng = sbid_letter_location & row_num
        sqq_rng = sqq_letter_location & row_num
        srt1_rng = srt1_letter_location & row_num
    
        'Becuase formulas are passed as a string, creates a string concatenating all necessary variables.
        Cells(i, su_name).Formula = "=VLOOKUP(" & sbid_rng & ",'apollo info'!A:G,7,0)"
        Cells(i, weight_needed).Formula = "=(" & sqq_rng & "/" & srt1_rng & ")*1.05"
        Cells(i, srt1).Formula = "=VLOOKUP(" & sbid_rng & ",'apollo info'!A:F,6,0)"
    Next i
    
    
    '---------------------------------------------------------------------------------------------------------------------------'
    
    
    'Select "seed qty update" Sheet.
    Sheets("seed qty update").Activate
    
    'Replaces old "last_col" variable with the count of columns in the newly activated sheet.
    last_col = WorksheetFunction.CountA(ActiveSheet.Rows(1))
    
    'Loops through every column header & creates a varibale with the value "True" if the header's name is equal to the given name.
    For i = 1 To last_col

        If Cells(1, i).Value = "Old Seed Qty" Then
            Old_Seed_Qty_exists = True
        End If

        If Cells(1, i).Value = "Seed Used" Then
            Seed_Used_exists = True
        End If

        If Cells(1, i).Value = "Seed Quantity" Then
            Seed_Quantity_exists = True
        End If

    Next i

    'If no header has the same name as one of the three provided above, create a column and give it one of the above headers.
    If Not Old_Seed_Qty_exists Then
        Range("B1").EntireColumn.Insert
        Range("B1").Value = "Old Seed Qty"
    End If

    If Not Seed_Used_exists Then
        Range("C1").EntireColumn.Insert
        Range("C1").Value = "Seed Used"
    End If

    If Not Seed_Quantity_exists Then
        Range("D1").EntireColumn.Insert
        Range("D1").Value = "Seed Quantity"
    End If
'
    'Creates a variable that stores the location of headers with the given name
    Dim old_seed_qty As Integer: old_seed_qty = Application.WorksheetFunction.Match("Old Seed Qty", ActiveSheet.Rows(1), 0)
    Dim seed_used As Integer: seed_used = Application.WorksheetFunction.Match("Seed Used", ActiveSheet.Rows(1), 0)
    Dim seed_quantity As Integer: seed_quantity = Application.WorksheetFunction.Match("Seed Quantity", ActiveSheet.Rows(1), 0)
    
    
    'Does the same as above, but because these columns must be added manually, there is a possibility that they are missing.
    'Because an error might occur, the following statement will allow the macro to ignore the error.
    On Error Resume Next
        Dim inv_bid As Integer: inv_bid = Application.WorksheetFunction.Match("Inv Bid", ActiveSheet.Rows(1), 0)
        Dim seed_quanitity_uom As Integer: seed_quanitity_uom = Application.WorksheetFunction.Match("Seed Quantity_UOM", ActiveSheet.Rows(1), 0)
    
    
    'If there is an error code other than "0" (i.e. not zero errors), then show a message and stop the code.
    If Err <> 0 Then
        MsgBox "One of the following headers was not found on sheet 'seed qty update'" & vbNewLine & "'Inv BID' and/or 'Seed Quantity_UOM'" & vbNewLine & "Resolve missing headers and restart macro."
        Exit Sub
    End If

    'Creates variable for a letter in the "al" array using the position of the given headers.
    bid_letter_location = al(inv_bid)
    old_seed_qty_letter_location = al(old_seed_qty)
    seed_used_letter_location = al(seed_used)
    
    
    'Loops through every non-empty row then concats the letter from the above variables with the number of current row (starting @ 2).
    'After creating variables, add formulas to each Cell as the loops continues.
    For i = 2 To WorksheetFunction.CountA(ActiveSheet.Columns(1))
    
        row_num = CStr(i)

        bid_rng = bid_letter_location & row_num
        old_seed_qty_rng = old_seed_qty_letter_location & row_num
        seed_used_rng = seed_used_letter_location & row_num
        
        'Becuase formulas are passed as a string, creates a string concatenating all necessary variables.
        Cells(i, old_seed_qty).Formula = "=VLOOKUP(" & bid_rng & ",'apollo info'!A:D,4,0)"
        Cells(i, seed_quantity).Formula = "=" & old_seed_qty_rng & "-" & seed_used_rng
        
        'The below string uses the equation "(weight_needed - sender_inv_bid + 1)" or "(y - x + 1)" to ensure the vlookup is returning the right column value.
        Cells(i, seed_used).Formula = "=VLOOKUP(" & bid_rng & ",'weight file STP'!" & sbid_letter_location & ":" & wn_letter_location & "," & CStr(weight_needed - sender_inv_bid + 1) & ",0)*VLOOKUP(" & bid_rng & ",'weight file STP'!" & sbid_letter_location & ":" & srt1_letter_location & "," & CStr(srt1 - sender_inv_bid + 1) & ",0)"
            
    Next i
    


End Sub


