# Excel-Automated-Invoice
Excel VBA Automated Invoice Template

## Overview
Uses Visual Basics and Excel to automatically copy invoice information from csv (Comma-separated values) format to a customized template.

## Instructions
1)  Copy csv information into Sheet1 under header.
2)  Run Script Clean_Data_Up if necessary to format csv data before invoice creation.
3)  Run Script Automated_Invoice to automatically copy formatted csv data into Sheet2 invoice templet.

## Code Clean_Data_Up

Sub Clean_Data_Up()
'Define Variables
Dim Lastrow As Long
Dim Str As String
Dim RemovePO As String

'Define Last Row with Data in Worksheet
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Remove PO from Order Number and Format
For i = 2 To Lastrow
 Str = Worksheets("Sheet1").Range("B" & i)
 RemovePO = Replace(Str, "PO ", "")
 Worksheets("Sheet1").Range("B" & i) = RemovePO
 Worksheets("Sheet1").Range("B" & i).NumberFormat = "000000;@"
 Next i

'Set Style for Range
With Worksheets("Sheet1").Range("A2:O2" & Lastrow)
    .Font.Name = "Tahoma"
    .Font.Size = 8
    .HorizontalAlignment = xlHAlignCenter
End With
 
'Format Date and Remove Time Stamp
 Worksheets("Sheet1").Range("A:A").NumberFormat = "mm-dd-yyyy;@"

'Loop Remove Zip Code +4
 For Each cell In Range("I2:I" & Lastrow)
    cell.Value = Left(cell.Value, 5)
 Next
 
'Format Currency
 Worksheets("Sheet1").Range("L:M").NumberFormat = "$#,##0.00_);($#,##0.00)"

End Sub

## Code Automated_Invoice

Sub Automated_Invoice()
'Define Variables
Dim FileName As String
Dim Lastrow As Long
Dim Arr(0 To 2) As String
Dim OriginalText As String
Dim CorrectedText As String
Dim podate As Range
Dim ponum As Range
Dim poname As Range
Dim pocompany As Range
Dim poaddress1 As Range
Dim poaddress2 As Range
Dim poqty As Range
Dim posku As Range
Dim poprice As Range
Dim pototal As Range
Dim poinvnum As Range
Dim multi As Interior
Dim x As Interior
Lastrow = Worksheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
OriginalText = Worksheets("Sheet1").Range("A10").Value

'Loop Copy/Paste/Print`
For i = 2 To Lastrow
Set poinvnum = Worksheets("Sheet1").Range("N" & i)
    'PO Number duplicate check
    If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 1).Value Then
        'Item Qty
        Set poqty = Worksheets("Sheet1").Range("J" & i + 1)
            Worksheets("Sheet2").Range("A24") = poqty
        'Item Sku
        Set posku = Worksheets("Sheet1").Range("K" & i + 1)
            Worksheets("Sheet2").Range("B24") = posku
        'Item Price
        Set poprice = Worksheets("Sheet1").Range("L" & i + 1)
            Worksheets("Sheet2").Range("H24") = poprice
        'Order Total
        Set pototal = Worksheets("Sheet1").Range("M" & i + 1)
            Worksheets("Sheet2").Range("I24") = pototal
        If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 2) Then
            'Item Qty
            Set poqty = Worksheets("Sheet1").Range("J" & i + 2)
                Worksheets("Sheet2").Range("A25") = poqty
            'Item Sku
            Set posku = Worksheets("Sheet1").Range("K" & i + 2)
                Worksheets("Sheet2").Range("B25") = posku
            'Item Price
            Set poprice = Worksheets("Sheet1").Range("L" & i + 2)
                Worksheets("Sheet2").Range("H25") = poprice
            'Order Total
            Set pototal = Worksheets("Sheet1").Range("M" & i + 2)
                Worksheets("Sheet2").Range("I25") = pototal
            If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 3) Then
                'Item Qty
                Set poqty = Worksheets("Sheet1").Range("J" & i + 3)
                    Worksheets("Sheet2").Range("A26") = poqty
                'Item Sku
                Set posku = Worksheets("Sheet1").Range("K" & i + 3)
                    Worksheets("Sheet2").Range("B26") = posku
                'Item Price
                Set poprice = Worksheets("Sheet1").Range("L" & i + 3)
                    Worksheets("Sheet2").Range("H26") = poprice
                'Order Total
                Set pototal = Worksheets("Sheet1").Range("M" & i + 3)
                    Worksheets("Sheet2").Range("I26") = pototal
                If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 4) Then
                    'Item Qty
                    Set poqty = Worksheets("Sheet1").Range("J" & i + 4)
                        Worksheets("Sheet2").Range("A27") = poqty
                    'Item Sku
                    Set posku = Worksheets("Sheet1").Range("K" & i + 4)
                        Worksheets("Sheet2").Range("B27") = posku
                    'Item Price
                    Set poprice = Worksheets("Sheet1").Range("L" & i + 4)
                        Worksheets("Sheet2").Range("H27") = poprice
                    'Order Total
                    Set pototal = Worksheets("Sheet1").Range("M" & i + 4)
                        Worksheets("Sheet2").Range("I27") = pototal
                    If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 5) Then
                        'Item Qty
                        Set poqty = Worksheets("Sheet1").Range("J" & i + 5)
                            Worksheets("Sheet2").Range("A28") = poqty
                        'Item Sku
                        Set posku = Worksheets("Sheet1").Range("K" & i + 5)
                            Worksheets("Sheet2").Range("B28") = posku
                        'Item Price
                        Set poprice = Worksheets("Sheet1").Range("L" & i + 5)
                            Worksheets("Sheet2").Range("H28") = poprice
                        'Order Total
                        Set pototal = Worksheets("Sheet1").Range("M" & i + 5)
                            Worksheets("Sheet2").Range("I28") = pototal
                        If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 6) Then
                            'Item Qty
                            Set poqty = Worksheets("Sheet1").Range("J" & i + 6)
                                Worksheets("Sheet2").Range("A29") = poqty
                            'Item Sku
                            Set posku = Worksheets("Sheet1").Range("K" & i + 6)
                                Worksheets("Sheet2").Range("B29") = posku
                            'Item Price
                            Set poprice = Worksheets("Sheet1").Range("L" & i + 6)
                                Worksheets("Sheet2").Range("H29") = poprice
                            'Order Total
                            Set pototal = Worksheets("Sheet1").Range("M" & i + 6)
                                Worksheets("Sheet2").Range("I29") = pototal
                            If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 7) Then
                                'Item Qty
                                Set poqty = Worksheets("Sheet1").Range("J" & i + 7)
                                    Worksheets("Sheet2").Range("A30") = poqty
                                'Item Sku
                                Set posku = Worksheets("Sheet1").Range("K" & i + 7)
                                    Worksheets("Sheet2").Range("B30") = posku
                                'Item Price
                                Set poprice = Worksheets("Sheet1").Range("L" & i + 7)
                                    Worksheets("Sheet2").Range("H30") = poprice
                                'Order Total
                                Set pototal = Worksheets("Sheet1").Range("M" & i + 7)
                                    Worksheets("Sheet2").Range("I30") = pototal
                                If poinvnum.Value = Worksheets("Sheet1").Range("N" & i + 8) Then
                                Debug.Print poinvnum.Value
                                Msg = "Too many Lines items, check debug window."
                                Else
                                    'Date
                                    Set podate = Worksheets("Sheet1").Range("A" & i)
                                        Worksheets("Sheet2").Range("G13:G14") = podate
                                    'PO Number
                                    Set ponum = Worksheets("Sheet1").Range("B" & i)
                                        Worksheets("Sheet2").Range("G15:G18") = ponum
                                    'Name
                                    Set poname = Worksheets("Sheet1").Range("C" & i)
                                        Worksheets("Sheet2").Range("B15") = poname
                                    'Company
                                    Set pocompany = Worksheets("Sheet1").Range("D" & i)
                                        Worksheets("Sheet2").Range("B16") = pocompany
                                    'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                                    Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
                                        If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                                            Worksheets("Sheet2").Range("B16") = poaddress1
                                            Else
                                            Worksheets("Sheet2").Range("B17") = poaddress1
                                        End If
                                    'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                                    Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
                                        If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                                            Worksheets("Sheet2").Range("B17") = poaddress2
                                            Else
                                            Worksheets("Sheet2").Range("B18") = poaddress2
                                        End If
                                    'Define Array for Combining City, State, Zip
                                    Arr(0) = Worksheets("Sheet1").Range("G" & i)
                                    Arr(1) = Worksheets("Sheet1").Range("H" & i)
                                    Arr(2) = Worksheets("Sheet1").Range("I" & i)
                                    OriginalText = Join(Arr)
                                    'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                                    If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                                        Worksheets("Sheet2").Range("B17").Value = Join(Arr)
                                        'Add Comma after City Only
                                        CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                        Worksheets("Sheet2").Range("B17").Value = CorrectedText
                                        ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
                                        Worksheets("Sheet2").Range("B18").Value = Join(Arr)
                                        'Add Comma after City Only
                                        CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                        Worksheets("Sheet2").Range("B18").Value = CorrectedText
                                        Else
                                        Worksheets("Sheet2").Range("B19").Value = Join(Arr)
                                        'Add Comma after City Only
                                        CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                        Worksheets("Sheet2").Range("B19").Value = CorrectedText
                                    End If
                                    'Item Qty
                                    Set poqty = Worksheets("Sheet1").Range("J" & i)
                                        Worksheets("Sheet2").Range("A23") = poqty
                                    'Item Sku
                                    Set posku = Worksheets("Sheet1").Range("K" & i)
                                        Worksheets("Sheet2").Range("B23") = posku
                                    'Item Price
                                    Set poprice = Worksheets("Sheet1").Range("L" & i)
                                        Worksheets("Sheet2").Range("H23") = poprice
                                    'Order Total
                                    Set pototal = Worksheets("Sheet1").Range("M" & i)
                                        Worksheets("Sheet2").Range("I23") = pototal
                                    'Invoice Number
                                        Worksheets("Sheet2").Range("G11:G12") = poinvnum
                                    'Set Indent for Customer Info
                                With Worksheets("Sheet2").Range("B15:B19")
                                    .IndentLevel = 1
                                End With
                                'Set Font Style for Range
                                With Worksheets("Sheet2").Range("B11:G19")
                                    .Font.Name = "Tahoma"
                                    .Font.Size = 8
                                End With
                                'Set Font Style for Range
                                With Worksheets("Sheet2").Range("A23:I29")
                                    .Font.Name = "Tahoma"
                                    .Font.Size = 8
                                End With
                            
                                'Read Cell Customer PO as FileName
                                FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)
                            
                                'Print to PDF
                                Worksheets("Sheet2").ExportAsFixedFormat _
                                Type:=x1TypePDF, _
                                FileName:=FileName, _
                                OpenAfterPublish:=False
                            
                                'Clear out Data for Next Invoice
                                Worksheets("Sheet2").Range("B15:B19").ClearContents
                                Worksheets("Sheet2").Range("G11:G19").ClearContents
                                Worksheets("Sheet2").Range("A23:I30").ClearContents
                                i = i + 7
                                End If
                            Else
                                'Date
                                Set podate = Worksheets("Sheet1").Range("A" & i)
                                    Worksheets("Sheet2").Range("G13:G14") = podate
                                'PO Number
                                Set ponum = Worksheets("Sheet1").Range("B" & i)
                                    Worksheets("Sheet2").Range("G15:G18") = ponum
                                'Name
                                Set poname = Worksheets("Sheet1").Range("C" & i)
                                    Worksheets("Sheet2").Range("B15") = poname
                                'Company
                                Set pocompany = Worksheets("Sheet1").Range("D" & i)
                                    Worksheets("Sheet2").Range("B16") = pocompany
                                'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                                Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
                                    If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                                        Worksheets("Sheet2").Range("B16") = poaddress1
                                        Else
                                        Worksheets("Sheet2").Range("B17") = poaddress1
                                    End If
                                'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                                Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
                                    If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                                        Worksheets("Sheet2").Range("B17") = poaddress2
                                        Else
                                        Worksheets("Sheet2").Range("B18") = poaddress2
                                    End If
                                'Define Array for Combining City, State, Zip
                                Arr(0) = Worksheets("Sheet1").Range("G" & i)
                                Arr(1) = Worksheets("Sheet1").Range("H" & i)
                                Arr(2) = Worksheets("Sheet1").Range("I" & i)
                                OriginalText = Join(Arr)
                                'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                                If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                                    Worksheets("Sheet2").Range("B17").Value = Join(Arr)
                                    'Add Comma after City Only
                                    CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                    Worksheets("Sheet2").Range("B17").Value = CorrectedText
                                    ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
                                    Worksheets("Sheet2").Range("B18").Value = Join(Arr)
                                    'Add Comma after City Only
                                    CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                    Worksheets("Sheet2").Range("B18").Value = CorrectedText
                                    Else
                                    Worksheets("Sheet2").Range("B19").Value = Join(Arr)
                                    'Add Comma after City Only
                                    CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                    Worksheets("Sheet2").Range("B19").Value = CorrectedText
                                End If
                                'Item Qty
                                Set poqty = Worksheets("Sheet1").Range("J" & i)
                                    Worksheets("Sheet2").Range("A23") = poqty
                                'Item Sku
                                Set posku = Worksheets("Sheet1").Range("K" & i)
                                    Worksheets("Sheet2").Range("B23") = posku
                                'Item Price
                                Set poprice = Worksheets("Sheet1").Range("L" & i)
                                    Worksheets("Sheet2").Range("H23") = poprice
                                'Order Total
                                Set pototal = Worksheets("Sheet1").Range("M" & i)
                                    Worksheets("Sheet2").Range("I23") = pototal
                                'Invoice Number
                                    Worksheets("Sheet2").Range("G11:G12") = poinvnum
                                'Set Indent for Customer Info
                            With Worksheets("Sheet2").Range("B15:B19")
                                .IndentLevel = 1
                            End With
                            'Set Font Style for Range
                            With Worksheets("Sheet2").Range("B11:G19")
                                .Font.Name = "Tahoma"
                                .Font.Size = 8
                            End With
                            'Set Font Style for Range
                            With Worksheets("Sheet2").Range("A23:I29")
                                .Font.Name = "Tahoma"
                                .Font.Size = 8
                            End With
                        
                            'Read Cell Customer PO as FileName
                            FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)
                        
                            'Print to PDF
                            Worksheets("Sheet2").ExportAsFixedFormat _
                            Type:=x1TypePDF, _
                            FileName:=FileName, _
                            OpenAfterPublish:=False
                        
                            'Clear out Data for Next Invoice
                            Worksheets("Sheet2").Range("B15:B19").ClearContents
                            Worksheets("Sheet2").Range("G11:G19").ClearContents
                            Worksheets("Sheet2").Range("A23:I30").ClearContents
                            i = i + 6
                            End If
                        Else
                            'Date
                            Set podate = Worksheets("Sheet1").Range("A" & i)
                                Worksheets("Sheet2").Range("G13:G14") = podate
                            'PO Number
                            Set ponum = Worksheets("Sheet1").Range("B" & i)
                                Worksheets("Sheet2").Range("G15:G18") = ponum
                            'Name
                            Set poname = Worksheets("Sheet1").Range("C" & i)
                                Worksheets("Sheet2").Range("B15") = poname
                            'Company
                            Set pocompany = Worksheets("Sheet1").Range("D" & i)
                                Worksheets("Sheet2").Range("B16") = pocompany
                            'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                            Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
                                If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                                    Worksheets("Sheet2").Range("B16") = poaddress1
                                    Else
                                    Worksheets("Sheet2").Range("B17") = poaddress1
                                End If
                            'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                            Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
                                If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                                    Worksheets("Sheet2").Range("B17") = poaddress2
                                    Else
                                    Worksheets("Sheet2").Range("B18") = poaddress2
                                End If
                            'Define Array for Combining City, State, Zip
                            Arr(0) = Worksheets("Sheet1").Range("G" & i)
                            Arr(1) = Worksheets("Sheet1").Range("H" & i)
                            Arr(2) = Worksheets("Sheet1").Range("I" & i)
                            OriginalText = Join(Arr)
                            'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                            If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                                Worksheets("Sheet2").Range("B17").Value = Join(Arr)
                                'Add Comma after City Only
                                CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                Worksheets("Sheet2").Range("B17").Value = CorrectedText
                                ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
                                Worksheets("Sheet2").Range("B18").Value = Join(Arr)
                                'Add Comma after City Only
                                CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                Worksheets("Sheet2").Range("B18").Value = CorrectedText
                                Else
                                Worksheets("Sheet2").Range("B19").Value = Join(Arr)
                                'Add Comma after City Only
                                CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                                Worksheets("Sheet2").Range("B19").Value = CorrectedText
                            End If
                            'Item Qty
                            Set poqty = Worksheets("Sheet1").Range("J" & i)
                                Worksheets("Sheet2").Range("A23") = poqty
                            'Item Sku
                            Set posku = Worksheets("Sheet1").Range("K" & i)
                                Worksheets("Sheet2").Range("B23") = posku
                            'Item Price
                            Set poprice = Worksheets("Sheet1").Range("L" & i)
                                Worksheets("Sheet2").Range("H23") = poprice
                            'Order Total
                            Set pototal = Worksheets("Sheet1").Range("M" & i)
                                Worksheets("Sheet2").Range("I23") = pototal
                            'Invoice Number
                                Worksheets("Sheet2").Range("G11:G12") = poinvnum
                            'Set Indent for Customer Info
                        With Worksheets("Sheet2").Range("B15:B19")
                            .IndentLevel = 1
                        End With
                        'Set Font Style for Range
                        With Worksheets("Sheet2").Range("B11:G19")
                            .Font.Name = "Tahoma"
                            .Font.Size = 8
                        End With
                        'Set Font Style for Range
                        With Worksheets("Sheet2").Range("A23:I29")
                            .Font.Name = "Tahoma"
                            .Font.Size = 8
                        End With
                    
                        'Read Cell Customer PO as FileName
                        FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)
                    
                        'Print to PDF
                        Worksheets("Sheet2").ExportAsFixedFormat _
                        Type:=x1TypePDF, _
                        FileName:=FileName, _
                        OpenAfterPublish:=False
                    
                        'Clear out Data for Next Invoice
                        Worksheets("Sheet2").Range("B15:B19").ClearContents
                        Worksheets("Sheet2").Range("G11:G19").ClearContents
                        Worksheets("Sheet2").Range("A23:I30").ClearContents
                        i = i + 5
                        End If
                    Else
                        'Date
                        Set podate = Worksheets("Sheet1").Range("A" & i)
                            Worksheets("Sheet2").Range("G13:G14") = podate
                        'PO Number
                        Set ponum = Worksheets("Sheet1").Range("B" & i)
                            Worksheets("Sheet2").Range("G15:G18") = ponum
                        'Name
                        Set poname = Worksheets("Sheet1").Range("C" & i)
                            Worksheets("Sheet2").Range("B15") = poname
                        'Company
                        Set pocompany = Worksheets("Sheet1").Range("D" & i)
                            Worksheets("Sheet2").Range("B16") = pocompany
                        'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                        Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
                            If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                                Worksheets("Sheet2").Range("B16") = poaddress1
                                Else
                                Worksheets("Sheet2").Range("B17") = poaddress1
                            End If
                        'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                        Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
                            If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                                Worksheets("Sheet2").Range("B17") = poaddress2
                                Else
                                Worksheets("Sheet2").Range("B18") = poaddress2
                            End If
                        'Define Array for Combining City, State, Zip
                        Arr(0) = Worksheets("Sheet1").Range("G" & i)
                        Arr(1) = Worksheets("Sheet1").Range("H" & i)
                        Arr(2) = Worksheets("Sheet1").Range("I" & i)
                        OriginalText = Join(Arr)
                        'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                        If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                            Worksheets("Sheet2").Range("B17").Value = Join(Arr)
                            'Add Comma after City Only
                            CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                            Worksheets("Sheet2").Range("B17").Value = CorrectedText
                            ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
                            Worksheets("Sheet2").Range("B18").Value = Join(Arr)
                            'Add Comma after City Only
                            CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                            Worksheets("Sheet2").Range("B18").Value = CorrectedText
                            Else
                            Worksheets("Sheet2").Range("B19").Value = Join(Arr)
                            'Add Comma after City Only
                            CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                            Worksheets("Sheet2").Range("B19").Value = CorrectedText
                        End If
                        'Item Qty
                        Set poqty = Worksheets("Sheet1").Range("J" & i)
                            Worksheets("Sheet2").Range("A23") = poqty
                        'Item Sku
                        Set posku = Worksheets("Sheet1").Range("K" & i)
                            Worksheets("Sheet2").Range("B23") = posku
                        'Item Price
                        Set poprice = Worksheets("Sheet1").Range("L" & i)
                            Worksheets("Sheet2").Range("H23") = poprice
                        'Order Total
                        Set pototal = Worksheets("Sheet1").Range("M" & i)
                            Worksheets("Sheet2").Range("I23") = pototal
                        'Invoice Number
                            Worksheets("Sheet2").Range("G11:G12") = poinvnum
                        'Set Indent for Customer Info
                    With Worksheets("Sheet2").Range("B15:B19")
                        .IndentLevel = 1
                    End With
                    'Set Font Style for Range
                    With Worksheets("Sheet2").Range("B11:G19")
                        .Font.Name = "Tahoma"
                        .Font.Size = 8
                    End With
                    'Set Font Style for Range
                    With Worksheets("Sheet2").Range("A23:I29")
                        .Font.Name = "Tahoma"
                        .Font.Size = 8
                    End With
                
                    'Read Cell Customer PO as FileName
                    FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)
                
                    'Print to PDF
                    Worksheets("Sheet2").ExportAsFixedFormat _
                    Type:=x1TypePDF, _
                    FileName:=FileName, _
                    OpenAfterPublish:=False
                
                    'Clear out Data for Next Invoice
                    Worksheets("Sheet2").Range("B15:B19").ClearContents
                    Worksheets("Sheet2").Range("G11:G19").ClearContents
                    Worksheets("Sheet2").Range("A23:I30").ClearContents
                    i = i + 4
                    End If
                Else
                    'Date
                    Set podate = Worksheets("Sheet1").Range("A" & i)
                        Worksheets("Sheet2").Range("G13:G14") = podate
                    'PO Number
                    Set ponum = Worksheets("Sheet1").Range("B" & i)
                        Worksheets("Sheet2").Range("G15:G18") = ponum
                    'Name
                    Set poname = Worksheets("Sheet1").Range("C" & i)
                        Worksheets("Sheet2").Range("B15") = poname
                    'Company
                    Set pocompany = Worksheets("Sheet1").Range("D" & i)
                        Worksheets("Sheet2").Range("B16") = pocompany
                    'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                    Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
                        If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                            Worksheets("Sheet2").Range("B16") = poaddress1
                            Else
                            Worksheets("Sheet2").Range("B17") = poaddress1
                        End If
                    'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                    Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
                        If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                            Worksheets("Sheet2").Range("B17") = poaddress2
                            Else
                            Worksheets("Sheet2").Range("B18") = poaddress2
                        End If
                    'Define Array for Combining City, State, Zip
                    Arr(0) = Worksheets("Sheet1").Range("G" & i)
                    Arr(1) = Worksheets("Sheet1").Range("H" & i)
                    Arr(2) = Worksheets("Sheet1").Range("I" & i)
                    OriginalText = Join(Arr)
                    'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                    If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                        Worksheets("Sheet2").Range("B17").Value = Join(Arr)
                        'Add Comma after City Only
                        CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                        Worksheets("Sheet2").Range("B17").Value = CorrectedText
                        ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
                        Worksheets("Sheet2").Range("B18").Value = Join(Arr)
                        'Add Comma after City Only
                        CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                        Worksheets("Sheet2").Range("B18").Value = CorrectedText
                        Else
                        Worksheets("Sheet2").Range("B19").Value = Join(Arr)
                        'Add Comma after City Only
                        CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                        Worksheets("Sheet2").Range("B19").Value = CorrectedText
                    End If
                    'Item Qty
                    Set poqty = Worksheets("Sheet1").Range("J" & i)
                        Worksheets("Sheet2").Range("A23") = poqty
                    'Item Sku
                    Set posku = Worksheets("Sheet1").Range("K" & i)
                        Worksheets("Sheet2").Range("B23") = posku
                    'Item Price
                    Set poprice = Worksheets("Sheet1").Range("L" & i)
                        Worksheets("Sheet2").Range("H23") = poprice
                    'Order Total
                    Set pototal = Worksheets("Sheet1").Range("M" & i)
                        Worksheets("Sheet2").Range("I23") = pototal
                    'Invoice Number
                        Worksheets("Sheet2").Range("G11:G12") = poinvnum
                    'Set Indent for Customer Info
                With Worksheets("Sheet2").Range("B15:B19")
                    .IndentLevel = 1
                End With
                'Set Font Style for Range
                With Worksheets("Sheet2").Range("B11:G19")
                    .Font.Name = "Tahoma"
                    .Font.Size = 8
                End With
                'Set Font Style for Range
                With Worksheets("Sheet2").Range("A23:I29")
                    .Font.Name = "Tahoma"
                    .Font.Size = 8
                End With
            
                'Read Cell Customer PO as FileName
                FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)
            
                'Print to PDF
                Worksheets("Sheet2").ExportAsFixedFormat _
                Type:=x1TypePDF, _
                FileName:=FileName, _
                OpenAfterPublish:=False
            
                'Clear out Data for Next Invoice
                Worksheets("Sheet2").Range("B15:B19").ClearContents
                Worksheets("Sheet2").Range("G11:G19").ClearContents
                Worksheets("Sheet2").Range("A23:I30").ClearContents
                i = i + 3
                End If
            Else
                'Date
                Set podate = Worksheets("Sheet1").Range("A" & i)
                    Worksheets("Sheet2").Range("G13:G14") = podate
                'PO Number
                Set ponum = Worksheets("Sheet1").Range("B" & i)
                    Worksheets("Sheet2").Range("G15:G18") = ponum
                'Name
                Set poname = Worksheets("Sheet1").Range("C" & i)
                    Worksheets("Sheet2").Range("B15") = poname
                'Company
                Set pocompany = Worksheets("Sheet1").Range("D" & i)
                    Worksheets("Sheet2").Range("B16") = pocompany
                'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
                    If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                        Worksheets("Sheet2").Range("B16") = poaddress1
                        Else
                        Worksheets("Sheet2").Range("B17") = poaddress1
                    End If
                'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
                    If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                        Worksheets("Sheet2").Range("B17") = poaddress2
                        Else
                        Worksheets("Sheet2").Range("B18") = poaddress2
                    End If
                'Define Array for Combining City, State, Zip
                Arr(0) = Worksheets("Sheet1").Range("G" & i)
                Arr(1) = Worksheets("Sheet1").Range("H" & i)
                Arr(2) = Worksheets("Sheet1").Range("I" & i)
                OriginalText = Join(Arr)
                'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
                If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                    Worksheets("Sheet2").Range("B17").Value = Join(Arr)
                    'Add Comma after City Only
                    CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                    Worksheets("Sheet2").Range("B17").Value = CorrectedText
                    ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
                    Worksheets("Sheet2").Range("B18").Value = Join(Arr)
                    'Add Comma after City Only
                    CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                    Worksheets("Sheet2").Range("B18").Value = CorrectedText
                    Else
                    Worksheets("Sheet2").Range("B19").Value = Join(Arr)
                    'Add Comma after City Only
                    CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                    Worksheets("Sheet2").Range("B19").Value = CorrectedText
                End If
                'Item Qty
                Set poqty = Worksheets("Sheet1").Range("J" & i)
                    Worksheets("Sheet2").Range("A23") = poqty
                'Item Sku
                Set posku = Worksheets("Sheet1").Range("K" & i)
                    Worksheets("Sheet2").Range("B23") = posku
                'Item Price
                Set poprice = Worksheets("Sheet1").Range("L" & i)
                    Worksheets("Sheet2").Range("H23") = poprice
                'Order Total
                Set pototal = Worksheets("Sheet1").Range("M" & i)
                    Worksheets("Sheet2").Range("I23") = pototal
                'Invoice Number
                    Worksheets("Sheet2").Range("G11:G12") = poinvnum
                'Set Indent for Customer Info
            With Worksheets("Sheet2").Range("B15:B19")
                .IndentLevel = 1
            End With
            'Set Font Style for Range
            With Worksheets("Sheet2").Range("B11:G19")
                .Font.Name = "Tahoma"
                .Font.Size = 8
            End With
            'Set Font Style for Range
            With Worksheets("Sheet2").Range("A23:I29")
                .Font.Name = "Tahoma"
                .Font.Size = 8
            End With
        
            'Read Cell Customer PO as FileName
            FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)
        
            'Print to PDF
            Worksheets("Sheet2").ExportAsFixedFormat _
            Type:=x1TypePDF, _
            FileName:=FileName, _
            OpenAfterPublish:=False
        
            'Clear out Data for Next Invoice
            Worksheets("Sheet2").Range("B15:B19").ClearContents
            Worksheets("Sheet2").Range("G11:G19").ClearContents
            Worksheets("Sheet2").Range("A23:I30").ClearContents
            i = i + 2
            End If
        Else
            'Date
            Set podate = Worksheets("Sheet1").Range("A" & i)
                Worksheets("Sheet2").Range("G13:G14") = podate
            'PO Number
            Set ponum = Worksheets("Sheet1").Range("B" & i)
                Worksheets("Sheet2").Range("G15:G18") = ponum
            'Name
            Set poname = Worksheets("Sheet1").Range("C" & i)
                Worksheets("Sheet2").Range("B15") = poname
            'Company
            Set pocompany = Worksheets("Sheet1").Range("D" & i)
                Worksheets("Sheet2").Range("B16") = pocompany
            'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
            Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
                If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                    Worksheets("Sheet2").Range("B16") = poaddress1
                    Else
                    Worksheets("Sheet2").Range("B17") = poaddress1
                End If
            'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
            Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
                If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                    Worksheets("Sheet2").Range("B17") = poaddress2
                    Else
                    Worksheets("Sheet2").Range("B18") = poaddress2
                End If
            'Define Array for Combining City, State, Zip
            Arr(0) = Worksheets("Sheet1").Range("G" & i)
            Arr(1) = Worksheets("Sheet1").Range("H" & i)
            Arr(2) = Worksheets("Sheet1").Range("I" & i)
            OriginalText = Join(Arr)
            'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
            If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                Worksheets("Sheet2").Range("B17").Value = Join(Arr)
                'Add Comma after City Only
                CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                Worksheets("Sheet2").Range("B17").Value = CorrectedText
                ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
                Worksheets("Sheet2").Range("B18").Value = Join(Arr)
                'Add Comma after City Only
                CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                Worksheets("Sheet2").Range("B18").Value = CorrectedText
                Else
                Worksheets("Sheet2").Range("B19").Value = Join(Arr)
                'Add Comma after City Only
                CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
                Worksheets("Sheet2").Range("B19").Value = CorrectedText
            End If
            'Item Qty
            Set poqty = Worksheets("Sheet1").Range("J" & i)
                Worksheets("Sheet2").Range("A23") = poqty
            'Item Sku
            Set posku = Worksheets("Sheet1").Range("K" & i)
                Worksheets("Sheet2").Range("B23") = posku
            'Item Price
            Set poprice = Worksheets("Sheet1").Range("L" & i)
                Worksheets("Sheet2").Range("H23") = poprice
            'Order Total
            Set pototal = Worksheets("Sheet1").Range("M" & i)
                Worksheets("Sheet2").Range("I23") = pototal
            'Invoice Number
                Worksheets("Sheet2").Range("G11:G12") = poinvnum
            'Set Indent for Customer Info
        With Worksheets("Sheet2").Range("B15:B19")
            .IndentLevel = 1
        End With
        'Set Font Style for Range
        With Worksheets("Sheet2").Range("B11:G19")
            .Font.Name = "Tahoma"
            .Font.Size = 8
        End With
        'Set Font Style for Range
        With Worksheets("Sheet2").Range("A23:I29")
            .Font.Name = "Tahoma"
            .Font.Size = 8
        End With
    
        'Read Cell Customer PO as FileName
        FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)
    
        'Print to PDF
        Worksheets("Sheet2").ExportAsFixedFormat _
        Type:=x1TypePDF, _
        FileName:=FileName, _
        OpenAfterPublish:=False
    
        'Clear out Data for Next Invoice
        Worksheets("Sheet2").Range("B15:B19").ClearContents
        Worksheets("Sheet2").Range("G11:G19").ClearContents
        Worksheets("Sheet2").Range("A23:I30").ClearContents
        i = i + 1
        End If
    Else
        'Date
        Set podate = Worksheets("Sheet1").Range("A" & i)
            Worksheets("Sheet2").Range("G13:G14") = podate
        'PO Number
        Set ponum = Worksheets("Sheet1").Range("B" & i)
            Worksheets("Sheet2").Range("G15:G18") = ponum
        'Name
        Set poname = Worksheets("Sheet1").Range("C" & i)
            Worksheets("Sheet2").Range("B15") = poname
        'Company
        Set pocompany = Worksheets("Sheet1").Range("D" & i)
            Worksheets("Sheet2").Range("B16") = pocompany
        'Address 1, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
        Set poaddress1 = Worksheets("Sheet1").Range("E" & i)
            If IsEmpty(Worksheets("Sheet2").Range("B16")) = True Then
                Worksheets("Sheet2").Range("B16") = poaddress1
                Else
                Worksheets("Sheet2").Range("B17") = poaddress1
            End If
        'Address 2, If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
        Set poaddress2 = Worksheets("Sheet1").Range("F" & i)
            If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
                Worksheets("Sheet2").Range("B17") = poaddress2
                Else
                Worksheets("Sheet2").Range("B18") = poaddress2
            End If
        'Define Array for Combining City, State, Zip
        Arr(0) = Worksheets("Sheet1").Range("G" & i)
        Arr(1) = Worksheets("Sheet1").Range("H" & i)
        Arr(2) = Worksheets("Sheet1").Range("I" & i)
        OriginalText = Join(Arr)
        'City, State, Zip & If Then Else Statement Checks to see if Previouly Line is Blank Then Copy Paste Next Line Accordingly
        If IsEmpty(Worksheets("Sheet2").Range("B17")) = True Then
            Worksheets("Sheet2").Range("B17").Value = Join(Arr)
            'Add Comma after City Only
            CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
            Worksheets("Sheet2").Range("B17").Value = CorrectedText
            ElseIf IsEmpty(Worksheets("Sheet2").Range("B18")) = True Then
            Worksheets("Sheet2").Range("B18").Value = Join(Arr)
            'Add Comma after City Only
            CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
            Worksheets("Sheet2").Range("B18").Value = CorrectedText
            Else
            Worksheets("Sheet2").Range("B19").Value = Join(Arr)
            'Add Comma after City Only
            CorrectedText = Replace(OriginalText, " ", ", ", Count:=1)
            Worksheets("Sheet2").Range("B19").Value = CorrectedText
        End If
        'Item Qty
        Set poqty = Worksheets("Sheet1").Range("J" & i)
            Worksheets("Sheet2").Range("A23") = poqty
        'Item Sku
        Set posku = Worksheets("Sheet1").Range("K" & i)
            Worksheets("Sheet2").Range("B23") = posku
        'Item Price
        Set poprice = Worksheets("Sheet1").Range("L" & i)
            Worksheets("Sheet2").Range("H23") = poprice
        'Order Total
        Set pototal = Worksheets("Sheet1").Range("M" & i)
            Worksheets("Sheet2").Range("I23") = pototal
        'Invoice Number
            Worksheets("Sheet2").Range("G11:G12") = poinvnum
        'Set Indent for Customer Info
    With Worksheets("Sheet2").Range("B15:B19")
        .IndentLevel = 1
    End With
    'Set Font Style for Range
    With Worksheets("Sheet2").Range("B11:G19")
        .Font.Name = "Tahoma"
        .Font.Size = 8
    End With
    'Set Font Style for Range
    With Worksheets("Sheet2").Range("A23:I29")
        .Font.Name = "Tahoma"
        .Font.Size = 8
    End With

    'Read Cell Customer PO as FileName
    FileName = "Inv_" & Worksheets("Sheet1").Range("N" & i) & " PO_" & Worksheets("Sheet1").Range("B" & i)

    'Print to PDF
    Worksheets("Sheet2").ExportAsFixedFormat _
    Type:=x1TypePDF, _
    FileName:=FileName, _
    OpenAfterPublish:=False

    'Clear out Data for Next Invoice
    Worksheets("Sheet2").Range("B15:B19").ClearContents
    Worksheets("Sheet2").Range("G11:G19").ClearContents
    Worksheets("Sheet2").Range("A23:I30").ClearContents
    End If
Next i

Application.CutCopyMode = False
 
End Sub
