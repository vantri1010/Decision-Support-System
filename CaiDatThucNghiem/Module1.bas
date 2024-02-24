Attribute VB_Name = "Module1"
Sub test()
Value1 = ThisWorkbook.Sheets(1).Range("A1").Value

Dim i As Integer
Dim j As Integer

For j = 1 To 16
'Hang cheo ma tran = 1 va background mau xanh
ThisWorkbook.Sheets(5).Cells(j, j) = 1
ThisWorkbook.Sheets(5).Cells(j, j).Interior.ColorIndex = 37
For i = j To 16
    '4 cai IF ngoai la de xet xem ben sheet 1 la low high hay ?
    If LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[very]*") Then
        '4 cai IF ben duoi nay la de xet nhung khia canh con lai so voi khia canh nay la j de in ra gia tri
        If LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1
         ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 2
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 5
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 8
        End If
    
    ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[very]*" Then
         If LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 2
         ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 3
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 7
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 9
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*medium*" Then
        If LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 2
         ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 3
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 4
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 6
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[very]*") Then
        If LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 5
         ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 7
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 4
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1 / 2
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & j)) Like "*[very]*" Then
        If LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 8
         ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 9
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 6
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(1).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(5).Range("" + Chr(i + 65) & j).Value = 1
        End If
    End If
    
  
Next i
Next j

For j = 1 To 16
For i = j + 1 To 17
'De in nua ma tran ahp con lai vi co gia tri dao nguoc nen phai co dong nay de khoi 1/0
If ThisWorkbook.Sheets(5).Range("" + Chr(i - 2 + 65) & j).Value <> 0 And i > j + 1 Then
 ThisWorkbook.Sheets(5).Range("" + Chr(j + 64) & i - 1).Value = 1 / (ThisWorkbook.Sheets(5).Range("" + Chr(i - 2 + 65) & j).Value)
End If
Next i
Next j

'Task
For j = 1 To 16
'Hang cheo ma tran = 1 va background mau xanh
ThisWorkbook.Sheets(6).Cells(j, j) = 1
ThisWorkbook.Sheets(6).Cells(j, j).Interior.ColorIndex = 37
For i = j To 16
    '4 cai IF ngoai la de xet xem ben sheet 1 la low high hay ?
    If LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[very]*") Then
        '4 cai IF ben duoi nay la de xet nhung khia canh con lai so voi khia canh nay la j de in ra gia tri
        If LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1
         ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 2
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 5
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 8
        End If
    
    ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[very]*" Then
         If LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 2
         ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 3
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 7
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 9
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*medium*" Then
        If LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 2
         ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 3
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 4
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 6
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[very]*") Then
        If LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 5
         ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 7
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 4
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1 / 2
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & j)) Like "*[very]*" Then
        If LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 8
         ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 9
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 6
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(2).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(6).Range("" + Chr(i + 65) & j).Value = 1
        End If
    End If
    
  
Next i
Next j

For j = 1 To 16
For i = j + 1 To 17
'De in nua ma tran ahp con lai vi co gia tri dao nguoc nen phai co dong nay de khoi 1/0
If ThisWorkbook.Sheets(6).Range("" + Chr(i - 2 + 65) & j).Value <> 0 And i > j + 1 Then
 ThisWorkbook.Sheets(6).Range("" + Chr(j + 64) & i - 1).Value = 1 / (ThisWorkbook.Sheets(6).Range("" + Chr(i - 2 + 65) & j).Value)
End If
Next i
Next j

'Struc


For j = 1 To 8
'Hang cheo ma tran = 1 va background mau xanh
ThisWorkbook.Sheets(7).Cells(j, j) = 1
ThisWorkbook.Sheets(7).Cells(j, j).Interior.ColorIndex = 37
For i = j To 8
    '4 cai IF ngoai la de xet xem ben sheet 1 la low high hay ?
    If LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[very]*") Then
        '4 cai IF ben duoi nay la de xet nhung khia canh con lai so voi khia canh nay la j de in ra gia tri
        If LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1
         ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 2
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 5
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 8
        End If
    
    ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[very]*" Then
         If LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 2
         ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 3
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 7
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 9
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*medium*" Then
        If LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 2
         ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 3
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 4
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 6
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[very]*") Then
        If LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 5
         ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 7
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 4
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1 / 2
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & j)) Like "*[very]*" Then
        If LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 8
         ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 9
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 6
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(3).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(7).Range("" + Chr(i + 65) & j).Value = 1
        End If
    End If
    
  
Next i
Next j

For j = 1 To 8
For i = j + 1 To 9
'De in nua ma tran ahp con lai vi co gia tri dao nguoc nen phai co dong nay de khoi 1/0
If ThisWorkbook.Sheets(7).Range("" + Chr(i - 2 + 65) & j).Value <> 0 And i > j + 1 Then
 ThisWorkbook.Sheets(7).Range("" + Chr(j + 64) & i - 1).Value = 1 / (ThisWorkbook.Sheets(7).Range("" + Chr(i - 2 + 65) & j).Value)
End If
Next i
Next j

'Tech

For j = 1 To 9
'Hang cheo ma tran = 1 va background mau xanh
ThisWorkbook.Sheets(8).Cells(j, j) = 1
ThisWorkbook.Sheets(8).Cells(j, j).Interior.ColorIndex = 37
For i = j To 9
    '4 cai IF ngoai la de xet xem ben sheet 1 la low high hay ?
    If LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[very]*") Then
        '4 cai IF ben duoi nay la de xet nhung khia canh con lai so voi khia canh nay la j de in ra gia tri
        If LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1
         ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 2
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 5
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 8
        End If
    
    ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[very]*" Then
         If LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 2
         ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 3
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 7
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 9
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*medium*" Then
        If LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 2
         ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 3
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 4
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 6
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[very]*") Then
        If LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 5
         ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 7
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 4
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1 / 2
        End If
        
    ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & j)) Like "*[very]*" Then
        If LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 8
         ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[low]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
           ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 9
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*medium*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 6
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And Not (LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*") Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 2
        ElseIf LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[high]*" And LCase(ThisWorkbook.Sheets(4).Range("B" & i + 1)) Like "*[very]*" Then
            ThisWorkbook.Sheets(8).Range("" + Chr(i + 65) & j).Value = 1
        End If
    End If
    
  
Next i
Next j

For j = 1 To 9
For i = j + 1 To 10
'De in nua ma tran ahp con lai vi co gia tri dao nguoc nen phai co dong nay de khoi 1/0
If ThisWorkbook.Sheets(8).Range("" + Chr(i - 2 + 65) & j).Value <> 0 And i > j + 1 Then
 ThisWorkbook.Sheets(8).Range("" + Chr(j + 64) & i - 1).Value = 1 / (ThisWorkbook.Sheets(8).Range("" + Chr(i - 2 + 65) & j).Value)
End If
Next i
Next j





End Sub
