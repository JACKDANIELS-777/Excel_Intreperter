Private Sub CommandButton1_Click()
    Dim txt As String
    Dim col As Collection
    Dim i As Integer
    Dim word As String

    txt = TextBox1.Text
    Set col = lex(txt)

    word = "" ' Initialize word

    For i = 1 To col.Count ' Use Count property
        word = col(i) ' Get the current word
        MsgBox word
        
        
        ' Handle if_print_cell_val
        If word = "if_print_cell_val" Then
            ' if_print_cell_val; A1 =; 10; print; A1; 10;
            If i + 6 <= col.Count Then
                If col(i + 2) = "=" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = ">" Then
                'MsgBox Evaluate("=" & col(i + 3))
                
                
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value > Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value < Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<>" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value <> Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                End If
            End If
            i = i + 6
            
        ' Handle if_print_val_cell
        ElseIf word = "if_print_val_cell" Then
            If i + 6 <= col.Count Then
                If col(i + 2) = "=" Then
                    If Evaluate(col(i + 1)) = ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = ">" Then
                    If Evaluate(col(i + 1)) > ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<" Then
                    If Evaluate(col(i + 1)) < ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<>" Then
                    If Evaluate(col(i + 1)) <> ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        MsgBox (col(i + 5))
                    End If
                End If
            End If
            i = i + 6
            
        ' Handle for_inc_cell_by_var
        ElseIf word = "for_inc_cell_by_var" Then
            If i + 6 <= col.Count Then
                Dim j As Integer
                For j = CInt(col(i + 2)) To CInt(col(i + 3))
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = ThisWorkbook.Sheets(1).Range(col(i + 5)).Value + j
                Next j
            End If
            i = i + 6
            
        ' Handle for_each
        ElseIf word = "for_each" Then
            Dim rng As Range
            

            ' foreach;cell;in;range;action
            If i + 1 <= col.Count Then
                If col(i + 1) = "cell" Then
                    ' foreach;cell;in;range;print;cell;
                    If i + 1 + 6 <= col.Count Then
                        If col(i + 2) = "in" And col(i + 4) = "print" And col(i + 5) = "cell" Then
                            Set rng = ThisWorkbook.Sheets(1).Range(col(i + 3))
                            If col(i + 6) = "+" Then
                                For Each cell In rng
                                   
                                    cell.Value = cell.Value + CInt(col(i + 7))
                                Next cell
                            ElseIf col(i + 6) = "-" Then
                                For Each cell In rng
                                    cell.Value = cell.Value - CInt(col(i + 7))
                                Next cell
                            ElseIf col(i + 6) = "*" Then
                                For Each cell In rng
                                    cell.Value = cell.Value * CInt(col(i + 7))
                                Next cell
                            ElseIf col(i + 6) = "/" Then
                                For Each cell In rng
                                    cell.Value = cell.Value / CInt(col(i + 7))
                                Next cell
                            End If
                            Set rng = Nothing
                        End If
                        i = i + 6
                    
                    
                    ElseIf i + 1 + 4 <= col.Count Then
                        If col(i + 2) = "in" And col(i + 4) = "msg" And col(i + 5) = "cell" Then
                            Set rng = ThisWorkbook.Sheets(1).Range(col(i + 3))
                            For Each cell In rng
                                MsgBox (cell.Value)
                            Next cell
                            Set rng = Nothing
                        End If
                        i = i + 5
                    End If
                End If
            End If
            
        ' Handle print
        ElseIf word = "print" Then
            'check if they want to have the formula in the cell not just the value
            
            ' Ensure there are enough elements in the collection
            If i + 2 <= col.Count Then
                Dim res As Variant
                
                 On Error Resume Next ' Ignore errors
                res = Evaluate(col(i + 2)) ' Attempt to evaluate the expression
                MsgBox res & "hey"
                'MsgBox CStr(res) ' Show the result of the evaluation
                If res = Err.Number Then
                res = col(i + 2)
                End If
                

    
    ' Reset error handling
                ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = res
                i = i + 2 ' Move index forward by 2 to skip the next two items
            Else
                MsgBox "Not enough arguments after 'print'."
            End If
            
        ' Handle clear_cell
        ElseIf word = "clear_cell" Then
            ' Ensure the next element is available
            If i + 1 <= col.Count Then
                ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = ""
                i = i + 1 ' Move index forward to skip the next item
            Else
                MsgBox "Not enough arguments after 'clear'."
            End If
            
        ' Handle clear_range
        ElseIf word = "clear_range" Then
            If i + 2 <= col.Count Then
                ThisWorkbook.Sheets(1).Range(col(i + 1), col(i + 2)).Value = ""
                i = i + 2
            End If
            
        ' Handle add_range_dup
        ElseIf word = "add_range_dup" Then
            If i + 3 <= col.Count Then
                ThisWorkbook.Sheets(1).Range(col(i + 1), col(i + 2)).Value = col(i + 3)
                i = i + 3
            End If
            
        ' Handle add_range_inc
        ElseIf word = "add_range_inc" Then
            If i + 3 <= col.Count Then
                Dim startCell As String
                Dim endCell As String
                Dim incrementValue As Integer
                Dim og_inc_value As Integer
                
                ' Get the arguments from the collection
                startCell = col(i + 1) ' e.g., "A1"
                endCell = col(i + 2)   ' e.g., "A10"
                incrementValue = CInt(col(i + 3)) ' e.g., 10
                og_inc_value = CInt(col(i + 3))
                
                ' Loop through each cell in the specified range
                For Each cell In ThisWorkbook.Sheets(1).Range(startCell & ":" & endCell)
                    cell.Value = incrementValue ' Increment the cell value
                    incrementValue = incrementValue + og_inc_value
                Next cell
                
                i = i + 3 ' Move index forward to skip processed items
            Else
                MsgBox "Not enough arguments after 'add_range_inc'."
            End If
            
        ' Handle msg
        ElseIf word = "msg" Then
            If i + 2 <= col.Count Then
                MsgBox col(i + 1), , col(i + 2)
                i = i + 2
            End If
            
        ' Handle msg_cell
        ElseIf word = "msg_cell" Then
            If i + 2 <= col.Count Then
                MsgBox ThisWorkbook.Sheets(1).Range(col(i + 1)).Value, , ThisWorkbook.Sheets(1).Range(col(i + 2))
                i = i + 2
            End If
            
        ' Handle del_sheet
        ElseIf word = "del_sheet" Then
            If i + 1 <= col.Count Then
                ThisWorkbook.Sheets(col(i + 1)).Delete
                i = i ' No need to increment i here
            End If
            
        ' Handle rename_sheet
        ElseIf word = "rename_sheet" Then
            If i + 2 <= col.Count Then
                ThisWorkbook.Sheets(col(i + 1)).Name = col(i + 2)
                i = i + 1
            End If
            
        ' Handle add_sheet
        ElseIf word = "add_sheet" Then
            If i + 1 <= col.Count Then
                ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = col(i + 1)
            End If
            
        ' Handle comment
        ElseIf word = "comment" Then
            If i + 1 <= col.Count Then
                i = i + 1 ' Increment i if needed
            End If
        End If
    Next i
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 ' Prevent the beep sound
        TextBox1.Value = TextBox1.Value & vbCrLf ' Add a new line
    End If
End Sub

