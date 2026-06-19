Function donotit()
Application.Volatile
donotit = Evaluate("=SUM(A1:A20)")
 End Function


Function lex(txt As String) As Collection
    Dim i As Integer
    Dim str As String
    Dim tok As String
    Dim forms As Collection
    Set forms = New Collection
    
    
    Dim col As Collection
    Set col = New Collection ' Initialize the collection
    Dim quoteCount As Integer ' Track the number of quotes
    Dim stack As Collection
    Set stack = New Collection
    Dim b_num As Integer
    b_num = 0
    quoteCount = 0
    str = ""
    tok = ""

    For i = 1 To Len(txt)
        Dim curchar As String
        curchar = Mid(txt, i, 1)

        If curchar = "'" Then
            quoteCount = quoteCount + 1 ' Increment the quote count

            If quoteCount Mod 2 = 1 Then
                ' Starting a new quote: capture everything until the next closing quote
                If tok <> "" Then
                    col.Add tok ' Add the accumulated token before entering a quote
                    tok = "" ' Reset for new content
                End If
                str = "" ' Reset for quoted string
            Else
                ' Closing the quote: add the captured string
                
                col.Add "statement"
                col.Add str ' Add the accumulated quoted content
                str = "" ' Reset for next content
            End If
        ElseIf quoteCount Mod 2 = 1 Then
            ' Inside a quote: accumulate characters (including brackets)
            str = str & curchar
            

        ElseIf curchar = "[" And quoteCount = 0 Then
            ' Start capturing for brackets
           
            If tok <> "" Then
                col.Add tok ' Add the accumulated token before entering brackets
                tok = "" ' Reset for new content
            End If
            If b_num = 1 Then
            str = str & curchar ' Include the bracket in the string
            End If
            
            b_num = 1
            stack.Add "open" ' Mark that we've encountered an opening bracket
            
            
        ElseIf b_num = 1 And curchar <> "]" Then
            ' Accumulate characters inside brackets
            str = str & curchar
            
        ElseIf curchar = "]" Then
            ' Closing a bracket: add the captured string
            
            
            If stack.Count > 0 Then
                stack.Remove stack.Count ' Remove the last opened bracket
                If stack.Count = 1 Then
                str = str & curchar ' Include the closing bracket in the string
                End If
                If stack.Count = 0 Then
                    col.Add "statement"
                    col.Add Mid(str, 1, Len(str) - 1) ' Add the accumulated bracketed content
                    MsgBox str & "str:"
                    str = "" ' Reset for next content
                    b_num = 0 ' Reset bracket number
                End If
                
            End If
        ElseIf curchar <> ";" Then
            If curchar <> vbCr And curchar <> vbLf Then
                tok = tok & curchar ' Accumulate characters
            End If
        Else
            ' Add the token to the collection if it's not empty
            If tok <> "" Then
                col.Add tok ' Add the token to the collection
            End If
            tok = "" ' Reset the token
        End If
    Next i

    ' Add any remaining token or string
    If tok <> "" Then
        col.Add tok
    End If
    If str <> "" Then
        If quoteCount Mod 2 = 1 Then
            col.Add "statement"
            col.Add str
        ElseIf stack.Count > 0 Then
            col.Add "statement"
            col.Add str
        End If
    End If

    ' Return the collection of tokens
    Set lex = col
End Function







Function parser(col As Collection)
Dim variableTable As Collection
    Set variableTable = New Collection
Dim arr_col As Collection
    Set arr_col = New Collection
    Dim filePath As String
            Dim fileNum As Integer
For i = 1 To col.Count
        word = col(i)
        'MsgBox word & "word"

        ' Handle if_print_cell_val
        If word = "if_print_cell_val" Then
            If i + 6 <= col.Count Then
                If col(i + 2) = "=" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = ">" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value > Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value < Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<>" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value <> Evaluate(col(i + 3)) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                End If
            End If
            i = i + 6
            
        ElseIf word = "if_statement_cell_val" Then
        'if;a>1;statement
            'MsgBox (col(i + 5) & "BGOUEAHOUGHOUAEH")
            If i + 5 <= col.Count Then
                If col(i + 2) = "=" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Evaluate(col(i + 3)) Then
                        parser lex(col(i + 5))
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = ">" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value > Evaluate(col(i + 3)) Then
                        parser lex(col(i + 5))
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value < Evaluate(col(i + 3)) Then
                        parser lex(col(i + 5))
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<>" Then
                    If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value <> Evaluate(col(i + 3)) Then
                        parser lex(col(i + 5))
                        'MsgBox (col(i + 5))
                    End If
                End If
            End If
            i = i + 5
            
            
        ' Handle if_print_val_cell
        ElseIf word = "if_print_val_cell" Then
            If i + 6 <= col.Count Then
                If col(i + 2) = "=" Then
                    If Evaluate(col(i + 1)) = ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = ">" Then
                    If Evaluate(col(i + 1)) > ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<" Then
                    If Evaluate(col(i + 1)) < ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                ElseIf col(i + 2) = "<>" Then
                    If Evaluate(col(i + 1)) <> ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                        ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = col(i + 6)
                        'MsgBox (col(i + 5))
                    End If
                End If
            End If
            i = i + 6
            
            'if..;a1;>;10;print;a1;10;else;print;a1;20
        ElseIf word = "statement" Then
            If i + 1 <= col.Count Then
            parser lex(col(i + 1))
            End If
            i = i + 1
        ElseIf word = "if_else_print_cell_val" Then
        
        If i + 10 <= col.Count Then
            If col(i + 2) = "=" Then
            
                If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
                
                
            ElseIf col(i + 2) = ">" Then
            
                If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value > Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
                
            ElseIf col(i + 2) = "<" Then
                If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value < Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
            ElseIf col(i + 2) = "<>" Then
            
            If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value <> Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
            
            End If
        
        End If
        i = i + 10
        
        ElseIf word = "if_else_print_val_cell" Then
        
        If i + 10 <= col.Count Then
            If col(i + 2) = "=" Then
            
                If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
                
                
            ElseIf col(i + 2) = ">" Then
            
                If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value > Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
                
            ElseIf col(i + 2) = "<" Then
                If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value < Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
            ElseIf col(i + 2) = "<>" Then
            
            If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value <> Evaluate(col(i + 3)) Then
                    ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
                Else
                    ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
                End If
            
            End If
        
        End If
        i = i + 10
        'if..;a1;>;a2;print;a1;10;else;print;a1;20;
        ElseIf word = "if_else_print_cells" Then
        
    If i + 10 <= col.Count Then
        If col(i + 2) = "=" Then
            
            If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        ElseIf col(i + 2) = ">" Then
            
            If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value > ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        ElseIf col(i + 2) = "<" Then
            If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value < ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        ElseIf col(i + 2) = "<>" Then
            
            If ThisWorkbook.Sheets(1).Range(col(i + 1)).Value <> ThisWorkbook.Sheets(1).Range(col(i + 3)).Value Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        End If
    
    End If
    i = i + 10
        
        ElseIf word = "if_else_print_vals" Then
        
    If i + 10 <= col.Count Then
        If col(i + 2) = "=" Then
            
            If Evaluate(col(i + 1)) = Evaluate(col(i + 3)) Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        ElseIf col(i + 2) = ">" Then
            
            If Evaluate(col(i + 1)) > Evaluate(col(i + 3)) Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        ElseIf col(i + 2) = "<" Then
            If Evaluate(col(i + 1)) < Evaluate(col(i + 3)) Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        ElseIf col(i + 2) = "<>" Then
            
            If Evaluate(col(i + 1)) <> Evaluate(col(i + 3)) Then
                ThisWorkbook.Sheets(1).Range(col(i + 5)).Value = Evaluate(col(i + 6))
            Else
                ThisWorkbook.Sheets(1).Range(col(i + 9)).Value = Evaluate(col(i + 10))
            End If
            
        End If
    
    End If
    i = i + 10

            
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

            If i + 1 <= col.Count Then
                If col(i + 1) = "cell" Then
                    If i + 1 + 6 <= col.Count Then
                        If col(i + 2) = "in" And col(i + 4) = "print" And col(i + 5) = "cell" Then
                            Set rng = ThisWorkbook.Sheets(1).Range(col(i + 3))
                            For Each cell In rng
                                cell.Value = cell.Value + CInt(col(i + 7))
                            Next cell
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
            If i + 2 <= col.Count Then
            
                Dim res As Variant
                On Error Resume Next
                res = Evaluate(col(i + 2))
                'MsgBox col(i + 2)
                If res = "" Then
                res = col(i + 2)
                End If
                
                If res = Err.Number Then res = col(i + 2)
                ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = res
                
                i = i + 2
                
                
            
            Else
                MsgBox "Not enough arguments after 'print'."
            End If
            
        ' Handle clear_cell
        ElseIf word = "clear_cell" Then
            If i + 1 <= col.Count Then
                ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = ""
                i = i + 1
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
                For Each cell In ThisWorkbook.Sheets(1).Range(col(i + 1) & ":" & col(i + 2))
                    cell.Value = CInt(col(i + 3)) ' Increment the cell value
                    'CInt(col(i + 3)) = CInt(col(i + 3)) + CInt(col(i + 3)) ' Update increment
                Next cell
                i = i + 3
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
            
        'deappend;A1;old;new;
        ElseIf word = "replace_val" Then
        If i + 3 <= col.Count Then
            If col(i + 3) = "" Then
            ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Replace(ThisWorkbook.Sheets(1).Range(col(i + 1)).Value, CStr(col(i + 2)), "")
            
            Else
            ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Replace(ThisWorkbook.Sheets(1).Range(col(i + 1)).Value, CStr(col(i + 2)), CStr(col(i + 3)))
            
            End If
            
            
            
        End If
        i = i + 3
        
        
        ElseIf word = "append_val" Then
        If i + 2 <= col.Count Then
            ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = ThisWorkbook.Sheets(1).Range(col(i + 1)).Value & col(i + 2)
        End If
        i = i + 2
        
        ' Handle input
        'input;prompt;cell;
ElseIf word = "input" Then
    If i + 2 <= col.Count Then
        Dim userInput As String
        userInput = InputBox(col(i + 1)) ' Display the prompt
        ThisWorkbook.Sheets(1).Range(col(i + 2)).Value = userInput ' Store input in specified cell
        i = i + 2 ' Move index forward
    End If

            ElseIf word = "print_len_cell" Then
        'len;cell;destination
        If i + 2 <= col.Count Then
        ThisWorkbook.Sheets(1).Range(col(i + 2)).Value = Len(ThisWorkbook.Sheets(1).Range(col(i + 1)).Value)
        
        
        End If
        i = i + 2
        'print;A1;A2;
        ElseIf word = "print_update_len_cell" Then
        If i + 2 <= col.Count Then
        ThisWorkbook.Sheets(1).Range(col(i + 2)).Value = CStr("=LEN(" & ThisWorkbook.Sheets(1).Range(col(i + 1)).Value & ")")
        End If
        
        i = i + 2
        
        ' Handle print_substr
        'pr....;A1;A2;1;2;
        ElseIf word = "print_substr" Then
            If i + 4 <= col.Count Then
            
                'MsgBox Mid(ThisWorkbook.Sheets(1).Range(col(i + 1)).Value, CInt(col(i + 3)), CInt(col(i + 4))) & "8"
                ThisWorkbook.Sheets(1).Range(col(i + 2)).Value = Mid(ThisWorkbook.Sheets(1).Range(col(i + 1)).Value, CInt(col(i + 3)), CInt(col(i + 4)))
                i = i + 4
            End If
        
        ' Handle print_concat
        'print..;A1;A2;Destination;
        ElseIf word = "print_concat" Then
        
            If i + 3 <= col.Count Then
                
                ThisWorkbook.Sheets(1).Range(col(i + 3)).Value = ThisWorkbook.Sheets(1).Range(col(i + 1)).Value & ThisWorkbook.Sheets(1).Range(col(i + 2)).Value
                
                i = i + 3
            End If
        
        ' Handle print_replace_substr
        ElseIf word = "print_replace_substr" Then
            If i + 3 <= col.Count Then
                MsgBox Replace(ThisWorkbook.Sheets(1).Range(col(i + 1)).Value, col(i + 2), col(i + 3))
                i = i + 3
            End If
        
        ' Handle comment
        ElseIf word = "comment" Then
            If i + 1 <= col.Count Then
                i = i + 1 ' Increment i if needed
            End If
            'mask =expr;newfunc
            ElseIf word = "mask_expr" Then
            
            If i + 2 <= col.Count Then
    Dim str_1 As String
    str_1 = col(i + 2) & " = Evaluate(""" & col(i + 1) & """)"

    Dim startLine As Long
    On Error Resume Next ' Ignore errors temporarily
    startLine = ThisWorkbook.VBProject.VBComponents("Module1").CodeModule.ProcStartLine(col(i + 2), vbext_pk_Proc)
    On Error GoTo 0 ' Turn error handling back on

    ' Check if startLine is greater than 0 (i.e., the procedure exists)
    If startLine > 0 Then
        ThisWorkbook.VBProject.VBComponents("Module1").CodeModule.DeleteLines startLine, _
            startLine + ThisWorkbook.VBProject.VBComponents("Module1").CodeModule.ProcCountLines(col(i + 2), vbext_pk_Proc) - 1
    End If

    ' Add the new function
    ThisWorkbook.VBProject.VBComponents("Module1").CodeModule.AddFromString "Function " & col(i + 2) & "()" & vbCrLf & "Application.Volatile" & vbCrLf & str_1 & vbCrLf & " End Function"

End If

i = i + 2

        'align;a1;left;
        ElseIf word = "align" Then
    If i + 2 <= col.Count Then
         
        Select Case col(i + 2)
            Case "left"
                ThisWorkbook.Sheets(1).Range(col(i + 1)).HorizontalAlignment = xlLeft
            Case "center"
                ThisWorkbook.Sheets(1).Range(col(i + 1)).HorizontalAlignment = xlCenter
            Case "right"
                ThisWorkbook.Sheets(1).Range(col(i + 1)).HorizontalAlignment = xlRight
        End Select
    End If
    i = i + 2
        ElseIf word = "highlight_bg_cell_color" Then
        If i + 4 <= col.Count Then
        ' Set the background color of the specified cell
        ThisWorkbook.Sheets(1).Range(col(i + 1)).Interior.Color = RGB(CInt(col(i + 2)), CInt(col(i + 3)), CInt(col(i + 4))) ' RGB values
        i = i + 4 ' Move index forward for RGB values
    End If
    '
        ElseIf word = "highlight_fg_cell_color" Then
    If i + 4 <= col.Count Then
        ' Set the font color of the specified cell
        ThisWorkbook.Sheets(1).Range(col(i + 1)).Font.Color = RGB(CInt(col(i + 2)), CInt(col(i + 3)), CInt(col(i + 4))) ' RGB values
        i = i + 4 ' Move index forward for RGB values
    End If
        'merge;range;
        ElseIf word = "merge" Then
        If i + 2 <= col.Count Then
        
        ThisWorkbook.Sheets(1).Range(col(i + 1)).Merge
       ' ActiveSheet.Range("A1:C1").Merge: ActiveSheet.Range("A1").Value = "Merged Cell"
       i = i + 2
       End If
       
       

        ElseIf word = "unmerge" Then
        If i + 1 <= col.Count Then
        ActiveSheet.Range(col(i + 1)).UnMerge

        i = i + 1
        
        End If
        ElseIf word = "add_hyperlink" Then
        If i + 2 <= col.Count Then
        
        ThisWorkbook.Sheets(1).Hyperlinks.Add Anchor:=ThisWorkbook.Sheets(1).Range(col(i + 1)), Address:=col(i + 2), TextToDisplay:="Click here"
        i = i + 2
        End If
        'lambda;funcname;expr;paramaters;
        ElseIf word = "lambda" Then
        If i + 3 <= col.Count Then
        ThisWorkbook.VBProject.VBComponents("Module1").CodeModule.AddFromString "Function " & col(i + 1) & "(" & col(i + 2) & ")" & vbCrLf & "Application.Volatile" & vbCrLf & col(i + 3) & vbCrLf & " End Function"
        
        i = i + 3
        
        End If
        
        ElseIf word = "format" Then
        If i + 1 <= col.Count Then
        
        If col(i + 1) = "bold" Then
        If i + 2 <= col.Count Then
        
        ThisWorkbook.Sheets(1).Range(col(i + 2)).Font.Bold = True
        i = i + 2
        End If
        
        
        ElseIf col(i + 1) = "number" Then
        
        If i + 3 <= col.Count Then
        
        ThisWorkbook.Sheets(1).Range(col(i + 2)).NumberFormat = CStr(col(i + 3))
        i = i + 3
        
        End If
        End If
        End If
        
        
        ElseIf word = "check_cell_value_bool" Then
        Dim check_var As Variant
        
        If i + 2 <= col.Count Then
        check_var = ThisWorkbook.Sheets(1).Range(col(i + 1)).Value
        MsgBox VarType(check_var) = vbInterger
        
        If VarType(ThisWorkbook.Sheets(1).Range(col(i + 1)).Value) = vbDouble Then
        MsgBox ("hey")
        End If
        
        
        
        i = i + 2
        End If
        
        ' Handle check_cell_value
ElseIf word = "check_cell_value" Then
    If i + 1 <= col.Count Then
        MsgBox ThisWorkbook.Sheets(1).Range(col(i + 1)).Value
        i = i + 1
    End If
        'print;var;a;
        ElseIf word = "print_var" Then
            If i + 2 <= col.Count Then
                If variableTable.Count > 0 Then
            For j = 1 To variableTable.Count
                ' Error checking: Ensure we're within bounds
                
                If j <= variableTable.Count Then
                    
                    If variableTable(j) = col(i + 2) Then
                        ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = variableTable(j + 1)
                        
                        
                        Exit For
                    End If
                End If
            Next j
        End If
        
            End If
            i = i + 2
        ElseIf word = "define_var" Then
    If i + 2 <= col.Count Then
        ' Check if variableTable has items before accessing
        If variableTable.Count > 0 Then
            For j = 1 To variableTable.Count
                ' Error checking: Ensure we're within bounds
                
                If j <= variableTable.Count Then
                    
                    If variableTable(j) = col(i + 1) Then
                        ' Remove variable and its value
                        variableTable.Remove j
                        variableTable.Remove j
                        
                        Exit For
                    End If
                End If
            Next j
        End If
        Dim res1 As Variant
                On Error Resume Next
                res1 = Evaluate(col(i + 2))
                'MsgBox col(i + 2)
                If res1 = "" Then
                res1 = col(i + 2)
                End If
                
                If res = Err.Number Then res = col(i + 2)
                
        ' Add the new variable name and value
        variableTable.Add col(i + 1)  ' Variable name
        
        variableTable.Add res1  ' Variable value
        
    End If
    i = i + 2

        ElseIf word = "get_var" Then
            If i + 1 <= col.Count Then
            
            MsgBox variableTable.Count
            For j = 1 To variableTable.Count
                
                If variableTable(j) = col(i + 1) Then
                MsgBox variableTable(j + 1)
                Exit For
                
                
                
                End If
                
                
                j = j + 1
            Next j
          i = i + 1
            
        End If
        
         ElseIf word = "form" Then
         Dim f As UserForm
        Dim btn As MSForms.CommandButton

    ' Create an instance of the UserForm
    
         'form;create;name;
         If i + 2 <= col.Count Then
            
         End If
         
    
    ' Use the name of your UserForm
         ' Clear existing controls (optional)
    
    Set btn = f.Controls.Add("Forms.CommandButton.1", "btnNew")
    With btn
        .Caption = "Click Me"
        .Left = 50
        .Top = 50
        .Width = 100
    End With

    ' Show the modified UserForm
    UserForm1.Show
    
    
         
        ElseIf word = "save" Then
    Dim newWorkbook As Workbook
    Dim sheet As Worksheet

    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    Set sheet = newWorkbook.Sheets(1)

    ' Write data to cell A1
    sheet.Cells(1, 1).Value = "This is some sample text."

    ' Save the new workbook
    newWorkbook.SaveAs "C:\path\to\your\file.xlsx" ' Change this to your desired file path
    newWorkbook.Close

    MsgBox "Data saved to Excel file."
        'switch;var;case10;statement;end switch;
        ElseIf word = "switch" Then
        Dim switch As Collection
        Set switch = New Collection
        Dim str As Variant
        Dim x As Integer
        Dim y As Integer
        
        
        For y = i + 2 To col.Count
        
        If str = "switch" Then
        ElseIf col(y) = "end switch" Then
        Exit For
        Else
        switch.Add col(y)
        
        End If
  
        Next y
        For j = 1 To variableTable.Count
                
                If variableTable(j) = col(i + 1) Then
               str = variableTable(j + 1)
                Exit For
                
                
                
                End If
                
                
                j = j + 1
            Next j
            
         For x = 1 To switch.Count
        MsgBox switch(x + 1) & "h" & str
        If CStr(switch(x)) = str Then
        MsgBox switch(x) & " h" & str
        parser lex(switch(x + 2))
        End If
        x = x + 2
        'MsgBox str
        Next x
        i = i + switch.Count + 1
        
        'function;name;args;statement;end_function;
        'ARR_COL;name;args;statement;statement;
        ElseIf word = "function" Then
        If i + 4 <= col.Count Then
        arr_col.Add col(i + 1)
        arr_col.Add col(i + 2)
        'col(i+3) is statement then the enxt one is the []
        arr_col.Add col(i + 4)
        End If
        i = i + 4
        
        ElseIf word = "call_function" Then
        Dim z As Integer
        Dim zz As Integer
        
        
        If i + 2 <= col.Count Then
        For z = 1 To arr_col.Count
        If col(i + 1) = arr_col(z) Then
        Dim s As Variant
        Dim ss As String
        Dim sss As Variant
        
        
        Dim new_col As Collection
        Set new_col = New Collection
        Dim new_col1 As Collection
        Set new_col1 = New Collection
        '(10,20)
        
        sss = col(i + 2)
        sss = Replace(sss, "(", "")
        sss = Replace(sss, ")", "")
        sss = Split(sss, ",")
        For zz = LBound(sss) To UBound(sss)
            new_col1.Add Trim(sss(zz)) ' Trim spaces around elements
        Next zz
        '(a,b)
        s = arr_col(z + 1)
        
        s = Replace(s, "(", "")
        s = Replace(s, ")", "")
        s = Split(s, ",")
        
        For zz = LBound(s) To UBound(s)
            new_col.Add Trim(s(zz)) ' Trim spaces around elements
        Next zz
        
        For zz = 1 To new_col.Count
         ss = CStr(ss) & "define_var;" & CStr(new_col(zz)) & ";" & CStr(new_col1(zz)) & ";"
         
        Next zz
        
        
        ss = ss & Mid(arr_col(z + 2), 1, Len(arr_col(z + 2)))
        parser lex(ss)
        
        Else
        z = z + 2
        
        End If
        
        Next z
        
        
        i = i + 2
        
        End If
        'load;file;
        ElseIf word = "load_run" Then
        If i + 1 <= col.Count Then
        Dim fileContent As String
        
        fileContent = LoadFileAsString(col(i + 1))
        parser lex(fileContent)
        
        
        End If
        ElseIf word = "py" Then
    Dim result As String
    Dim scriptPath As String
    
    scriptPath = "C:\Users\elrai\vba_interperter_code\a.py" ' Adjust as needed
    result = RunPythonScript(scriptPath)
    
    ' If you're reading the output, you can display it
    'MsgBox result
    
    ElseIf word = "random" Then
             
    If i + 4 <= col.Count Then
    
    
    ' Specify the path for the file
    filePath = "C:\Users\elrai\vba_interperter_code\a.py" ' Change this to your desired path
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write data to the Python script
    Print #fileNum, "import random"
    Print #fileNum, "r = random.randint(" & col(i + 1) & "," & col(i + 2) & ")" ' Corrected the assignment with spaces around "="
    Print #fileNum, "with open('C:\\Users\\elrai\\vba_interperter_code\\output.txt', 'w') as file:" ' Escaped backslashes
    Print #fileNum, vbTab & "file.write(str(r))  # Write the random integer to the file" ' Corrected to write r as a string
    Print #fileNum, "print(r)"
    
    ' Close the file
    Close #fileNum
    
        result = RunPythonScript("C:\Users\elrai\vba_interperter_code\a.py")
        ThisWorkbook.Sheets(1).Range(col(i + 4)).Value = result
        i = i + 4
        End If
        
    ElseIf word = "py_input" Then
    'py_input;a1;prompt;
    If i + 2 <= col.Count Then
    ' Specify the path for the file
    filePath = "C:\Users\elrai\vba_interperter_code\a.py" ' Change this to your desired path
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write data to the Python script
    Print #fileNum, "X = input('" & col(i + 2) & "')"
    Print #fileNum, "with open('C:\\Users\\elrai\\vba_interperter_code\\output.txt', 'w') as file:" ' Escaped backslashes
    Print #fileNum, vbTab & "file.write(X)" ' Corrected to write r as a string
    Print #fileNum, "input('Press any key to exit')"
    
    ' Close the file
    Close #fileNum
    result = RunPythonScript("C:\Users\elrai\vba_interperter_code\a.py")
    ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = result
    i = i + 2
    End If
    
        ElseIf word = "py_input_eval" Then
    'py_input;a1;prompt;
    If i + 2 <= col.Count Then
    ' Specify the path for the file
    filePath = "C:\Users\elrai\vba_interperter_code\a.py" ' Change this to your desired path
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write data to the Python script
    Print #fileNum, "X = input('" & col(i + 2) & "')"
    Print #fileNum, "with open('C:\\Users\\elrai\\vba_interperter_code\\output.txt', 'w') as file:" ' Escaped backslashes
    Print #fileNum, vbTab & "file.write(X)" ' Corrected to write r as a string
    Print #fileNum, "input('Press any key to exit')"
    
    ' Close the file
    Close #fileNum
    result = RunPythonScript("C:\Users\elrai\vba_interperter_code\a.py")
    ThisWorkbook.Sheets(1).Range(col(i + 1)).Value = Evaluate(result)
    i = i + 2
    End If
    
        ElseIf word = "tk" Then
        CreateTkinterScript
        
        ElseIf word = "python" Then
        Dim com As String
        
       
        com = "while True: " & vbCrLf & _
      "    command = input('Enter a command (or ''exit'' to quit): '); " & vbCrLf & _
      "    if command.lower() == 'exit': break; " & vbCrLf & _
      "    elif command.lower() == 'no': print('hey doos'); break;" & vbCrLf & _
      "    try: exec(command);" & vbCrLf & _
      "    except Exception as e: print(f'Error: {e}');" & vbCrLf & _
      "input()"



     ' Now call the Python script
     Shell "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe -c """ & com & """", vbNormalFocus

    'creates custom tkinter form
    
     
    
    'Shell "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe " & " -c " & com, vbNormalFocus
    'Shell "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe " & "C:\Users\elrai\vba_interperter_code\output.py", vbNormalFocus
    
        'Shell "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe"
        
        ElseIf word = "eval()" Then
        'eval()
        If i + 1 <= col.Count Then
        i = i + 1
        End If
       
        ElseIf word = "create_tk" Then
    Dim params(6) As Variant ' To hold parameters: Name, Height, Width, Resizeable, Background, OnTop
    

    ' Default values
    params(0) = vbNullString ' Name
    params(1) = 200          ' Height
    params(2) = 300          ' Width
    params(3) = False        ' Resizeable
    params(4) = "white"      ' Background color
    params(5) = False        ' OnTop
    params(6) = False
    

    ' Parse arguments
    'creat...;all;name;hei;wid;resize;window_col_bg;ontop;diabled
    If i + 8 <= col.Count And col(i + 1) = "all" Then
        params(0) = col(i + 2)
        params(1) = CInt(col(i + 3))
        params(2) = CInt(col(i + 4))
        params(3) = col(i + 5)
        params(4) = col(i + 6)
        params(5) = col(i + 7)
        params(6) = col(i + 8)
        i = i + 8
    Else
        For j = 1 To 5
            If i + 2 <= col.Count Then
                Select Case col(i + 1)
                    Case "name"
                        params(0) = col(i + 2)
                        i = i + 2
                    Case "height"
                        params(1) = CInt(col(i + 2))
                        i = i + 2
                    Case "width"
                        params(2) = CInt(col(i + 2))
                        i = i + 2
                    Case "resize"
                        params(3) = True
                        i = i + 2
                    Case "window_bg"
                        params(4) = col(i + 2)
                        i = i + 2
                    Case "window_ontop"
                        params(5) = True
                        i = i + 2
                End Select
            End If
        Next j
    End If

    ' Call CreateCustomTkForm with the parameters
    Call CreateCustomTkForm(params)
        ElseIf word = "show_custom_msgbox" Then
        If i + 2 <= col.Count Then
        show_custom_msgbox col(i + 1), col(i + 2)
        i = i + 2
        End If
        
        ElseIf word = "show_c_msgbox_y_n" Then
    Dim response As String
    Dim responseFile As String
    responseFile = "C:\Users\elrai\vba_interperter_code\output.txt" ' Adjust as necessary

    ' Call the custom message box
    Call show_custom_msgbox("Confirmation", "Do you want to proceed?")

    ' Wait for the response file to be created/updated
    'Do While Dir(responseFile) = "" ' Wait until the response file exists
     '   Application.Wait Now + TimeValue("00:00:01") ' Check every second
    'Loop

    ' Read the response from the file
    
    
    response = LoadFileAsString(responseFile)
    

    ' Determine what was clicked
    If response = "Yes" Then
        MsgBox "You clicked Yes"
    ElseIf response = "No" Then
        MsgBox "You clicked No"
    End If
    fileNum = FreeFile
    ' Clear the response file after reading
    Open responseFile For Output As #fileNum
    Print #fileNum, "" ' Write an empty string to clear the file
    Close #fileNum



    ' Optionally, delete the response file after reading

        
        ElseIf word = "" Then
        
       End If
        
    
    Next i


End Function

Sub CreateCustomTkForm(Optional ByVal params As Variant)
    Dim filePath As String
    Dim pythonCode As String
    Dim pythonExe As String

    ' Path to the Python file
    filePath = "C:\Users\elrai\vba_interperter_code\a.py" ' Adjust as necessary

    ' Path to the Python executable (using pythonw.exe to hide console)
    pythonExe = "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe" ' Adjust to your Python installation path
    
    
    'params 0 is name
    'params 1 is height
    'params 2 is width
    'params 3 is resizeable
    'params 4 is winodw_bg
    'params 5 is ontop
    'params 6 is the disabled
    ' Build Python code
    pythonCode = "import tkinter as tk" & vbCrLf
    pythonCode = pythonCode & vbCrLf
    pythonCode = pythonCode & "root = tk.Tk()" & vbCrLf
    pythonCode = pythonCode & "root.title('" & params(0) & "')" & vbCrLf
    pythonCode = pythonCode & "root.geometry('" & params(2) & "x" & params(1) & "')" & vbCrLf
    
    If params(3) = True Then
        pythonCode = pythonCode & "root.resizable(width=True, height=True)" & vbCrLf
    End If
    
    If params(5) = True Then
        pythonCode = pythonCode & "root.attributes('-topmost', True)" & vbCrLf
    End If
    
    If params(6) = True Then
    pythonCode = pythonCode & "root.attributes('-disabled', True)" & vbCrLf
    
    End If
    

    pythonCode = pythonCode & "root.configure(bg='" & params(4) & "')" & vbCrLf
    pythonCode = pythonCode & "root.mainloop()" & vbCrLf

    ' Save the Python code to the file and execute
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, pythonCode
    Close #fileNum
    
    ' Execute the Python script
    Shell pythonExe & " " & filePath, vbHide
    Application.Wait Now + TimeValue("00:00:005")
End Sub


'End Sub

' Function to get the variable value from the collection
Private Function GetVariableValue(varTable As Collection, varName As String) As Variant
    If varTable Is Nothing Then
        GetVariableValue = "Variable table is not initialized"
        Exit Function
    End If
    
    On Error Resume Next
    
    GetVariableValue = varTable(varName)
    If Err.Number <> 0 Then
        GetVariableValue = varTable(varName)
        Err.Clear
    End If
    
    On Error GoTo 0
End Function

Function LoadFileAsString(filePath As String) As String
    Dim fileContent As String
    Dim fileNum As Integer
    MsgBox filePath
    ' Get the next available file number
    fileNum = FreeFile
    
    ' Open the file for input
    Open filePath For Input As #fileNum
    
    ' Read the entire content of the file
    fileContent = Input$(LOF(fileNum), fileNum)
    
    ' Close the file
    Close #fileNum
    
    ' Return the file content
    LoadFileAsString = fileContent
End Function

Function RunPythonScript(scriptPath As String, Optional pythonExe As String = "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe") As String
    ' Declare variables
    Dim command As String          ' Command string to execute the Python script
    Dim fileContent As String      ' Variable to hold the content read from the output file
    Dim filePath As String         ' Path to the output text file
    Dim wsh As Object              ' WScript.Shell object for running the command
    
    ' Construct the command to execute the Python script
    command = pythonExe & " """ & scriptPath & """"
    
    ' Create an instance of the WScript.Shell object
    Set wsh = CreateObject("WScript.Shell")
    
    ' Execute the command; parameters:
    ' 0 = hide the window during execution
    ' True = wait for the script to finish before proceeding
    wsh.Run command, 1, True
    
    ' Define the path to the output text file
    filePath = "C:\Users\elrai\vba_interperter_code\output.txt" ' Ensure this matches your Python script's output path
    
    ' Open the output file for input
    Open filePath For Input As #1
    ' Read the entire content of the file into fileContent
    fileContent = Input$(LOF(1), 1)
    ' Close the file after reading
    Close #1
        
    ' Return the content read from the output file
    RunPythonScript = fileContent
End Function


'Function RunPythonScript_nowait(scriptPath As String, Optional pythonExe As String = "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe") As String
    ' Declare variables
'    Dim command As String          ' Command string to execute the Python script
'    Dim fileContent As String      ' Variable to hold the content read from the output file
'    Dim filePath As String         ' Path to the output text file
'    Dim wsh As Object              ' WScript.Shell object for running the command
    
'    ' Command to execute the Python script
'    command = pythonExe & " """ & scriptPath & """"
    
'    ' Create a WScript.Shell object
'    Set wsh = CreateObject("WScript.Shell")
    
'    ' Execute the command and wait for it to complete
'    wsh.Run command, 0, True ' 0 = hide the window, True = wait for completion
    
'    ' Path to the output text file
'    filePath = "C:\Users\elrai\vba_interperter_code\output.txt" ' Ensure this matches your Python script's output path
    
'    ' Read the output file
'    Open filePath For Input As #1
'    fileContent = Input$(LOF(1), 1) ' Read the entire content of the file
'    Close #1 ' Close the file after reading
        
'    ' Return the content from the text file (if needed)
'    RunPythonScript = fileContent
'End Function



Sub CreateTkinterScript()
    Dim filePath As String
    Dim fileNum As Integer
    Dim pythonCode As String
    Dim pythonExe As String

    ' Path to the Python file
    filePath = "C:\Users\elrai\vba_interperter_code\a.py" ' Adjust as necessary

    ' Path to the Python executable (using pythonw.exe to hide console)
    pythonExe = "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe" ' Adjust to your Python installation path

    ' Start building the Python Tkinter code
    pythonCode = "import tkinter as tk" & vbCrLf
    pythonCode = pythonCode & "def highlight_text(event):" & vbCrLf
    pythonCode = pythonCode & vbTab & "text_content = text.get('1.0', 'end-1c')" & vbCrLf
    pythonCode = pythonCode & vbTab & "text.delete('1.0', 'end')" & vbCrLf
    pythonCode = pythonCode & vbTab & "text.insert('1.0', text_content)" & vbCrLf
    pythonCode = pythonCode & vbTab & "text.tag_remove('highlight', '1.0', 'end')" & vbCrLf
    pythonCode = pythonCode & vbTab & "text.tag_remove('msg_highlight', '1.0', 'end')" & vbCrLf
    
    ' Lists of words to highlight
    pythonCode = pythonCode & vbTab & "red_words = ['print;','if_else_cell_val;']" & vbCrLf
    pythonCode = pythonCode & vbTab & "blue_words = ['msg;']" & vbCrLf

    ' Highlighting words in the lists
    pythonCode = pythonCode & vbTab & "for word in red_words:" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & "start = '1.0'" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & "while True:" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "start = text.search(word, start, stopindex='end')" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "if not start:" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & vbTab & "break" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "end = f'{start}+{len(word)}c'" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "text.tag_add('highlight', start, end)" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "start = end" & vbCrLf

    pythonCode = pythonCode & vbTab & "for word in blue_words:" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & "start = '1.0'" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & "while True:" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "start = text.search(word, start, stopindex='end')" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "if not start:" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & vbTab & "break" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "end = f'{start}+{len(word)}c'" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "text.tag_add('msg_highlight', start, end)" & vbCrLf
    pythonCode = pythonCode & vbTab & vbTab & vbTab & "start = end" & vbCrLf

    pythonCode = pythonCode & vbCrLf
    pythonCode = pythonCode & "root = tk.Tk()" & vbCrLf
    pythonCode = pythonCode & "root.title('Highlight Text')" & vbCrLf

    ' Create a Text widget and bind the key release event
    pythonCode = pythonCode & "text = tk.Text(root, font=('Arial', 14), wrap='word')" & vbCrLf
    pythonCode = pythonCode & "text.pack(pady=20)" & vbCrLf
    pythonCode = pythonCode & "text.tag_configure('highlight', foreground='red')" & vbCrLf
    pythonCode = pythonCode & "text.tag_configure('msg_highlight', foreground='blue')" & vbCrLf
    pythonCode = pythonCode & "text.bind('<KeyRelease>', highlight_text)" & vbCrLf

    ' Run the Tkinter main loop
    pythonCode = pythonCode & "root.mainloop()" & vbCrLf

    ' Open the Python file for output and write the code
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, pythonCode
    Close #fileNum

    ' Run the Python script using pythonw.exe to hide the console
    Shell pythonExe & " " & filePath, vbHide
End Sub




