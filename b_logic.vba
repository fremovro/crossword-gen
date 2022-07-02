'Генерация кроссвордов
Sub generate_crosswords(words As tWords, ByRef crosswords As tCrosswords, ByRef err_check As Boolean)
    Dim intersections As tIntersections
    intersections.count = 0
    '2.1) Заполнить пересечения у всех слов
    Call fill_intersection(words, intersections)
    '2.2) Сгенерировать все возможные варианты
    Call generate(words, crosswords, err_check, intersections)
    '2.3) Отобрать все заполненные кроссворды
    Call selection(crosswords, err_check)
End Sub
Sub generate(words As tWords, ByRef crosswords As tCrosswords, ByRef err_check As Boolean, intersections As tIntersections)
    Dim len_arr(), temp_arr() As Variant, _
        f As Boolean, _
        c As Integer, t As Integer
    len_arr = Array(11, 9, 9, 7, 7, 7, 6)
    t = 0
    ReDim crosswords.item(0)
    For i = 0 To 6
        c = crosswords.count
        For j = 0 To c
            f = True
            t = 0
            For k = 0 To words.count - 1
                Call drop_used(words, crosswords, intersections, CInt(len_arr(i)), CInt(i), CInt(j)) 'Сбросить метки использования, если все подходящие слова использованы
                If (Len(words.item(k).content) = len_arr(i)) Then 'Проверка совместимости по длине
                    If (let_check(words.item(k), crosswords.item(j), intersections, CInt(i))) Then 'Проверка совместимости по пересечениям
                        If (words.item(k).used = False) Then 'Если слово не использовалось, то оно сразу же используется
                            Call fill_crossword(crosswords, words.item(k), CInt(j), CInt(i), f)
                            words.item(k).used = True 'Отмечается как использованное (реализация пункта 3.6)
                        Else 'Иначе откладывается на потом(запоминается индекс) (реализация пункта 3.6)
                            ReDim temp_arr(t)
                            temp_arr(t) = k
                            t = t + 1
                        End If
                    End If
                End If
            Next
            For k = 0 To t - 1 'Использование всех подходящих слов, отложенных на потом (реализация пункта 3.6)
                Call fill_crossword(crosswords, words.item(temp_arr(k)), CInt(j), CInt(i), f)
            Next
        Next
    Next
End Sub
Sub fill_crossword(ByRef crosswords As tCrosswords, word As tWord, j As Integer, i As Integer, ByRef f As Boolean)
    If (f) Then
        ReDim Preserve crosswords.item(j).content.item(i)
        crosswords.item(j).content.item(i) = word
        crosswords.item(j).content.count = crosswords.item(j).content.count + 1
        f = False
    Else
        crosswords.count = crosswords.count + 1
        ReDim Preserve crosswords.item(crosswords.count)
        If i <> 0 Then
            crosswords.item(crosswords.count) = crosswords.item(j)
            crosswords.item(crosswords.count).content.count = crosswords.item(crosswords.count).content.count - 1
        End If
        ReDim Preserve crosswords.item(crosswords.count).content.item(i)
        crosswords.item(crosswords.count).content.item(i) = word
        crosswords.item(crosswords.count).content.count = crosswords.item(crosswords.count).content.count + 1
    End If
End Sub
Sub drop_used(ByRef words As tWords, crosswords As tCrosswords, intersections As tIntersections, my_len As Integer, i As Integer, j As Integer)
    Dim f As Boolean
    f = True
    For k = 0 To words.count - 1
        If (Len(words.item(k).content) = my_len) Then
            If (let_check(words.item(k), crosswords.item(j), intersections, CInt(i))) Then
                If (words.item(k).used = False) Then
                    f = False
                End If
            End If
        End If
    Next
    If (f) Then
        For k = 0 To words.count - 1
            If (Len(words.item(k).content) = my_len) Then
                If (let_check(words.item(k), crosswords.item(j), intersections, CInt(i))) Then
                    words.item(k).used = False
                End If
            End If
        Next
    End If
End Sub
Function let_check(word As tWord, crossword As tCrossword, intersections As tIntersections, i As Integer) As Boolean
    Dim f_1, f_2, f_3 As Boolean
    f_1 = False
    f_2 = False
    f_3 = False
    If (i > crossword.content.count) Then GoTo f_res
    Select Case (i)
    Case 0
        GoTo t_res
    Case 1
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(0).content And _
            intersections.item(j).pos_1 = 8 And intersections.item(j).pos_2 = 2) Then GoTo t_res
        Next
    Case 2
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(1).content And _
            intersections.item(j).pos_1 = 5 And intersections.item(j).pos_2 = 5) Then GoTo t_res
        Next
    Case 3
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(0).content And _
            intersections.item(j).pos_1 = 5 And intersections.item(j).pos_2 = 1) Then
                f_1 = True
            End If
        Next
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(2).content And _
            intersections.item(j).pos_1 = 2 And intersections.item(j).pos_2 = 4) Then
                f_2 = True
            End If
        Next
    If (f_1 And f_2) Then GoTo t_res
    Case 4
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(0).content And _
            intersections.item(j).pos_1 = 11 And intersections.item(j).pos_2 = 1) Then
                f_1 = True
            End If
        Next
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(2).content And _
            intersections.item(j).pos_1 = 8 And intersections.item(j).pos_2 = 4) Then
                f_2 = True
            End If
        Next
        If (f_1 And f_2) Then GoTo t_res
    Case 5
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(3).content And _
            intersections.item(j).pos_1 = 7 And intersections.item(j).pos_2 = 1) Then
                f_1 = True
            End If
        Next
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(1).content And _
            intersections.item(j).pos_1 = 8 And intersections.item(j).pos_2 = 4) Then
                f_2 = True
            End If
        Next
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(4).content And _
            intersections.item(j).pos_1 = 7 And intersections.item(j).pos_2 = 7) Then
                f_3 = True
            End If
        Next
        If (f_1 And f_2 And f_3) Then GoTo t_res
    Case 6
        For j = 0 To intersections.count - 1
            If (intersections.item(j).i_2.content = word.content And intersections.item(j).i_1.content = crossword.content.item(0).content And _
            intersections.item(j).pos_1 = 6 And intersections.item(j).pos_2 = 5) Then GoTo t_res
        Next
    End Select
    GoTo f_res
t_res:
    let_check = True
    Exit Function
f_res:
    let_check = False
End Function
Sub selection(ByRef crosswords As tCrosswords, err_check As Boolean)
    Dim temp As tCrosswords, _
        c As Integer
    c = 0
    For i = 0 To crosswords.count
        If (crosswords.item(i).content.count = 7) Then
            ReDim Preserve temp.item(c)
            temp.item(c) = crosswords.item(i)
            c = c + 1
        End If
    Next
    crosswords = temp
    crosswords.count = c
End Sub
Sub fill_intersection(words As tWords, ByRef intersections As tIntersections)
    Dim temp_word As tWord
    For i = 0 To words.count - 1
        temp_word = words.item(i)
        For j = 0 To words.count - 1
            Select Case Len(temp_word.content)
            Case 11
                Select Case Len(words.item(j).content)
                Case 9
                    If (Mid(temp_word.content, 8, 1) = Mid(words.item(j).content, 2, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 8, 2)
                    End If
                Case 7
                    If (Mid(temp_word.content, 5, 1) = Mid(words.item(j).content, 1, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 5, 1)
                    End If
                    If (Mid(temp_word.content, 11, 1) = Mid(words.item(j).content, 1, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 11, 1)
                    End If
                Case 6
                    If (Mid(temp_word.content, 6, 1) = Mid(words.item(j).content, 5, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 6, 5)
                    End If
                End Select
            Case 9
                Select Case Len(words.item(j).content)
                Case 9
                    If (Mid(temp_word.content, 5, 1) = Mid(words.item(j).content, 5, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 5, 5)
                    End If
                Case 7
                    If (Mid(temp_word.content, 2, 1) = Mid(words.item(j).content, 4, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 2, 4)
                    End If
                    If (Mid(temp_word.content, 8, 1) = Mid(words.item(j).content, 4, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 8, 4)
                    End If
                End Select
            Case 7
                If (Len(words.item(j).content) = 7) Then
                    If (Mid(temp_word.content, 7, 1) = Mid(words.item(j).content, 1, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 7, 1)
                    End If
                    If (Mid(temp_word.content, 7, 1) = Mid(words.item(j).content, 7, 1)) Then
                        Call i_add(intersections, temp_word, words.item(j), 7, 7)
                    End If
                End If
            End Select
        Next
    Next
End Sub
Sub i_add(ByRef intersections As tIntersections, temp_word As tWord, word As tWord, a As Integer, b As Integer)
    ReDim Preserve intersections.item(intersections.count)
    intersections.item(intersections.count).i_1 = temp_word
    intersections.item(intersections.count).i_2 = word
    intersections.item(intersections.count).pos_1 = a
    intersections.item(intersections.count).pos_2 = b
    intersections.count = intersections.count + 1
End Sub