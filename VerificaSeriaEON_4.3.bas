Attribute VB_Name = "Module1"
'Verificare Seria Versiune: 4.3'
'Data: 13 iulie 2022'

Sub VerificareSerie()
    Dim FindString As String
    Dim Rng As Range
    Dim lastColumn As Long
    Dim s As String
    Dim container As String
    Dim an As String
    Dim anUser As String
    Dim colSerieCorecta As Integer
    Dim colContainer As Integer
    Dim colName As String
    Dim colSerie As Integer
    Dim colAssesmentCode As Integer
    Dim colTipMontaj As Integer
    Dim colIndexA As Integer
    Dim colIndexR As Integer
    Dim colEchipament As Integer
    Dim colAn As Integer
    Dim fileEON As String
    Dim ziua As String
    Dim IndexA As String
    Dim IndexR As String
    Dim path As String
    Dim screi As Boolean
    Dim SArray() As String
    Dim poz As Integer
    
        
    fisa = InputBox("Seriile care se cauta se gasesc pe fisa", , "Sheet1")
    If StrPtr(fisa) = 0 Then
        Exit Sub
    End If
        
    container = InputBox("Numar container", , "0")
    If StrPtr(fisa) = 0 Then
        Exit Sub
    End If
        
    anUser = InputBox("Anul (pentru serii fara an inclus)", , "0")
     If StrPtr(anUser) = 0 Then
        Exit Sub
    End If
    
    colSerieCorecta = 0
    colContainer = 0
    colAssesmentCode = 0
    colTipMontaj = 0
    colIndex = 0
    lastColumn = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    For i = 1 To lastColumn
        If UCase(Cells(1, i).Value) = "SERIE" Then colSerie = i
        If UCase(Cells(1, i).Value) = "DESCRIERE" Then colAssesmentCode = i
        If UCase(Cells(1, i).Value) = "TIP MONTAJ" Then colTipMontaj = i
        If UCase(Cells(1, i).Value) = "INDEX DEMONTARE ACTIV" Then colIndexA = i
        If UCase(Cells(1, i).Value) = "INDEX DEMONTARE REACTIV" Then colIndexR = i
        If UCase(Cells(1, i).Value) = "COD ECHIPAMENT" Then colEchipament = i
        If UCase(Cells(1, i).Value) = "AN DE FABRICATIE" Then colAn = i
        If Cells(1, i).Value = "Container" Then colContainer = i
        If Cells(1, i).Value = "Serie corecta" Then colSerieCorecta = i
    Next i
    If colSerie = 0 Then
        MsgBox "Lipsa coloana Serie"
        Exit Sub
    End If
    If colAssesmentCode = 0 Then
        MsgBox "Lipsa coloana Assesment Code"
        Exit Sub
    End If
    If colTipMontaj = 0 Then
        MsgBox "Lipsa coloana Tip Montaj"
        Exit Sub
    End If
    If colEchipament = 0 Then
        MsgBox "Lipsa coloana Cod echipament"
        Exit Sub
    End If
    If colIndexA = 0 Then
        MsgBox "Lipsa coloana Index Activ. Indexul va fi NULL"
    End If
    If colIndexR = 0 Then
        MsgBox "Lipsa coloana Index Reactiv. Indexul va fi NULL"
    End If
    
    If colSerieCorecta = 0 Then
        colSerieCorecta = lastColumn + 1
        colContainer = lastColumn + 2
        Cells(1, colSerieCorecta) = "Serie corecta"
        Cells(1, colContainer) = "Container"
    End If
        
    'parametri nume fisier
    path = ThisWorkbook.path & "\"
    ziua = DAY(Date) & MONTH(Date) & YEAR(Date)
    'deschid fisierele in care se vor scrie seriile
    fileEON = path & "EON_" & ziua & ".csv"
    Open fileEON For Append As #1
        
    colName = Split(Cells(1, colSerie).Address, "$")(1) 'numele coloanei unde sunt seriile
    Let CopyRange = colName & Startrow & ":" & colName & Lastrow
    colNumber = Range(colName & 1).Column
        
    FindString = " "
    While FindString <> vbNullString
        FindString = InputBox("Citeste seria")
        If (FindString = vbNullString) Then
            Result = MsgBox("Terminat?", vbYesNo + vbQuestion)
            If Result = vbYes Then
                FindString = ""
            Else
                container = InputBox("Numar container", , "0")
                anUser = InputBox("Anul (pentru serii fara an inclus)", , "0")
                FindString = " "
            End If
        End If
        If StrPtr(FindString) = 0 Then
            Exit Sub
        End If
        FindString = Replace(FindString, "|", "")
        If Trim(FindString) <> "" Then
            SArray = Split(FindString)
            If (UBound(SArray) - LBound(SArray) + 1) = 3 Then 'trei parti
                Serie = Val(SArray(2))
                an = SArray(1)
                If CInt(an) < 50 Then
                    an = "20" + an 'an de 2 cifre completat la 4 in An fabricatie
                Else
                    an = "19" + an
                End If
            ElseIf Len(FindString) = 16 And Mid(FindString, 1, 4) = "1001" Then       'citeste 10 caractere + an
                an = Mid(FindString, 5, 2)
                If CInt(an) < 50 Then
                    an = "20" + an 'an de 2 cifre completat la 4 in An fabricatie
                Else
                    an = "19" + an
                End If
                Serie = Mid(FindString, 7, 10)
                Serie = CLng(Serie)
            ElseIf Len(FindString) = 16 And Mid(FindString, 1, 4) = "1002" Then   'citeste 10 caractere, sterge 0 din fata + an
                an = Mid(FindString, 5, 2)
                If CInt(an) < 50 Then
                    an = "20" + an 'an de 2 cifre completat la 4 in An fabricatie
                Else
                    an = "19" + an
                End If
                Serie = Mid(FindString, 7, 10)
                Serie = CLng(Serie)
            ElseIf Len(FindString) = 16 And Mid(FindString, 1, 4) = "1009" Then   'citeste tot
                an = Mid(FindString, 5, 2)
                If CInt(an) < 50 Then
                    an = "20" + an 'an de 2 cifre completat la 4 in An fabricatie
                Else
                    an = "19" + an
                End If
                Serie = FindString
            ElseIf Mid(FindString, 1, 3) = "101" Then
                an = Mid(FindString, 4, 2)
                If CInt(an) < 50 Then
                    an = "20" + an 'an de 2 cifre completat la 4 in An fabricatie
                Else
                    an = "19" + an
                End If
                Serie = FindString
            ElseIf InStr(FindString, "/") > 0 Then 'serie cu an
                poz = InStr(FindString, "/")
                Serie = Val(Mid(FindString, 1, poz - 1))
                an = Mid(FindString, poz + 1, Len(FindString))
            Else              'alt tip serie
                Serie = Val(FindString) 'sterge 0 din fata daca exista
                an = anUser
            End If
            
            With Sheets(fisa).Range(CopyRange)
                Set Rng = .Find(What:=Serie, _
                                After:=.Cells(.Cells.Count), _
                                LookIn:=xlValues, _
                                LookAt:=xlPart, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                If Not Rng Is Nothing Then
                    Application.Goto Rng, True
                    Rng.Interior.ColorIndex = 37
                    Cells(Rng.Row, lastColumn).NumberFormat = "@"
                    If (UBound(SArray) - LBound(SArray) + 1) = 3 Then 'trei parti
                        Serie = SArray(2) & "/" & an
                    ElseIf Len(FindString) = 16 And Mid(FindString, 1, 4) = "1001" Then       'citeste 10 caractere + an
                        Serie = Mid(FindString, 7, 10) & "/" & an
                    ElseIf Len(FindString) = 16 And Mid(FindString, 1, 4) = "1002" Then   'citeste 10 caractere, sterge 0 din fata + an
                        Serie = Serie & "/" & an
                    ElseIf Len(FindString) = 16 And Mid(FindString, 1, 4) = "1009" Then   'citeste tot
                        Serie = Serie & "/" & an
                    ElseIf Mid(FindString, 1, 3) = "101" Then   'citeste tot
                        Serie = Serie & "/" & an
                    ElseIf InStr(FindString, "/") > 0 Then 'serie cu /
                        Serie = Serie & "/" & an
                    Else             'alt tip serie
                        Serie = FindString & "/" & an
                    End If
                    Cells(Rng.Row, colSerieCorecta) = Serie
                    Cells(Rng.Row, colContainer) = container

                    'scriu in fisier
                    IndexA = Cells(Rng.Row, colIndexA)
                    IndexR = Cells(Rng.Row, colIndexR)
                    
                    'verific anul
                    scrie = True
                    Result = vbYes
                    If an <> Cells(Rng.Row, colAn) Then Result = MsgBox("Ani diferiti " & an & "/" & Cells(Rng.Row, colAn) & " - Continuati?", vbYesNo + vbQuestion)
                    If Result = vbNo Then scrie = False
                    If scrie = True Then
                        Print #1, Serie & "#" & IndexA & "#" & IndexR & "#" & container & "#" & Cells(Rng.Row, colTipMontaj) & "#" & Cells(Rng.Row, colAssesmentCode) & "#" & Cells(Rng.Row, colEchipament)
                    End If
                Else
                    MsgBox "Nu exista " + FindString
                End If
            End With
        End If
    Wend
    'inchid fisierele
    Close #1
    If (FileLen(fileEON) = 0) Then Kill fileEON
End Sub
