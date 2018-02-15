Attribute VB_Name = "modGame"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public iFont As New StdFont
Public bGaming As Boolean
Public bPause As Boolean
Public bGameOver As Boolean

Public Col As Integer '10 Colunas
Public Row As Integer '15 Linhas

Public iBloco As Integer
Public iBlocoSub As Integer
Public iNextBloco As Integer

Public Tela(9, 14) As Integer
Public Peca(3, 3) As Integer

Public iAlt As Integer, iLarg As Integer

Public Pontos As Long
Public Level As Integer

Public ReqDown As Boolean
Public ReqLeft As Boolean
Public ReqRight As Boolean

Sub PreencheCollision()
Dim x As Integer, y As Integer
For x = 0 To 3
    For y = 0 To 3
        If Peca(x, y) <> 0 Then Tela(Col + x, Row + y) = Peca(x, y)
    Next y
Next x
End Sub

Function Collision() As Boolean
If Row + iAlt = 15 Then Collision = True: Exit Function
Dim x As Integer, y As Integer
For x = 0 To 3
    For y = 0 To 3
        If Peca(x, y) <> 0 Then
            If Tela(Col + x, Row + y + 1) <> 0 Then Collision = True: Exit Function
        End If
    Next y
Next x
End Function

Sub ChangeCollision()
Dim x As Integer, y As Integer
For x = 0 To 3
    For y = 0 To 3
        If Peca(x, y) <> 0 Then
            If Tela(Col + x, Row + y) <> 0 Then
                Row = (Row + y) - iAlt
            End If
        End If
    Next y
Next x
End Sub

Sub SetPecaValues()
Dim i As Integer, ii As Integer
Dim rTemp As RECT

For i = 0 To 3
    For ii = 0 To 3
        Peca(i, ii) = 0
    Next ii
Next i

If iBloco = 0 Then
    Peca(0, 0) = 1: Peca(1, 0) = 1: Peca(0, 1) = 1: Peca(1, 1) = 1
    iLarg = 2: iAlt = 2
ElseIf iBloco = 1 Then
    If iBlocoSub = 0 Then
        Peca(0, 0) = 4: Peca(1, 0) = 4: Peca(2, 0) = 4: Peca(3, 0) = 4
        iLarg = 4: iAlt = 1
    ElseIf iBlocoSub = 1 Then
        Peca(0, 0) = 4: Peca(0, 1) = 4: Peca(0, 2) = 4: Peca(0, 3) = 4
        iLarg = 1: iAlt = 4
    End If
ElseIf iBloco = 2 Then
    If iBlocoSub = 0 Then
        Peca(0, 0) = 3: Peca(1, 0) = 3: Peca(0, 1) = 3: Peca(0, 2) = 3
        iLarg = 2: iAlt = 3
    ElseIf iBlocoSub = 1 Then
        Peca(0, 0) = 3: Peca(1, 0) = 3: Peca(2, 0) = 3: Peca(2, 1) = 3
        iLarg = 3: iAlt = 2
    ElseIf iBlocoSub = 2 Then
        Peca(1, 0) = 3: Peca(1, 1) = 3: Peca(0, 2) = 3: Peca(1, 2) = 3
        iLarg = 2: iAlt = 3
    ElseIf iBlocoSub = 3 Then
        Peca(0, 0) = 3: Peca(0, 1) = 3: Peca(1, 1) = 3: Peca(2, 1) = 3
        iLarg = 3: iAlt = 2
    End If
ElseIf iBloco = 3 Then
    If iBlocoSub = 0 Then
        Peca(0, 0) = 6: Peca(1, 0) = 6: Peca(1, 1) = 6: Peca(1, 2) = 6
        iLarg = 2: iAlt = 3
    ElseIf iBlocoSub = 1 Then
        Peca(2, 0) = 6: Peca(0, 1) = 6: Peca(1, 1) = 6: Peca(2, 1) = 6
        iLarg = 3: iAlt = 2
    ElseIf iBlocoSub = 2 Then
        Peca(0, 0) = 6: Peca(0, 1) = 6: Peca(0, 2) = 6: Peca(1, 2) = 6
        iLarg = 2: iAlt = 3
    ElseIf iBlocoSub = 3 Then
        Peca(0, 0) = 6: Peca(1, 0) = 6: Peca(2, 0) = 6: Peca(0, 1) = 6
        iLarg = 3: iAlt = 2
    End If
ElseIf iBloco = 4 Then
    If iBlocoSub = 0 Then
        Peca(1, 0) = 2: Peca(0, 1) = 2: Peca(1, 1) = 2: Peca(2, 1) = 2
        iLarg = 3: iAlt = 2
    ElseIf iBlocoSub = 1 Then
        Peca(0, 0) = 2: Peca(0, 1) = 2: Peca(1, 1) = 2: Peca(0, 2) = 2
        iLarg = 2: iAlt = 3
    ElseIf iBlocoSub = 2 Then
        Peca(0, 0) = 2: Peca(1, 0) = 2: Peca(2, 0) = 2: Peca(1, 1) = 2
        iLarg = 3: iAlt = 2
    ElseIf iBlocoSub = 3 Then
        Peca(1, 0) = 2: Peca(0, 1) = 2: Peca(1, 1) = 2: Peca(1, 2) = 2
        iLarg = 2: iAlt = 3
    End If
ElseIf iBloco = 5 Then
    If iBlocoSub = 0 Then
        Peca(0, 0) = 4: Peca(1, 0) = 4: Peca(1, 1) = 4: Peca(2, 1) = 4
        iLarg = 3: iAlt = 2
    ElseIf iBlocoSub = 1 Then
        Peca(1, 0) = 4: Peca(0, 1) = 4: Peca(1, 1) = 4: Peca(0, 2) = 4
        iLarg = 2: iAlt = 3
    End If
ElseIf iBloco = 6 Then
    If iBlocoSub = 0 Then
        Peca(1, 0) = 5: Peca(2, 0) = 5: Peca(0, 1) = 5: Peca(1, 1) = 5
        iLarg = 3: iAlt = 2
    ElseIf iBlocoSub = 1 Then
        Peca(0, 0) = 5: Peca(0, 1) = 5: Peca(1, 1) = 5: Peca(1, 2) = 5
        iLarg = 2: iAlt = 3
    End If
End If

surfPeca.BltColorFill rPeca, 1
For i = 0 To 3
    For ii = 0 To 3
        If Peca(i, ii) <> 0 Then
            rTemp.Left = (Peca(i, ii) - 1) * 20
            rTemp.Right = rTemp.Left + 20
            rTemp.Bottom = 20
            surfPeca.BltFast i * 20, ii * 20, surfBlocos, rTemp, DDBLTFAST_WAIT
        End If
    Next ii
Next i
End Sub

Sub DesenhaBackBuffer()
Dim i As Integer, ii As Integer
Dim rTemp As RECT
For i = 0 To 9
    For ii = 0 To 14
        If Tela(i, ii) <> 0 Then
            rTemp.Left = (Tela(i, ii) - 1) * 20
            rTemp.Right = rTemp.Left + 20
            rTemp.Bottom = 20
            BackBuffer.BltFast i * 20, ii * 20, surfBlocos, rTemp, DDBLTFAST_WAIT
        End If
    Next ii
Next i
End Sub

Sub ChecaLinhasCompletas()
Dim x As Integer, y As Integer
Dim x2 As Integer, y2 As Integer
For y = Row To Row + iAlt - 1
    For x = 0 To 9
        If Tela(x, y) = 0 Then GoTo Quebra
    Next x
    For y2 = y To 1 Step -1
        For x2 = 0 To 9
            Tela(x2, y2) = Tela(x2, y2 - 1)
        Next x2
    Next y2
    Pontos = Pontos + 10
    If Pontos = 100 Then Level = 1
    If Pontos = 200 Then Level = 2
    If Pontos = 300 Then Level = 3
    If Pontos = 400 Then Level = 4
    If Pontos = 500 Then Level = 5
    If Pontos = 600 Then Level = 6
    If Pontos = 700 Then Level = 7
    If Pontos = 800 Then Level = 8
    If Pontos = 900 Then Level = 9
    If Pontos = 1000 Then Level = 10
    DrawLevelPontos
Quebra:
Next y
End Sub

Function LeftColision() As Boolean
If Col = 0 Then Exit Function
Dim x As Integer, y As Integer
For x = 0 To 3
    For y = 0 To 3
        If Peca(x, y) <> 0 Then
            If Tela(Col + x - 1, Row + y) <> 0 Then LeftColision = True: Exit Function
        End If
    Next y
Next x
End Function

Function RightColision() As Boolean
If Col + iLarg = 10 Then Exit Function
Dim x As Integer, y As Integer
For x = 0 To 3
    For y = 0 To 3
        If Peca(x, y) <> 0 Then
            If Tela(Col + x + 1, Row + y) <> 0 Then RightColision = True: Exit Function
        End If
    Next y
Next x
End Function

Sub DrawLevelPontos()
Dim rTemp As RECT
rTemp.Right = 80: rTemp.Top = 120: rTemp.Bottom = 150
BackBufferStatus.BltColorFill rTemp, 1
BackBufferStatus.DrawText (80 - frmMain.TextWidth(Level)) / 2, 120, Level, False
rTemp.Right = 80: rTemp.Top = 170: rTemp.Bottom = 200
BackBufferStatus.BltColorFill rTemp, 1
BackBufferStatus.DrawText (80 - frmMain.TextWidth(Pontos)) / 2, 170, Pontos, False
End Sub

Sub NovasPecas()
Dim t As Integer
Dim rPecaTemp As RECT
Randomize Time
iBlocoSub = 0
iBloco = Rnd() * 6
SetPecaValues
rPecaTemp.Top = 35
rPecaTemp.Right = 80
rPecaTemp.Bottom = 95
BackBufferStatus.BltColorFill rPecaTemp, 1
rPecaTemp.Top = 0
rPecaTemp.Right = iLarg * 20
rPecaTemp.Bottom = iAlt * 20
BackBufferStatus.BltFast (80 - rPecaTemp.Right) / 2, 35, surfPeca, rPecaTemp, DDBLTFAST_WAIT
t = iBloco
iBloco = iNextBloco
iNextBloco = t
SetPecaValues
Row = 0: Col = 5 - Int(iLarg / 2)
End Sub
