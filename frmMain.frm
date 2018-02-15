VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4635
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicBlocos 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   720
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   300
      ScaleWidth      =   1800
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox PicBlocos 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   720
      Picture         =   "frmMain.frx":1F6C
      ScaleHeight     =   300
      ScaleWidth      =   1800
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   3120
      ScaleHeight     =   4470
      ScaleWidth      =   1170
      TabIndex        =   1
      Top             =   75
      Width           =   1200
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   75
      ScaleHeight     =   4470
      ScaleWidth      =   2970
      TabIndex        =   0
      Top             =   75
      Width           =   3000
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuSubMenu 
         Caption         =   "New game"
         Index           =   0
      End
      Begin VB.Menu mnuSubMenu 
         Caption         =   "Pause"
         Index           =   1
      End
      Begin VB.Menu mnuSubMenu 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSubMenu 
         Caption         =   "About"
         Index           =   3
      End
      Begin VB.Menu mnuSubMenu 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSubMenu 
         Caption         =   "Exit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMain.Show
frmMain.Caption = "Initializing DirectX 7..."
DoEvents
If Not Init_DX Then GameOver = True
If Not Init_Surfaces Then GameOver = True
Call ConfigIntro
frmMain.Caption = "The Sky is Falling!"

Dim rPecaTemp As RECT
Dim lMOD As Long
lMOD = DX.TickCount()

Do
Sleep (80)
DoEvents
If bGaming = True Then
    If bPause = True Then
        BackBuffer.DrawText (picGame.Width - TextWidth("P A U S E")) / 2, (picGame.Height - TextHeight("P A U S E")) / 2, "P A U S E", False
    Else

        BackBuffer.BltColorFill rBackBuffer, 1
        rPecaTemp.Right = iLarg * 20: rPecaTemp.Bottom = iAlt * 20
        BackBuffer.BltFast Col * 20, Row * 20, surfPeca, rPecaTemp, DDBLTFAST_WAIT
        DesenhaBackBuffer
        
        If ReqLeft = True Then
            If Col > 0 And LeftColision = False Then Col = Col - 1
        End If
        If ReqRight = True Then
            If Col + iLarg < 10 And RightColision = False Then Col = Col + 1
        End If
        If ReqDown = True Then
            If Collision = False Then Row = Row + 1
        End If
        
        If DX.TickCount() - lMOD > 1000 - (Level * 100) Then
            If Collision = True Then
                Call PreencheCollision
                ChecaLinhasCompletas
                NovasPecas
                If Collision = True Then bGameOver = True: bGaming = False
            Else
                Row = Row + 1
            End If
            lMOD = DX.TickCount()
        End If
    End If
    
ElseIf bGaming = False And bGameOver = False Then
    Dim rTemp As RECT
    rTemp.Left = 70: rTemp.Top = 100: rTemp.Right = 200: rTemp.Bottom = 200
    BackBuffer.BltColorFill rTemp, 1
    BackBuffer.BltFast 70, 100, surfPeca, rPeca, DDBLTFAST_WAIT
    iBloco = iBloco + 1
    If iBloco = 7 Then iBloco = 0
    SetPecaValues
ElseIf bGameOver = True Then
    BackBuffer.BltColorFill rBackBuffer, 1
    
    iFont.Italic = False: iFont.Size = 20: iFont.Bold = True: Me.FontSize = 20
    BackBuffer.SetFont iFont: BackBuffer.SetForeColor vbRed
    BackBuffer.DrawText 10, 100, "Game Over!", False
    
    iFont.Size = 12
    BackBuffer.SetFont iFont: BackBuffer.SetForeColor vbWhite
    BackBuffer.DrawText (200 - TextWidth("Score:" & Pontos)) / 2, 140, "Score:", False
    BackBuffer.DrawText (200 - TextWidth("Level:" & Level)) / 2, 155, "Level:", False
    
    iFont.Italic = True
    BackBuffer.SetFont iFont: BackBuffer.SetForeColor vbWhite
    BackBuffer.DrawText ((200 - TextWidth("Score:")) / 2) + TextWidth("Score:") + 10, 140, Pontos, False
    BackBuffer.DrawText ((200 - TextWidth("Level:")) / 2) + TextWidth("Level:") + 10, 155, Level, False
End If
Flip
Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End_DX
End
End Sub

Private Sub mnuSubMenu_Click(Index As Integer)
If Index = 0 Then
    Dim t As Integer
    Pontos = 0: Level = 0
    NovasPecas
    BackBuffer.BltColorFill rBackBuffer, 1
    bPause = False: bGaming = True: bGameOver = False
    Dim i As Integer, ii As Integer
    For i = 0 To 9
        For ii = 0 To 14
            Tela(i, ii) = 0
        Next ii
    Next i
ElseIf Index = 1 Then
    If bPause = True Then bPause = False Else: bPause = True
ElseIf Index = 3 Then
    MsgBox "The Sky is Falling v1.0 " & vbCrLf & "Author: Leandro Del Arco", vbInformation, "The Sky is Falling!"
ElseIf Index = 5 Then
    End
End If
End Sub

Private Sub picGame_KeyDown(KeyCode As Integer, Shift As Integer)
If bGaming = False Then Exit Sub
If KeyCode = vbKeyP Then
    If bPause = True Then bPause = False Else: bPause = True
ElseIf KeyCode = vbKeyLeft Then
    If bPause = False Then ReqLeft = True
ElseIf KeyCode = vbKeyRight Then
    If bPause = False Then ReqRight = True
ElseIf KeyCode = vbKeyUp Then
    If bPause = True Then Exit Sub
    Select Case iBloco
    Case 0: iBlocoSub = 0
    Case 1: iBlocoSub = iBlocoSub + 1: If iBlocoSub = 2 Then iBlocoSub = 0
    Case 2: iBlocoSub = iBlocoSub + 1: If iBlocoSub = 4 Then iBlocoSub = 0
    Case 3: iBlocoSub = iBlocoSub + 1: If iBlocoSub = 4 Then iBlocoSub = 0
    Case 4: iBlocoSub = iBlocoSub + 1: If iBlocoSub = 4 Then iBlocoSub = 0
    Case 5: iBlocoSub = iBlocoSub + 1: If iBlocoSub = 2 Then iBlocoSub = 0
    Case 6: iBlocoSub = iBlocoSub + 1: If iBlocoSub = 2 Then iBlocoSub = 0
    End Select
    Call SetPecaValues
    If Col + iLarg > 9 Then Col = 10 - iLarg
    If Row + iAlt > 15 Then Row = 15 - iAlt
    Call ChangeCollision
ElseIf KeyCode = vbKeyDown Then
    If bPause = False Then ReqDown = True
End If
End Sub

Sub Flip()
DDClipper.SetHWnd picGame.hWnd: PrimarySurf.SetClipper DDClipper
DX.GetWindowRect picGame.hWnd, rScr
PrimarySurf.Blt rScr, BackBuffer, rBackBuffer, DDBLT_WAIT

DDClipper.SetHWnd picStatus.hWnd: PrimarySurf.SetClipper DDClipper
DX.GetWindowRect picStatus.hWnd, rScr
PrimarySurf.Blt rScr, BackBufferStatus, rBackBufferStatus, DDBLT_WAIT
End Sub

Sub ConfigIntro()
With frmMain: .Font = "verdana": .FontSize = 12: .FontItalic = True: End With
Set iFont = frmMain.Font

BackBuffer.SetFont iFont
BackBuffer.SetForeColor vbWhite
BackBuffer.BltColorFill rBackBuffer, 1
BackBuffer.DrawText (picGame.Width - TextWidth("The Sky is Falling v1.0")) / 2, 10, "The Sky is Falling v1.0", False
BackBuffer.DrawText (picGame.Width - TextWidth("Leandro Del Arco")) / 2, 30, "Leandro Del Arco", False

BackBufferStatus.SetForeColor vbWhite
BackBufferStatus.DrawText 25, 10, "Next", False
BackBufferStatus.DrawText 21, 100, "Level", False
BackBufferStatus.DrawText 17, 150, "Score", False
DrawLevelPontos
End Sub

Private Sub picGame_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then ReqLeft = False
If KeyCode = vbKeyRight Then ReqRight = False
If KeyCode = vbKeyDown Then ReqDown = False
End Sub
