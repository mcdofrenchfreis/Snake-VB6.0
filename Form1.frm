VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Snake"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape box 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   0
      Left            =   135
      Top             =   135
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Snake Game for ICTPRO4 [SPACE]
' ICT - CAMARIN [CAMICTM41]
' Subject: ICTPRO4
' Instructor: Sir Louie Lee Sarte
' Group Members:
' Dela Cruz, Mar Yvan S.
' Villarino, Bea Carelle.
' Florita, Romulo.


Option Explicit

' Set the Game Speed by Changing the Value (default=50)
' - Mar Yvan Dela Cruz
Private Const vbGameSpeed As Long = 50

Private Const vbBackground As Long = &HC0E0FF
Private Const vbGridColour As Long = vbGreen
Private Const vbWallColour As Long = vbGreen
Private Const vbBonusColour As Long = vbRed
Private Const vbSnakeColour As Long = vbBlue



Private doLoop As Boolean, gotBonus As Boolean, facingUp As Boolean, facingDown As Boolean
Private facingLeft As Boolean, facingRight As Boolean, lastTickCount As Long, occupiedSquares() As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            If facingRight Then Exit Sub
            facingLeft = True: facingRight = False: facingUp = False: facingDown = False
        Case vbKeyRight
            If facingLeft Then Exit Sub
            facingRight = True: facingLeft = False: facingUp = False: facingDown = False
        Case vbKeyUp
            If facingDown Then Exit Sub
            facingUp = True: facingDown = False: facingLeft = False: facingRight = False
        Case vbKeyDown
            If facingUp Then Exit Sub
            facingDown = True: facingUp = False: facingLeft = False: facingRight = False
    End Select
End Sub


Private Sub Form_Load()
    Dim i As Long: BackColor = vbGridColour: box(0).BackColor = vbBackground
    AutoRedraw = True: loadBoard: facingDown = True: ReDim occupiedSquares(4)
    For i = 0 To 4
        occupiedSquares(i) = 50 + (5 - i)
        box(occupiedSquares(i)).BackColor = vbSnakeColour
    Next
    Me.Show: SetFocus: doGameLoop
End Sub


Private Sub addBonus()
    Dim n As Long
lblRetry:
    n = (Int((1560 - 39 + 1) * Rnd + 39))
    If (box(n).BackColor = vbWallColour) Or (box(n).BackColor = vbSnakeColour) Then GoTo lblRetry
    box(n).BackColor = vbBonusColour
End Sub


Private Sub loadBoard()
    Dim i As Long
    For i = 1 To 1599
        Load box(i)
        With box(i)
            .Move 9 + (Int((i Mod 40) * .Width)), 9 + (Int((i / 40)) * .Height)
            .Visible = True
        End With
    Next
    For i = 0 To 1599
        If (i <= 39) Or (i >= 1560) Then box(i).BackColor = vbWallColour
        If (i Mod 40) = 0 Then box(i).BackColor = vbWallColour
        If (i Mod 40) = 39 Then box(i).BackColor = vbWallColour
    Next
    addBonus
End Sub



Private Sub doGameLoop()
    doLoop = True
    Do While doLoop
        DoEvents
        If (GetTickCount - lastTickCount) >= vbGameSpeed Then
            lastTickCount = GetTickCount
            Select Case True
                Case facingUp:
                    If (Not box(occupiedSquares(0) - 40).BackColor = vbWallColour) And (Not box(occupiedSquares(0) - 40).BackColor = vbSnakeColour) Then
                        If Not box(occupiedSquares(0) - 40).BackColor = vbBonusColour Then
                            doodleBackSquare
                        Else
                            ReDim Preserve occupiedSquares(UBound(occupiedSquares) + 1)
                            gotBonus = True
                        End If
                        box(occupiedSquares(0) - 40).BackColor = vbSnakeColour
                        shiftDownArray occupiedSquares, occupiedSquares(0) - 40
                        If gotBonus Then
                            gotBonus = False
                            box(occupiedSquares(UBound(occupiedSquares))).BackColor = vbSnakeColour
                            addBonus
                        End If
                    Else
                        doLoop = False
                        doDead
                    End If
                Case facingDown:
                    If (Not box(occupiedSquares(0) + 40).BackColor = vbWallColour) And (Not box(occupiedSquares(0) + 40).BackColor = vbSnakeColour) Then
                        If Not box(occupiedSquares(0) + 40).BackColor = vbBonusColour Then
                            doodleBackSquare
                        Else
                            ReDim Preserve occupiedSquares(UBound(occupiedSquares) + 1)
                            gotBonus = True
                        End If
                        box(occupiedSquares(0) + 40).BackColor = vbSnakeColour
                        shiftDownArray occupiedSquares, occupiedSquares(0) + 40
                        If gotBonus Then
                            gotBonus = False
                            box(occupiedSquares(UBound(occupiedSquares))).BackColor = vbSnakeColour
                            addBonus
                        End If
                    Else
                        doLoop = False
                        doDead
                    End If
                Case facingLeft:
                    If (Not box(occupiedSquares(0) - 1).BackColor = vbWallColour) And (Not box(occupiedSquares(0) - 1).BackColor = vbSnakeColour) Then
                        If Not box(occupiedSquares(0) - 1).BackColor = vbBonusColour Then
                            doodleBackSquare
                        Else
                            ReDim Preserve occupiedSquares(UBound(occupiedSquares) + 1)
                            gotBonus = True
                        End If
                        box(occupiedSquares(0) - 1).BackColor = vbSnakeColour
                        shiftDownArray occupiedSquares, occupiedSquares(0) - 1
                        If gotBonus Then
                            gotBonus = False
                            box(occupiedSquares(UBound(occupiedSquares))).BackColor = vbSnakeColour
                            addBonus
                        End If
                    Else
                        doLoop = False
                        doDead
                    End If
                Case facingRight:
                    If (Not box(occupiedSquares(0) + 1).BackColor = vbWallColour) And (Not box(occupiedSquares(0) + 1).BackColor = vbSnakeColour) Then
                        If Not box(occupiedSquares(0) + 1).BackColor = vbBonusColour Then
                            doodleBackSquare
                        Else
                            ReDim Preserve occupiedSquares(UBound(occupiedSquares) + 1)
                            gotBonus = True
                        End If
                        box(occupiedSquares(0) + 1).BackColor = vbSnakeColour
                        shiftDownArray occupiedSquares, occupiedSquares(0) + 1
                        If gotBonus Then
                            gotBonus = False
                            box(occupiedSquares(UBound(occupiedSquares))).BackColor = vbSnakeColour
                            addBonus
                        End If
                    Else
                        doLoop = False
                        doDead
                    End If
            End Select
        End If
    Loop
End Sub


Private Sub doodleBackSquare()
    box(occupiedSquares(UBound(occupiedSquares))).BackColor = vbBackground
End Sub


Private Sub shiftDownArray(ByRef arr() As Long, ByVal newTopIndexValue As Long)
    Dim i As Long, x() As Long: x = arr
    For i = 1 To UBound(arr)
        arr(i) = x(i - 1)
    Next
    arr(0) = newTopIndexValue
End Sub


Private Sub doDead()
    MsgBox "Oh no! you're dead, reload the game to reset.", vbCritical Or vbOKOnly, "Snake"
    End
End Sub


Private Sub Form_Unload(Cancel As Integer)
    doLoop = False
    End
End Sub
