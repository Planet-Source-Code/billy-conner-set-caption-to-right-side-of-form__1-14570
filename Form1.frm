VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Just A Sample"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Set caption to right side"
      Height          =   330
      Left            =   255
      TabIndex        =   1
      Tag             =   "Set Caption To Left"
      Top             =   480
      Width           =   2505
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make form style backwards"
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   945
      Width           =   2475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = -20
Private Const SWP_NOZORDER = 4
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const WS_EX_LAYOUTRTL = &H400000
Private Const WS_EX_RIGHT = &H1000


Private Sub Command1_Click()
Const Fwd As String = "Make form style normal"
Const Bkw As String = "Make form style backwards"
Static OnOff As Boolean

OnOff = Not (OnOff)
SetWindowStyleEx Me.hWnd, WS_EX_LAYOUTRTL, OnOff, True

If OnOff Then 'Update The Command Button
    Command1.Caption = Fwd
Else
    Command1.Caption = Bkw
End If

End Sub


Private Sub Command2_Click()
Static OnOff As Boolean
Const Lft As String = "Set caption to left side"
Const rgt As String = "Set caption to right side"

OnOff = Not (OnOff)
SetWindowStyleEx Me.hWnd, WS_EX_RIGHT, OnOff, True

If OnOff Then 'Update The Command Button
    Command2.Caption = Lft
Else
    Command2.Caption = rgt
End If

End Sub
Private Sub SetWindowStyleEx(wnd As Long, NewStyle As Long, fAdd As Boolean, Optional fRedraw As Boolean = True)
Dim CurStyle As Long
CurStyle = GetWindowLong(wnd, GWL_EXSTYLE)
If fAdd And (CurStyle And NewStyle) = 0 Then
    ' Setting the new style and it is not already set...
    CurStyle = CurStyle Or NewStyle
ElseIf (Not fAdd) And (CurStyle And NewStyle) Then
    ' Removing the new style and it's already set...
    CurStyle = CurStyle And (Not NewStyle)
End If
SetWindowLong wnd, GWL_EXSTYLE, CurStyle
If fRedraw Then
    SetWindowPos wnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE
End If

End Sub
