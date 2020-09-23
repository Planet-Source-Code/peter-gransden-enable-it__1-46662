VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enable-It"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   705
      Left            =   2355
      ScaleHeight     =   645
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   90
      Width           =   915
      Begin VB.Image imgTarget 
         Height          =   480
         Left            =   180
         Picture         =   "frmMain.frx":0000
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable"
      Height          =   315
      Left            =   390
      TabIndex        =   0
      Top             =   540
      Width           =   1530
   End
   Begin VB.Label lblClassMane 
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   285
      Width           =   2145
   End
   Begin VB.Label Label2 
      Caption         =   "Drag the + cursor to the button/control tick box and release it, then use the enable/disable button."
      Height          =   585
      Left            =   45
      TabIndex        =   4
      Top             =   960
      Width           =   3330
   End
   Begin VB.Image imgCross 
      Height          =   480
      Left            =   345
      Picture         =   "frmMain.frx":030A
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgNull 
      Height          =   15
      Left            =   270
      Picture         =   "frmMain.frx":0614
      Top             =   2055
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label1 
      Caption         =   "Window Handle:"
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   45
      Width           =   1230
   End
   Begin VB.Label lblHWnd 
      Caption         =   "0000000000"
      Height          =   240
      Left            =   1260
      TabIndex        =   2
      Top             =   60
      Width           =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/******************************************************************************
'Name: frmMain.frm (frmMain)
'
'Description: Main form for this project contains the use interface controls.
'
'Date Updated: 04/July/2003.
'
'Author: Peter Gransden.
'/******************************************************************************

'/******************************************************************************
Private Sub cmdEnable_Click()
'/******************************************************************************
'Description: Button event to enable or disable a window/control,
'When you press the button it will send the EnableWindow API
'call to the to the windows handle, this
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
    
    'This only works like a switch just sends ether an ON or Off,
    'It doesn't correspond to the current windows/controls state.
    If ControllEnabled = True Then
        cmdEnable.Caption = "Enable"
        EnableWindow lblHWnd.Caption, 0 ' Disable < This is the API call
        ControllEnabled = False ' Sets swich
    Else
        cmdEnable.Caption = "Disable"
        EnableWindow lblHWnd.Caption, 1 ' Enable < This is the API call
        ControllEnabled = True ' Sets Swich
    End If

End Sub


'/******************************************************************************
Private Sub imgTarget_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'/******************************************************************************
'Description: This changes the mouse cursor to +, and sets Targeting to True.
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
    
    'We are Targeting
    Targeting = True
    'Blanks the target icon in the Picture Box
    imgTarget.Picture = imgNull.Picture
    'Sets the mouse cursor to Custom.
    Me.MousePointer = 99
    'Sets the mouse cursor to the +, in the picture box
    Me.MouseIcon = imgCross.Picture

End Sub

'/******************************************************************************
Private Sub imgTarget_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'/******************************************************************************
'Description: Gets all of the information about what's underneath the mouse cursor.
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
Dim Child As Long ' Holds the Child's Hwnd
Dim WindowX As Long ' Holds then X posishin of the mouse cursor
Dim WindowY As Long ' Holds then Y posishin of the mouse cursor
Dim sName As String ' Holds the display name of the control
Dim sClassName As String * 255 ' Holds the class name returned by the GetClassName API
Dim TempHwnd As Long ' Holds the Hind of an object
    
    'If we aren't targeting then do nothing.
    If Targeting = False Then Exit Sub
        'Call to get the mouse position.
        Call GetCursorPos(CursorPosition)
        'Get the windows Handle from the cursors position
        TempHwnd = WindowFromPoint(CursorPosition.x, CursorPosition.y)
        'Find the point on a window.
        GetWindowPoint TempHwnd, WindowX, WindowY
        'Find the Child object on a window (if any).
        Child = ChildWindowFromPoint(TempHwnd, WindowX, WindowY)
        
        'Ether use the child or windows Hwnd
        If Child = 0 Then
            'Get the class name of a window.
            Call GetClassName(TempHwnd, sClassName, 255)
            lblHWnd.Caption = TempHwnd
            ControllEnabled = True
        Else
            'Get the class name of the Child
            Call GetClassName(Child, sClassName, 255)
            lblHWnd.Caption = Child
            ControllEnabled = False
        End If
        
        'Format and display the information.
        sName = Trim(left(sClassName, InStr(sClassName, vbNullChar) - 1))
        lblClassMane.Caption = sName

    
End Sub

'/******************************************************************************
Private Sub imgTarget_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'/******************************************************************************
'Description: This changes the mouse cursor back to default
'and sets Targeting to False.
'
'Inputs: None
'
'Returns: None
'/******************************************************************************
    
    'We are not Targeting anymore.
    Targeting = False
    'Sets the picture box back the + cursor
    imgTarget.Picture = imgCross.Picture
    'Sets the mouse cursor back to default.
    Me.MousePointer = 0

End Sub

