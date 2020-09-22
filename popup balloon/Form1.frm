VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Popup message - UPDATED!"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Exit Program"
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   6960
      Width           =   3015
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Remove System Tray Icon"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   7440
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add To System Tray"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   6960
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Standard - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Warning - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "None - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Error - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   240
      Picture         =   "Form1.frx":0000
      Top             =   480
      Width           =   3735
   End
   Begin VB.Image Image5 
      Height          =   1035
      Left            =   240
      Picture         =   "Form1.frx":C9DE
      Top             =   5760
      Width           =   3765
   End
   Begin VB.Image Image4 
      Height          =   1050
      Left            =   240
      Picture         =   "Form1.frx":195E4
      Top             =   4440
      Width           =   3765
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   240
      Picture         =   "Form1.frx":264DE
      Top             =   3120
      Width           =   3765
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   240
      Picture         =   "Form1.frx":330E4
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Line Line6 
      X1              =   6600
      X2              =   240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      X1              =   6600
      X2              =   240
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line4 
      X1              =   6600
      X2              =   240
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line3 
      X1              =   6600
      X2              =   240
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   240
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATED! - Don't forget to Vote!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'::::::::::::::::::::::::::::::::::::::::::::::::::::
' Popup Balloon UPDATED!
' Now Supports Different Icons
'::::::::::::::::::::::::::::::::::::::::::::::::::::


Private Sub Command1_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
'Popup the balloon - Error Icon
Popup_error "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command2_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - No Icon
Popup_none "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command3_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - Error Warning
Popup_warning "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command4_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - Info Icon
Popup_info "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command5_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - Error Standard
Popup "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command6_Click()

'Removes Icon
Shell_NotifyIcon NIM_DELETE, m_IconData

'Exits Program
End

End Sub

Private Sub Command7_Click()

'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
End Sub

Private Sub Command8_Click()
Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

Private Sub Form_Load()

 'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

'get rid of the icon in the system tray
 Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

