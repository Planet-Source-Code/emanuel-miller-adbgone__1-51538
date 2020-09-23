VERSION 5.00
Begin VB.Form frmContribute 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmContribute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtHosts 
      Height          =   1245
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtSite 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblHosts 
      BackStyle       =   0  'Transparent
      Caption         =   "List the host(s) here:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblSite 
      BackStyle       =   0  'Transparent
      Caption         =   "Website you found the host(s):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "This information will be sent and added to the hosts list after evaluation."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmContribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblStatus_Click()

End Sub


Private Sub cmdSend_Click()

    '' User error checking statements
    If Len(txtSite.Text) = 0 Then
        MsgBox "You must enter the website that you found this advertisement (host).", vbInformation, build_title
        Exit Sub
    End If
    
    If Len(txtHosts.Text) = 0 Then
        MsgBox "You must add atleast one host.", vbInformation, build_title
        Exit Sub
    End If
    
    '' Send the data to the specified email address
    contribute_Hosts txtSite.Text, txtHosts.Text
    
End Sub

Private Sub Form_Load()
    
    '' Disable the icon menu
    menu_cancel_flag = True
    Me.Caption = build_title & " -- Contribute"
    
End Sub

Private Sub lblHost_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '' Enable the icon menu
    menu_cancel_flag = False
    
End Sub
