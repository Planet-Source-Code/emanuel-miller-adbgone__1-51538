VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AdBGone Encoder"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Code:"
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
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2003-2004 Emanuel Miller"
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
      TabIndex        =   5
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "This encoder will work with AdBGone 1.3"
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
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Author: Emanuel Miller"
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
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image imgLogo 
      Height          =   2550
      Left            =   120
      Picture         =   "frmMain.frx":164A
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Email:"
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
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdEncode_Click()
    
    If InStr(txtEmail.Text, "@") = 0 Then Exit Sub
    If InStr(txtEmail.Text, ".") = 0 Then Exit Sub
    If Len(txtEmail.Text) = 0 Then Exit Sub
    
    txtCode.Text = reg_Encode(txtEmail.Text)
    
End Sub

Private Sub txtCode_Change()

End Sub

Private Sub txtEmail_Change()

End Sub
