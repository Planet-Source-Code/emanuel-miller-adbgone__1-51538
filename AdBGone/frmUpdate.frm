VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image imgUpdate 
      Height          =   960
      Left            =   120
      Picture         =   "frmUpdate.frx":164A
      Top             =   75
      Width           =   960
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Form_Load()
    
    '' Set the dialog caption
    Me.Caption = build_title & " -- Status"
    
End Sub

