VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbCode 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   4455
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register"
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label lblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Code:"
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
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   1335
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Enter Code"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy Now!"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Image imgLogo 
      Height          =   2550
      Left            =   120
      Picture         =   "frmReg.frx":164A
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label lblBuy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Register Today!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4455
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub cmdBuy_Click()

    reg_buyNow

End Sub


Private Sub cmdCancel_Click()

If reg_count > 300 Then
lblInfo.Caption = "You have reached the unregistered protection limit!" & Chr(13) & Chr(10) & _
                  "If you wish to continue to use AdBGone, you must register." & Chr(13) & Chr(10) & _
                  "Registration is cheap, fast, and for life." & Chr(13) & Chr(10) & _
                  "AdBGone will be disabled until you register."
Else
    lblInfo.Caption = "Why wait for AdBGone to block 300 advertisements.." & Chr(13) & Chr(10) & _
                      "Buy AdBGone now!" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                      "Wait a few moments and click Try It to continue your trial." & Chr(13) & Chr(10) & _
                      "AdBGone will only block " & (300 - reg_count) & " more advertisements."
                      
End If
pbCode.Visible = False
    lblInfo.Height = 1215
    
End Sub

Private Sub cmdClose_Click()

    If reg_count > 300 Then
        end_Main
    Else
        menu_cancel_flag = False
        Unload Me
    End If
    
End Sub

Private Sub cmdReg_Click()
    
    lblInfo.Caption = "Enter the information that you received in your Email, then click Register."
    pbCode.Visible = True
    txtEmail.SetFocus
    lblInfo.Height = 480
    
End Sub



Private Sub cmdRegister_Click()

    If Len(txtEmail.Text) = 0 Then
        txtEmail.SetFocus
        lblEmail.FontUnderline = True
        lblEmail.ForeColor = &HC0&
        Beep
        Exit Sub
    Else
        lblEmail.ForeColor = &H80000012
        lblEmail.FontUnderline = False
    End If
    
    If Len(txtCode.Text) = 0 Then
        txtCode.SetFocus
        lblCode.FontUnderline = True
        lblCode.ForeColor = &HC0&
        Beep
        Exit Sub
    Else
        lblCode.ForeColor = &H80000012
        lblCode.FontUnderline = False
    End If
    
    txtEmail.Enabled = False
    txtCode.Enabled = False
    cmdCancel.Enabled = False
    cmdRegister.Enabled = False
    
    Dim check_it As Boolean
    
    check_it = reg_checkCode(txtCode.Text, txtEmail.Text)
    
    If check_it Then
        GoTo code_valid
    Else
        GoTo code_invalid
    End If
    
code_valid:
    lblInfo.Caption = "This code seems valid!" & Chr(13) & Chr(10) & "Saving Registration..."
    timeout 2
    If Not reg_saveReg(txtCode.Text, txtEmail.Text) Then
        Call SetTopMostWindow(frmReg.hwnd, False)
        MsgBox "There was an error saving your registration, please contact 3miller@sbcglobal.net", vbCritical, build_title
        end_Main
    Else
        Call SetTopMostWindow(frmReg.hwnd, False)
        MsgBox "Your registration was saved." & Chr(13) & Chr(10) & "Enjoy " & build_title & "!", vbInformation, build_title
    End If
    Unload Me
    timeout 1
    Main
    Exit Sub
    
code_invalid:
    lblInfo.Caption = "The Email or Registration Code you entered is invalid!" & Chr(13) & Chr(10) & "Please contact 3miller@sbcglobal.net for assistance."
    Beep
    timeout 0.1
    Beep
    timeout 0.1
    Beep
    GoTo ender
    

ender:
    txtEmail.Enabled = True
    txtCode.Enabled = True
    cmdCancel.Enabled = True
    cmdRegister.Enabled = True
    
    
End Sub

Private Sub Form_Load()

menu_cancel_flag = True

If reg_count > 300 Then
lblInfo.Caption = "You have reached the unregistered protection limit!" & Chr(13) & Chr(10) & _
                  "If you wish to continue to use AdBGone, you must register." & Chr(13) & Chr(10) & _
                  "Registration is cheap, fast, and for life." & Chr(13) & Chr(10) & _
                  "AdBGone will be disabled until you register."
Else
    lblInfo.Caption = "Why wait for AdBGone to block 300 advertisements.." & Chr(13) & Chr(10) & _
                      "Buy AdBGone now!" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                      "Wait a few moments and click Try It to continue your trial." & Chr(13) & Chr(10) & _
                      "AdBGone will only block " & (300 - reg_count) & " more advertisements."
                      
End If

End Sub

Private Sub lblReg_Click()

End Sub

