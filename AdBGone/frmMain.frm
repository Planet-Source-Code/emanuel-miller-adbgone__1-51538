VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
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
      TabIndex        =   7
      Top             =   5040
      Width           =   4455
   End
   Begin VB.Label lblRegistered 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Label lblPopCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PopUps blocked:"
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
      Top             =   4560
      Width           =   4455
   End
   Begin VB.Label lblAdcount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ads blocked:"
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
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Image imgLogo 
      Height          =   2550
      Left            =   120
      Picture         =   "frmMain.frx":164A
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":AD17
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label lblADBGONE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.ccorpsoft.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2000
      MouseIcon       =   "frmMain.frx":ADE1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Author: Emanuel Miller     [                                                    ]"
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
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label lblBuild 
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub reset_links()

    '' Check to see if the website label is underlined or not
    If Not lblADBGONE.FontUnderline Then
        lblADBGONE.FontUnderline = True
    End If
    
End Sub


Private Sub Form_Load()

    With frmMain
        '' Make sure this form is not visible before
        '' we update the labels on the form
        .Visible = False
        '' Update the caption and the labels with the version
        .lblBuild = build_title & " (v" & build_version & ")"
        .Caption = build_title & " -- About"
        '' Check to see if replace ads with is used
        '' if it is, show the ad replacement count
        '' else notify the user that the counter is
        '' disabled
        If use_server = 1 Then
            .lblAdcount.Caption = "Ads blocked: " & advertisement_count
        Else
            .lblAdcount.Caption = "Ads blocked: counter disabled"
        End If
        
        If block_popups_config Then
            .lblPopCount.Caption = "Popups blocked: " & popup_count
        Else
            .lblPopCount.Caption = "Popups blocked: Popup protection disabled"
        End If
        
        If reg_registered Then
            lblRegistered.Caption = "Registered to " & reg_regemail
        Else
            lblRegistered.Caption = "Unregistered"
        End If
        
    End With
    
End Sub

Private Sub Form_LostFocus()
    
    '' Just make sure that the link is underlined
    reset_links
    
End Sub


Private Sub lblCADBGONE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '' If the mouse is down, make the color brighter
    If Button = 1 Then
        lblADBGONE.ForeColor = &HFF8080
    End If
    
End Sub

Private Sub lblADBGONE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With lblADBGONE
        '' If the mouse is down check to see if
        '' the cursor is hovering over the label
        '' if it is, increase the color of the label
        '' if it isnt, reset the color
        If Button = 1 Then
            If x > (.Left / Screen.TwipsPerPixelX) And x < .Width And y > 0 And y < .Height Then
                .ForeColor = &HFF8080
            Else
                .ForeColor = &HC00000
            End If
        End If
    End With
    
End Sub

Private Sub lblADBGONE_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    With lblADBGONE
        '' If the mouse cursor is over the label
        '' when the button is released
        '' open a browser pointed to the adbgone website
        If Button = 1 Then
            .ForeColor = &HC00000
            If x > (.Left / Screen.TwipsPerPixelX) And x < .Width And y > 0 And y < .Height Then
                On Error Resume Next
                Call Shell("rundll32.exe url.dll,FileProtocolHandler http://www.ccorpsoft.com", vbNormalFocus)
            End If
        End If
    End With
    

End Sub

