VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbSlab1 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2415
      ScaleWidth      =   4215
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtHTML 
         Enabled         =   0   'False
         Height          =   735
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1100
         Width           =   3015
      End
      Begin VB.OptionButton optRepHTML 
         Caption         =   "HTML:"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1050
         Width           =   1215
      End
      Begin VB.PictureBox pbColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   9
         ToolTipText     =   "Click to change"
         Top             =   760
         Width           =   975
      End
      Begin MSComDlg.CommonDialog cmnDlg 
         Left            =   3720
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chkUseServer 
         Caption         =   "Replace ads with:"
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
         Top             =   120
         Width           =   1935
      End
      Begin VB.OptionButton optRepBG 
         Caption         =   "Background color:"
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
         Left            =   360
         TabIndex        =   4
         Top             =   760
         Width           =   1695
      End
      Begin VB.OptionButton optRepLogo 
         Caption         =   "Logo"
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lblAdInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Advertisement counter is disabled when 'Replace ads with' is disabled."
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
         TabIndex        =   6
         Top             =   1920
         Width           =   3975
      End
   End
   Begin VB.PictureBox pbSlab0 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2415
      ScaleWidth      =   4215
      TabIndex        =   7
      Top             =   600
      Width           =   4215
      Begin VB.CheckBox chkPopUp 
         Caption         =   "Enable Pop-Up protection"
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
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chkClearIE 
         Caption         =   "Auto clear IE Cache when Enabled"
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
         TabIndex        =   12
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox chkLAB 
         Caption         =   "Load when Windows starts"
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
         TabIndex        =   8
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmConfig.frx":164A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin MSComctlLib.TabStrip tbStrip 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advertisements"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub chkUseServer_Click()
    
    '' Enable/Disable the correct options depending on
    '' if Replace Ads with is checked or not
    If chkUseServer.Value = 1 Then
        optRepLogo.Enabled = True
        optRepBG.Enabled = True
        If optRepBG.Value Then
            pbColor.Enabled = True
        End If
        optRepHTML.Enabled = True
        If optRepHTML.Value Then
            txtHTML.Enabled = True
        End If
    Else
        optRepLogo.Enabled = False
        optRepBG.Enabled = False
        optRepHTML.Enabled = False
        txtHTML.Enabled = False
    End If
    
End Sub

Private Sub cmdSave_Click()

    '' If we encounter an error, skip to the err_ender goto tag
    On Error GoTo err_ender
    
    '' Check to see if we should use the ad replacement server
    If chkUseServer.Value = 1 Then
        use_server = 1
        '' update the tray caption to display if it is enabled
        '' or not
        If frmMenu.menuProton.Checked Then
            tray_update frmMenu, build_title & " (enabled)" & vbCrLf & "Ads blocked: " & advertisement_count & vbCrLf & "PopUps blocked: " & popup_count
        Else
            tray_update frmMenu, build_title & " (disabled)" & vbCrLf & "Ads blocked: " & advertisement_count & vbCrLf & "PopUps blocked: " & popup_count
        End If
    Else
        use_server = 0
        If frmMenu.menuProton.Checked Then
            tray_update frmMenu, build_title & " (enabled)" & vbCrLf & "Ads blocked: counter disabled" & vbCrLf & "PopUps blocked: " & popup_count
        Else
            tray_update frmMenu, build_title & " (disabled)" & vbCrLf & "Ads blocked: counter disabled" & vbCrLf & "PopUps blocked: " & popup_count
        End If
    End If
    
    '' Enable/disable load at boot depending on selection
    If chkLAB.Value = 1 Then
        load_at_boot = True
        enable_LoadAtBoot
    Else
        load_at_boot = False
        disable_LoadAtBoot
    End If
    
    '' Enable/disable Clearing of IE's cache
    If chkClearIE.Value = 1 Then
        clear_ie_cache = True
    Else
        clear_ie_cache = False
    End If
    
    '' Select the server_mode depending on the selection
    If optRepLogo.Value Then
        server_mode = 1
    ElseIf optRepBG.Value Then
        server_mode = 2
    Else
        server_mode = 3
    End If
    
    '' Enable/disable popup protection
    If chkPopUp.Value = 1 Then
        block_popups_config = True
    Else
        block_popups_config = False
    End If
    
    '' Set the server color to the selection made
    server_color = pbColor.BackColor
    
    '' Set the use_server value to the selection made
    use_server = chkUseServer.Value
    
    '' Check to see if protection is already enabled
    '' If it is enabled, and Replace ads with is checked
    '' we must enable the advertisement replacement
    '' server, if not disable it
    If frmMenu.menuProton.Checked Then
        If chkUseServer.Value = 1 Then
            start_adv_host
        Else
            stop_adv_host
        End If
    End If
    
    If chkPopUp.Value = 1 Then
        block_popups = True
    Else
        block_popups = False
    End If
    
    '' Set the server_html to the user defined html
    server_html = txtHTML.Text
    
    '' Make sure the file is closed
    Close #1
    '' Open the custom html file for writing and print
    '' the user defined html
    Open html_file For Output As #1
        Print #1, server_html
    Close #1
    
    '' Unload the configuration dialog
    Unload Me
    
    '' Write the configuration to the registry
    save_Config
    Exit Sub
    
err_ender:
    MsgBox "There was an error saving your configuration.", vbCritical, build_title
    Unload Me
    
End Sub



Private Sub Form_Load()
    
    '' Disable icon menu
    menu_cancel_flag = True
    Me.Caption = build_title & " -- Protection Configuration"
    
    '' Set the value of use server to the current setting
    chkUseServer.Value = use_server

    '' Check to see if we should set loading at boot
    If load_at_boot Then
        chkLAB.Value = 1
    Else
        chkLAB.Value = 0
    End If
    
    '' Check to see if we should set Clear IE Cache
    If clear_ie_cache Then
        chkClearIE.Value = 1
    Else
        chkClearIE.Value = 0
    End If
    
    '' Check to see if we should set popup protection
    If block_popups_config Then
        chkPopUp.Value = 1
    Else
        chkPopUp.Value = 0
    End If
    
    '' Check to see what server mode option should be set
    If server_mode = 1 Then
        optRepLogo.Value = True
        optRepBG.Value = False
        optRepHTML.Value = False
    ElseIf server_mode = 2 Then
        optRepBG.Value = True
        optRepLogo.Value = False
        optRepHTML.Value = False
    ElseIf server_mode = 3 Then
        optRepHTML.Value = True
        optRepLogo.Value = False
        optRepBG.Value = False
    End If
    
    
    '' Check to see if we should set the replace ads with
    If chkUseServer.Value = 1 Then
        optRepLogo.Enabled = True
        optRepBG.Enabled = True
        optRepHTML.Enabled = True
        If optRepBG.Value = True Then
            pbColor.Enabled = True
        Else
            pbColor.Enabled = False
        End If
        If optRepHTML.Value = True Then
            txtHTML.Enabled = True
        Else
            txtHTML.Enabled = False
        End If
    Else
        optRepLogo.Enabled = False
        optRepBG.Enabled = False
        optRepHTML.Enabled = False
        txtHTML.Enabled = False
        pbColor.Enabled = False
    End If
    
    '' Load html to the custom html text box
    txtHTML.Text = server_html
    
    '' Set the background color box
    pbColor.BackColor = server_color
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '' Enable icon menu
    menu_cancel_flag = False
    
End Sub


Private Sub optRepBG_Click()
    
    pbColor.Enabled = True
    txtHTML.Enabled = False

End Sub

Private Sub optRepHTML_Click()

    pbColor.Enabled = False
    txtHTML.Enabled = True
    
End Sub

Private Sub optRepLogo_Click()

    pbColor.Enabled = False
    txtHTML.Enabled = False
    
End Sub



Private Sub pbColor_Click()
    
    '' Incase of an error, just ignore
    On Error Resume Next
    '' Open up a color selection box
    With cmnDlg
            .DialogTitle = build_title & " -- Select a color"
            .CancelError = True
            .ShowColor
            '' Check if they clicked cancel
            If Err Then Exit Sub
            '' Set the color
            pbColor.BackColor = .Color
    End With
    
End Sub

Private Sub tbStrip_Click()

    '' Display the correct picturebox depending on
    '' which tab is selected
    Select Case tbStrip.SelectedItem.Index
        Case 1:
            pbSlab0.Visible = True
            pbSlab1.Visible = False
        
        Case 2:
            pbSlab0.Visible = False
            pbSlab1.Visible = True
    End Select
End Sub
