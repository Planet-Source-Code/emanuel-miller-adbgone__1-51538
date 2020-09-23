VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHosts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmHosts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   2760
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHosts.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHosts.frx":1A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHosts.frx":1DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHosts.frx":21A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHosts.frx":256A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbEdit 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   979
      ButtonWidth     =   1217
      ButtonHeight    =   979
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstHosts 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "HOST"
         Object.Width           =   7832
      EndProperty
   End
End
Attribute VB_Name = "frmHosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function save_hosts_file()
    
    '' If we encounter an error, jump to the err_ender goto tag
    On Error GoTo err_ender
    '' Open the hosts_file for writing
    Open hosts_file For Output As #1
    '' Dump the header to the file first
    Print #1, pound_data & Chr(13) & Chr(10)
    
    With frmHosts.lstHosts
        '' Go through each item in the list, dumping one by one
        For x = 1 To .ListItems.Count
            Print #1, localhost & " " & .ListItems(x).Text
        Next x
    End With
    '' Close the file
    Close #1
    '' Return without an error
    save_hosts_file = 0
    Exit Function

err_ender:
    '' Return with an error
    save_hosts_file = 1
    
            
End Function

Private Sub Form_Load()
    
    '' Disable the icon menu
    menu_cancel_flag = True
    
End Sub

Private Sub Form_Resize()

    '' Resize functions, if we run into an error
    '' it is because the user sized the window
    '' to small, ignore it
    On Error Resume Next
    lstHosts.Width = Me.Width - 350
    lstHosts.Height = Me.Height - 1220
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '' Enable the icon menu and clear the
    '' lsthost list
    menu_cancel_flag = False
    lstHosts.ListItems.Clear
    
End Sub

Private Sub lstHosts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '' If right mouse button was clicked
    '' Popup a small menu and set the caption
    '' of the remove menu item to the current
    '' selection of the list
    With frmMenu
    If Button = 2 Then
        .menuERemove.Caption = "Remove " & lstHosts.SelectedItem & "..."
        Me.PopupMenu frmMenu.menuMain2
    End If
    End With
    
End Sub

Private Sub tbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)

    '' If we run into an error chances are it was
    '' because the user forgot to select an item
    '' before clicking remove, ignore it
    On Error Resume Next
    
    '' Decide what to do depending on what button
    '' was selected
    '' Load was selected
    If Button.Index = 1 Then
        
        '' Show an open dialog
        With cmnDlg
            '' Set the path to the application path
            .InitDir = App.Path
            '' Set default extention
            .DefaultExt = "*.txt"
            '' Set the title of the dialog
            .DialogTitle = build_title & " -- Load host list"
            '' Set the default filters
            .Filter = "Text Files|*.txt|"
            '' Select the main filter
            .FilterIndex = 1
            '' If cancel is selected, raise an error
            .CancelError = True
            '' Show the dialog
            .ShowOpen
            
            '' If cancel was selected, just return
            If Err Then
                Exit Sub
            End If
            
            '' Notify the user that we are loading
            '' the selected file
            Me.Caption = build_title & " -- Loading host file..."
            
            '' Temporary variable
            Dim input_line As String
            '' If we run into an error ignore it
            On Error Resume Next
            '' Make sure the file is closed
            Close #1
            '' Open the file for reading
            Open cmnDlg.FileName For Input As #1
                '' Loop until the end of file
                While Not EOF(1)
                    '' Read a line of input into
                    '' the temporary variable
                    Line Input #1, input_line
                    '' If the variable is not empty, add this item
                    If Not Len(input_line) = 0 Then
                        lstHosts.ListItems.Add , , input_line
                    End If
                Wend
            '' Close the file
            Close #1
            '' Clear the variable
            input_line = ""
            
            '' Update the dialog caption
            frmHosts.Caption = build_title & " -- Hosts (" & frmHosts.lstHosts.ListItems.Count & ")"
        End With
            
    '' Save was selected
    ElseIf Button.Index = 2 Then
    
        '' Disable the dialog to prevent user error
        Me.Enabled = False
        '' Notify the user of activity
        Me.Caption = build_title & " -- Saving hosts..."
        '' Attemp to save the file, if anything but 0 is returned
        '' Notify the user that an error occured while saving
        save_it = save_hosts_file
        If save_it = 0 Then
            Me.Enabled = True
        Else
            MsgBox "Could not save the hosts file.", vbCritical, build_title
            Me.Enabled = True
        End If
        '' Update the caption
        Me.Caption = build_title & " -- Hosts (" & lstHosts.ListItems.Count & ")"
        
        '' If protection is already enabled, re-enable it so the new hosts
        '' are loaded and used
        If frmMenu.menuProton.Checked = True Then
            enable_Protection
        End If
    
    '' Add was selected (index 3 was a spacer)
    ElseIf Button.Index = 4 Then
        
        '' Temporary variable
        Dim add_me As String
        '' Show an input dialog for data entry
        add_me = InputBox("Enter the host you wish to add.", build_title, "")
        '' If cancel is selected, return
        If Err <> 0 Or Len(add_me) = 0 Then Exit Sub
        '' Add this item to the list
        frmHosts.lstHosts.ListItems.Add , , add_me
        
    '' Remove was selected
    ElseIf Button.Index = 5 Then
        
        '' Remove the selected item
        With frmHosts.lstHosts
            .ListItems.Remove .SelectedItem.Index
        End With
        
        '' Update dialog caption with the current list count
        frmHosts.Caption = build_title & " -- Hosts (" & frmHosts.lstHosts.ListItems.Count & ")"
    
    '' Close was selected
    Else
        '' Unload dialog
        Unload Me
        
    End If
    
End Sub
