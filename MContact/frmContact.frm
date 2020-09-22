VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmContact 
   Caption         =   "MContact"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   Icon            =   "frmContact.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":3032
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":3D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":45E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":4EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":5B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":6474
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":6D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":7628
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":7942
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContact.frx":7C5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      ItemData        =   "frmContact.frx":8536
      Left            =   4200
      List            =   "frmContact.frx":8538
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6375
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cboDay 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtRemark 
         Height          =   915
         Left            =   960
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   3600
         Width           =   5175
      End
      Begin VB.TextBox txtWebsite 
         Height          =   315
         Left            =   960
         MaxLength       =   255
         TabIndex        =   32
         Top             =   3240
         Width           =   5175
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   960
         MaxLength       =   255
         TabIndex        =   30
         Top             =   2880
         Width           =   5175
      End
      Begin VB.TextBox txtTelHome2 
         Height          =   315
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   28
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtTelHome 
         Height          =   315
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   26
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtTelFax2 
         Height          =   315
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   23
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtTelFax 
         Height          =   315
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   21
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtTelOffice2 
         Height          =   315
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtTelOffice 
         Height          =   315
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtTelMobile2 
         Height          =   315
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtEnglishName 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   5
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox txtChineseName 
         Height          =   360
         IMEMode         =   1  'ON
         Left            =   960
         MaxLength       =   20
         TabIndex        =   7
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtTelMobile 
         Height          =   315
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Remark:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Website:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "(2)"
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "(1)"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "(2)"
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "(1)"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "(2)"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "(1)"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "(2)"
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "(1)"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Email:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Home No:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fax No:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Office No:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblEnglishName 
         Caption         =   "Name (E):"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblChineseName 
         Caption         =   "Name (O):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblTelMobile 
         Caption         =   "Mobile No:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblDOB 
         Caption         =   "D.O.B.:"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblRecord 
         Caption         =   "Record No:"
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblRecordNo 
         Alignment       =   1  'Right Justify
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Exit (Esc)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Clear (Ctrl+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save Contact (Ctrl+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DELETE"
            Object.ToolTipText     =   "Delete Contact (Ctrl+D)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Find Contact (Ctrl+F)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BIRTHDAY"
            Object.ToolTipText     =   "Birthday Contact (Ctrl+B)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIRST"
            Object.ToolTipText     =   "First Record (Ctrl+Home)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PREVIOUS"
            Object.ToolTipText     =   "Previous (Ctrl+Left)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEXT"
            Object.ToolTipText     =   "Next (Ctrl+Right)"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LAST"
            Object.ToolTipText     =   "Last Record (Ctrl+End)"
            ImageIndex      =   14
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Database: Microsoft Access 2000 Database (Password protected)
'File name: Contact.mdb
'Database Password: Load from GenWord function
'Set Project Reference to
'Microsoft ActiveX Data Objects 2.7 Library or newer
Option Explicit
Dim rstAll As ADODB.Recordset

Private Sub cboDay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboMonth.SetFocus
    End If
End Sub

Private Sub cboMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboYear.SetFocus
    End If
End Sub

Private Sub cboYear_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelMobile.SetFocus
        txtTelMobile.SelStart = Len(txtTelMobile.Text)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyE And (Shift And vbCtrlMask) Then
        If vbYes = MsgBox("Clear text fields?", vbQuestion + vbYesNo, App.Title) Then
            Clear
        End If
    ElseIf KeyCode = vbKeyS And (Shift And vbCtrlMask) Then
        SaveData
    ElseIf KeyCode = vbKeyD And (Shift And vbCtrlMask) Then
        DeleteData
    ElseIf KeyCode = vbKeyF And (Shift And vbCtrlMask) Then
        frmFind.Show
        gblnOpen = True
        Unload Me
    ElseIf KeyCode = vbKeyB And (Shift And vbCtrlMask) Then
        frmFindBirthday.Show
        gblnOpen = True
        Unload Me
    ElseIf KeyCode = vbKeyHome And (Shift And vbCtrlMask) Then
        If Not rstAll.BOF Then
            rstAll.MoveFirst
            LoadDetail
        End If
    ElseIf KeyCode = vbKeyLeft And (Shift And vbCtrlMask) Then
        If Not (rstAll.EOF Or rstAll.BOF) Then rstAll.MovePrevious
        If rstAll.BOF Then
            If Not (rstAll.EOF Or rstAll.BOF) Then rstAll.MoveLast
        End If
        LoadDetail
    ElseIf KeyCode = vbKeyRight And (Shift And vbCtrlMask) Then
        If Not (rstAll.EOF Or rstAll.BOF) Then rstAll.MoveNext
        If rstAll.EOF Then
            If Not (rstAll.EOF Or rstAll.BOF) Then rstAll.MoveFirst
        End If
        LoadDetail
    ElseIf KeyCode = vbKeyEnd And (Shift And vbCtrlMask) Then
        If Not rstAll.EOF Then
            rstAll.MoveLast
            LoadDetail
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & " by Aeric Poon"
Me.Caption = Me.Caption & " (" & App.Major & "." & App.Minor & " Built " & App.Revision & ")"
'On Error GoTo ShowError
    
    'Set cn = New ADODB.Connection
    'cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    'cn.ConnectionString = "Data Source=" & App.Path & "\MyDatabase.mdb"
    'cn.Properties("Jet OLEDB:Database Password") = "report"
    'cn.Open
    
    GenerateDateList
    ReadSettings
    SetOtherNameFont
    
    If gblnWithPassword = False Then Toolbar1.Buttons("BIRTHDAY").Enabled = False
    
    If gblnOpen = False Then
        If OpenDatabase = False Then
            MsgBox "Unable to open Database", vbCritical, "Open Database"
            End
        Else
            LoadAllData
        End If
    End If
    
    UpdateBirthday
'    Exit Sub
'ShowError:
'    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ShowError
    If gblnOpen = False Then 'If False means not opening other form
        CloseDatabase
        
'        If cn.State = adStateOpen Then
'            cn.Close
'        End If

        If FileExists(App.Path & "\Backup.MDB") Then
            Kill App.Path & "\Backup.MDB"
        End If
        If CompactDB(App.Path & "\" & gstrDatabasePath, App.Path & "\Backup.MDB") = False Then
            MsgBox Error, vbExclamation, App.Title
        End If
        
        Kill App.Path & "\" & gstrDatabasePath
        FileCopy App.Path & "\Backup.MDB", App.Path & "\" & gstrDatabasePath
        Kill App.Path & "\Backup.MDB" 'Optional
    Else
        UpdateBirthday
    End If
    Exit Sub
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "CLOSE"
        Unload Me
    Case "NEW"
        Clear
    Case "SAVE"
        SaveData
    Case "DELETE"
        DeleteData
    Case "FIND"
        'Find
        frmFind.Show
        gblnOpen = True
        Unload Me
    Case "BIRTHDAY"
        'Find birthday
        frmFindBirthday.Show
        gblnOpen = True
        Unload Me
    Case "FIRST"
        If Not rstAll.BOF Then
            rstAll.MoveFirst
            LoadDetail
        End If
    Case "PREVIOUS"
        If rstAll.EOF And rstAll.EOF Then Exit Sub
        'If Not (rstAll.BOF Or rstAll.RecordCount = 0) Then
            rstAll.MovePrevious
        'Else '
        If rstAll.BOF Then
            'If Not (rstAll.EOF Or rstAll.BOF) Then
            rstAll.MoveLast
        'Else
        '    rstAll.MovePrevious
        End If
        LoadDetail
    Case "NEXT"
        If rstAll.EOF And rstAll.EOF Then Exit Sub
        rstAll.MoveNext
        If rstAll.EOF Then
            'If Not (rstAll.EOF Or rstAll.BOF) Then
            rstAll.MoveFirst
        End If
        LoadDetail
    Case "LAST"
        If Not rstAll.EOF Then
            rstAll.MoveLast
            LoadDetail
        End If
End Select
End Sub

Private Sub txtChineseName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboDay.SetFocus
    End If
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtWebsite.SetFocus
        txtWebsite.SelStart = Len(txtWebsite.Text)
    End If
End Sub

Private Sub txtEnglishName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtChineseName.SetFocus
        txtChineseName.SelStart = Len(txtChineseName.Text)
    End If
End Sub

Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vbYes = MsgBox("Do you want to save?", vbQuestion + vbYesNo, App.Title) Then
            SaveData
        End If
        txtEnglishName.SetFocus
    End If
End Sub

Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelFax2.SetFocus
        txtTelFax2.SelStart = Len(txtTelFax2.Text)
    End If
End Sub

Private Sub txtTelFax2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelHome.SetFocus
        txtTelHome.SelStart = Len(txtTelHome.Text)
    End If
End Sub

Private Sub txtTelHome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelHome2.SetFocus
        txtTelHome2.SelStart = Len(txtTelHome2.Text)
    End If
End Sub

Private Sub txtTelHome2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtEmail.SetFocus
        txtEmail.SelStart = Len(txtEmail.Text)
    End If
End Sub

Private Sub txtTelMobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelMobile2.SetFocus
        txtTelMobile2.SelStart = Len(txtTelMobile2.Text)
    End If
End Sub

Private Sub LoadDetail()
If rstAll.BOF Or rstAll.EOF Then Exit Sub
Toolbar1.Buttons("DELETE").Enabled = True

On Error GoTo ShowError
    If Len(rstAll!ContactID) = 3 Then
        lblRecordNo.Caption = "0" & rstAll!ContactID
    ElseIf Len(rstAll!ContactID) = 2 Then
        lblRecordNo.Caption = "00" & rstAll!ContactID
    ElseIf Len(rstAll!ContactID) = 1 Then
        lblRecordNo.Caption = "000" & rstAll!ContactID
    Else
        lblRecordNo.Caption = rstAll!ContactID
    End If
    '==========English Name==========
        If rstAll!EnglishName <> "" Then
            txtEnglishName.Text = rstAll!EnglishName
        Else
            txtEnglishName.Text = ""
        End If
    '==========Chinese Name==========
        If rstAll!ChineseName <> "" Then
            txtChineseName.Text = rstAll!ChineseName
        Else
            txtChineseName.Text = ""
        End If
    '============DOB Month===========
        If rstAll!DOBMonth <> "" Then
            cboMonth.Text = MonthName(rstAll!DOBMonth)
        Else
            cboMonth.Text = "  "
        End If
    '=============DOB Day============
        If rstAll!DOBDay <> "" Then
            If CInt(rstAll!DOBDay) < 10 Then
                cboDay.Text = "0" & rstAll!DOBDay
            Else
                cboDay.Text = rstAll!DOBDay
            End If
        Else
            cboDay.Text = "  "
        End If
    '============DOB Year============
        If rstAll!DOBYear <> "" Then
            If CInt(rstAll!DOBYear) > Year(Now) Then
                MsgBox "D.O.B. year=" & CInt(rstAll!DOBYear) & " > System year=" & _
                Year(Now) & "!" & vbCrLf & "Please change System year!", vbExclamation, "LoadDetail"
                cboYear.Text = "  "
            Else
                cboYear.Text = rstAll!DOBYear
            End If
        Else
            cboYear.Text = "  "
        End If
    '==========Mobile Phone No==========
        If rstAll!Mobile1 <> "" Then
            txtTelMobile.Text = rstAll!Mobile1
        Else
            txtTelMobile.Text = ""
        End If
    '=========Mobile Phone No 2=========
        If rstAll!Mobile2 <> "" Then
            txtTelMobile2.Text = rstAll!Mobile2
        Else
            txtTelMobile2.Text = ""
        End If
        
    '==========Office Phone No==========
        If rstAll!Work1 <> "" Then
            txtTelOffice.Text = rstAll!Work1
        Else
            txtTelOffice.Text = ""
        End If
    '=========Office Phone No 2=========
        If rstAll!Work2 <> "" Then
            txtTelOffice2.Text = rstAll!Work2
        Else
            txtTelOffice2.Text = ""
        End If
    '===============Fax No==============
        If rstAll!Fax1 <> "" Then
            txtTelFax.Text = rstAll!Fax1
        Else
            txtTelFax.Text = ""
        End If
    '==============Fax No 2=============
        If rstAll!Fax2 <> "" Then
            txtTelFax2.Text = rstAll!Fax2
        Else
            txtTelFax2.Text = ""
        End If
    '==============Home No==============
        If rstAll!Home1 <> "" Then
            txtTelHome.Text = rstAll!Home1
        Else
            txtTelHome.Text = ""
        End If
    '=============Home No 2=============
        If rstAll!Home2 <> "" Then
            txtTelHome2.Text = rstAll!Home2
        Else
            txtTelHome2.Text = ""
        End If
    '===============Email===============
        If rstAll!Email <> "" Then
            txtEmail.Text = rstAll!Email
        Else
            txtEmail.Text = ""
        End If
    '==============Website==============
        If rstAll!Website <> "" Then
            txtWebsite.Text = rstAll!Website
        Else
            txtWebsite.Text = ""
        End If
    '===============Remark==============
        If rstAll!Remark <> "" Then
            txtRemark.Text = rstAll!Remark
        Else
            txtRemark.Text = ""
        End If
        txtEnglishName.SelStart = 0
        txtEnglishName.SelLength = Len(txtEnglishName)
    Exit Sub
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "LoadDetail"
End Sub

Private Sub SaveData()
Dim rstData As ADODB.Recordset  'Contact
Dim mlngContactID As Long

If Trim(txtEnglishName.Text) = "" Then
    MsgBox "Please enter English Name.", vbExclamation + vbOKOnly, "Required Field"
    txtEnglishName.SetFocus
    Exit Sub
End If
    
'If txtChineseName.Text = "" Then
'    MsgBox "Please enter Chinese Name.", vbExclamation + vbOKOnly, "Required Field"
'    txtChineseName.SetFocus
'    Exit Sub
'End If

'If Trim(txtRecordNo.Text) <> "" And IsNumeric(txtRecordNo) = False Then
'    MsgBox "Invalid Record No", vbExclamation, App.Title
'    Exit Sub
'End If

'Check Gregorian Calendar
If Trim(cboYear.Text) <> "" Then
    If CInt(cboYear.Text) < 1582 Then
        MsgBox "Please enter Year greater or equal to 1852.", vbExclamation + vbOKOnly, "Invalid Year"
        cboYear.SetFocus
        Exit Sub
    End If
End If

'On Error GoTo ShowError
       
    Set rstData = New ADODB.Recordset
'    Debug.Print rstAll!AutoNum
    
    If lblRecordNo.Caption = "0000" Then
        mlngContactID = GetMaxRecordNum + 1
    Else
        mlngContactID = CLng(lblRecordNo.Caption)
    End If
    
    rstData.Open "SELECT * FROM Contact WHERE ContactID = " & mlngContactID, cn, adOpenStatic, adLockOptimistic
    
    If rstData.EOF Then
        rstData.AddNew
        rstData!ContactID = mlngContactID
    'Else
    '    If rstData.EOF Then
    '        MsgBox "Contact ID not found", vbExclamation, App.Title
    '        Exit Sub
    '    End If
    End If
    
    rstData!EnglishName = Trim(txtEnglishName.Text)
    rstData!ChineseName = Trim(txtChineseName.Text)
    
    'Month
    If cboMonth.ListIndex > 0 Then
        rstData!DOBMonth = CInt(cboMonth.ListIndex)
    Else
        rstData!DOBMonth = Null
    End If
    'Day
    If cboDay.ListIndex > 0 Then
        If cboMonth.ListIndex > 0 And cboYear.ListIndex > 0 Then
            rstData!DOBDay = CorrectDay(CInt(cboDay.Text), CInt(cboMonth.ListIndex), CInt(cboYear.Text))
        Else
            rstData!DOBDay = CInt(cboDay.Text)
        End If
    Else
        rstData!DOBDay = Null
    End If
    'Year
    If cboYear.ListIndex > 0 Then
        rstData!DOBYear = CInt(cboYear.Text)
    Else
        rstData!DOBYear = Null
    End If
    
    rstData!Mobile1 = Trim(txtTelMobile.Text)
    rstData!Mobile2 = Trim(txtTelMobile2.Text)
    rstData!Work1 = Trim(txtTelOffice.Text)
    rstData!Work2 = Trim(txtTelOffice2.Text)
    rstData!Fax1 = Trim(txtTelFax.Text)
    rstData!Fax2 = Trim(txtTelFax2.Text)
    rstData!Home1 = Trim(txtTelHome.Text)
    rstData!Home2 = Trim(txtTelHome2.Text)
    rstData!Email = Trim(txtEmail.Text)
    rstData!Website = Trim(txtWebsite.Text)
    rstData!Remark = Trim(txtRemark.Text)
    'rstData!LastSaved = Format(Now, "d/M/yyyy h:mm:ss AMPM")
    rstData.Update
    
    UpdateBirthday
    rstAll.Requery
    'rstAll.MoveLast
    LoadRecord mlngContactID
    LoadDetail
    
    MsgBox "Contact successfully saved.", vbInformation + vbOKOnly, App.Title
'    LoadAllData
    Exit Sub
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "SaveData"
End Sub

Public Sub LoadAllData()
On Error GoTo ShowError
    Set rstAll = New ADODB.Recordset
    rstAll.Open "SELECT * FROM Contact ORDER BY ContactID", cn, adOpenForwardOnly, adLockOptimistic
    
    If Not rstAll.EOF Then 'Or rstContact.RecordCount = 0 Then
        rstAll.MoveFirst
        LoadDetail
    Else
        Clear
    End If
    Exit Sub
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "LoadAllData"
End Sub

Private Sub Clear()
txtEnglishName.Text = ""
txtChineseName.Text = ""
lblRecordNo.Caption = "0000"

cboDay.ListIndex = 0
cboMonth.ListIndex = 0
cboYear.ListIndex = 0

txtTelMobile.Text = ""
txtTelMobile2.Text = ""
txtTelOffice.Text = ""
txtTelOffice2.Text = ""
txtTelFax.Text = ""
txtTelFax2.Text = ""
txtTelHome.Text = ""
txtTelHome2.Text = ""
txtEmail.Text = ""
txtWebsite.Text = ""
txtRemark.Text = ""

'Toolbar1.Buttons("DELETE").Enabled = False
End Sub

Private Sub SetOtherNameFont()
On Error GoTo ShowError
    If gstrFontName = "SimHei" Or gstrFontName = "SimSun" Then
        txtChineseName.Font.Name = gstrFontName
        txtChineseName.Font.Size = 12
        txtChineseName.Font.Charset = 134
    Else
        txtChineseName.Font = "MS Sans Serif"
        txtChineseName.FontSize = 8
    End If
    Exit Sub
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "SetOtherNameFont"
End Sub

Private Sub DeleteData()
    If lblRecordNo.Caption = "0000" Then
        MsgBox "Cannot Delete Empty Record.", vbExclamation, App.Title
        Exit Sub
    End If
    If Not MsgBox("Delete Contact #" & lblRecordNo.Caption & " from Database?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Exit Sub
    End If
On Error GoTo ShowError
    rstAll.Delete
    Clear
    MsgBox "Contact is deleted.", vbInformation + vbOKOnly, App.Title
    rstAll.Requery
    If Not (rstAll.EOF Or rstAll.BOF) Then rstAll.MoveLast
    LoadDetail
    Exit Sub
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "DeleteData"
End Sub

Private Function GetMaxRecordNum() As Long
    Dim rst As ADODB.Recordset
    Dim strSQL As String

On Error GoTo ShowError
    Set rst = New ADODB.Recordset
    strSQL = "SELECT Max(ContactID)AS MaxNo FROM [Contact]"
    rst.Open strSQL, cn, adOpenForwardOnly, adLockOptimistic
    If rst!MaxNo <> "" Then
        GetMaxRecordNum = CLng(rst!MaxNo)
    Else
        GetMaxRecordNum = 0
    End If

    rst.Close
    Set rst = Nothing
    Exit Function
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "GetMaxRecordNum"
End Function

Private Sub txtTelMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelOffice.SetFocus
        txtTelOffice.SelStart = Len(txtTelOffice.Text)
    End If
End Sub

Private Sub txtTelOffice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelOffice2.SetFocus
        txtTelOffice2.SelStart = Len(txtTelOffice2.Text)
    End If
End Sub

Private Sub txtTelOffice2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTelFax.SetFocus
        txtTelFax.SelStart = Len(txtTelFax.Text)
    End If
End Sub

Private Sub txtWebsite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRemark.SetFocus
        txtRemark.SelStart = Len(txtRemark.Text)
    End If
End Sub

Public Sub LoadRecord(lngContactID As Long)
On Error GoTo ShowError
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    
    Set rst = New ADODB.Recordset
    strSQL = "SELECT ContactID, EnglishName, ChineseName," & _
            " DOBMonth, DOBDay, DOBYear, Mobile1, Mobile2" & _
            " FROM Contact WHERE ContactID = " & lngContactID
    rst.Open strSQL, cn, adOpenStatic, adLockReadOnly
    
    Clear
    If Not rst.EOF Then
        'rstAll.Find "ContactID = " & rst!ContactID
        'rstAll!ContactID = rst!ContactID
        rstAll.MoveFirst
        Do Until rstAll!ContactID = rst!ContactID
            rstAll.MoveNext
        Loop
        LoadDetail
    Else
        MsgBox "Contact not found", vbInformation
    End If

    'txtEnglishName.SelStart = Len(txtName.Text)
    'txtEnglishName.SetFocus
    Toolbar1.Buttons("DELETE").Enabled = True
    
    rst.Close
    Set rst = Nothing
    Exit Sub
ShowError:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "LoadRecord"
End Sub

Private Sub GenerateDateList()
Dim y As Integer
Dim m As Integer
Dim d As Integer

Dim t As Integer

'Year
cboYear.AddItem "  "
If Year(Now) - 99 < 1582 Or Year(Now) > 9999 Then
    t = 1901 'Default Year
    For y = 2000 To t Step -1
        cboYear.AddItem y
    Next
Else
    t = Year(Now) - 99
    For y = Year(Now) To t Step -1
        cboYear.AddItem y
    Next
End If
'Month
cboMonth.AddItem "  "
For m = 1 To 12
cboMonth.AddItem MonthName(m)
Next
'Day
cboDay.AddItem "  "
For d = 1 To 31
If d < 10 Then
cboDay.AddItem "0" & d
Else
cboDay.AddItem d
End If
Next
End Sub

Private Function CompactDB(strOriginalFileName As String, strDestinationFileName As String) As Boolean
Dim oJetEngine As New JRO.JetEngine

On Error GoTo errHandle

Dim strSource As String
Dim strDestination As String

strSource = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & strOriginalFileName & ";" & _
            "Jet Oledb:Database Password=" & gstrPassword & ";" & _
            "Jet OLEDB:Engine Type=5;"
            
strDestination = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & strDestinationFileName & ";" & _
          "Jet Oledb:Database Password=" & gstrPassword & ";" & _
          "Jet OLEDB:Engine Type=5;"

oJetEngine.CompactDatabase strSource, strDestination
Set oJetEngine = Nothing
CompactDB = True
Exit Function
errHandle:
    MsgBox "Error#: " & Err.Number & vbCrLf & Err.Description, vbExclamation, "CompactDB"
    Set oJetEngine = Nothing
    CompactDB = False
End Function
