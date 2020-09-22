VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Contact"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4445
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7832
      _Version        =   393216
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimHei"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtSearchText 
      BeginProperty Font 
         Name            =   "SimHei"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "Found : 0 Contacts"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5130
      Width           =   1350
   End
   Begin VB.Label lblFindWhat 
      Caption         =   "Find Contacts :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim intSource As Integer
Dim rs As ADODB.Recordset

Private Sub cmdFind_Click()
    ClearData
    If Trim(txtSearchText.Text) = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    RefreshData
    Screen.MousePointer = vbDefault
End Sub

Private Sub ColumnHead()
    MSFlexGrid1.Clear
    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.Text = "Rec# "
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColWidth(1) = 2480
    MSFlexGrid1.Text = "Name "
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.ColWidth(2) = 1000
    MSFlexGrid1.Text = "Other " '"ÖÐÎÄ "
    'MSFlexGrid1.CellFontName = gstrFontName
    'MSFlexGrid1.CellFontSize = 9.75
        
    MSFlexGrid1.Col = 3
    MSFlexGrid1.ColWidth(3) = 1100
    MSFlexGrid1.Text = "DOB "
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.ColWidth(4) = 1200
    MSFlexGrid1.Text = "Mobile 1 "
End Sub

Private Sub ClearData()
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 5
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.FixedCols = 1
    Call ColumnHead
    lblCount.Caption = "Found : 0 Contacts"
End Sub

Private Sub RefreshData()
Dim strSQL As String
Dim strSearch As String
Dim R As Integer
Dim I As Integer

On Error GoTo errMessage
    strSearch = CheckString(Trim(txtSearchText.Text))
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM Contact" & _
            " WHERE EnglishName LIKE '%" & strSearch & "%'" & _
            " OR ChineseName LIKE '%" & strSearch & "%'"
            If IsNumeric(strSearch) Then
                strSQL = strSQL & " OR ContactID = " & CLng(strSearch)
            End If
            If IsDate(strSearch) Then
                strSQL = strSQL & " OR Birthday = #" & CDate(strSearch) & "#"
            End If
    strSQL = strSQL & _
            " OR Mobile1 = '" & strSearch & "'" & _
            " OR Mobile2 = '" & strSearch & "'" & _
            " OR Work1 = '" & strSearch & "'" & _
            " OR Work2 = '" & strSearch & "'" & _
            " OR Fax1 = '" & strSearch & "'" & _
            " OR Fax2 = '" & strSearch & "'" & _
            " OR Home1 = '" & strSearch & "'" & _
            " OR Home2 = '" & strSearch & "'" & _
            " OR Email LIKE '%" & strSearch & "%'" & _
            " OR Website LIKE '%" & strSearch & "%'" & _
            " OR Remark LIKE '%" & strSearch & "%'"
    rs.Open strSQL, cn
    
    If Not rs.EOF Then
        While Not rs.EOF
            R = R + 1
            rs.MoveNext
        Wend
        lblCount.Caption = "Found : " & R & " Contacts"
        MSFlexGrid1.Rows = R + 1
        ColumnHead
        If R > 17 Then MSFlexGrid1.ColWidth(1) = 2220
        rs.MoveFirst
        While Not rs.EOF
            I = I + 1
            MSFlexGrid1.TextMatrix(I, 0) = rs!ContactID
            MSFlexGrid1.TextMatrix(I, 1) = rs!EnglishName
            MSFlexGrid1.TextMatrix(I, 2) = rs!ChineseName
            
            If IsNull(rs!Birthday) Or Not IsDate(rs!Birthday) Then
                MSFlexGrid1.TextMatrix(I, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(I, 3) = Format(rs!Birthday, "dd Mmm")
            End If
            If IsNull(rs!Mobile1) Then
                MSFlexGrid1.TextMatrix(I, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(I, 4) = rs!Mobile1
            End If
            rs.MoveNext
        Wend
    End If
    Exit Sub
errMessage:
    MsgBox Error, vbExclamation, "RefreshData"
End Sub

Private Sub Form_Load()
    Call ClearData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmContact.Show
End Sub

Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Row > 0 Then
        If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) <> "" Then
            frmContact.LoadRecord MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
            Unload Me
        End If
    End If
End Sub

Private Sub txtSearchText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFind.SetFocus
    End If
End Sub
