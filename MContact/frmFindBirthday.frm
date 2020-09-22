VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFindBirthday 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Birthday Contact"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmFindBirthday.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3960
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6985
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
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4935
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   1
         CustomFormat    =   "dd MMMM"
         Format          =   20774915
         UpDown          =   -1  'True
         CurrentDate     =   39181
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   1
         CustomFormat    =   "dd MMMM"
         Format          =   20774915
         UpDown          =   -1  'True
         CurrentDate     =   39181
      End
      Begin VB.Label Label1 
         Caption         =   "Between :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "And :"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "Found : 0 Contacts"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5080
      Width           =   1350
   End
   Begin VB.Label lblFindWhat 
      AutoSize        =   -1  'True
      Caption         =   "Find Contacts' Birthdays"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmFindBirthday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim intSource As Integer
Private rs As ADODB.Recordset

Private Sub cmdFind_Click()
    Call ClearData
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
    MSFlexGrid1.ColWidth(1) = 3070
    MSFlexGrid1.Text = "Name "
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.ColWidth(2) = 1000
    MSFlexGrid1.Text = "Other " '"ÖÐÎÄ "
    'MSFlexGrid1.CellFontName = gstrFontName '"NSimSun"
    'MSFlexGrid1.CellFontSize = 9.75
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.ColWidth(3) = 1100
    MSFlexGrid1.Text = "DOB "
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.ColWidth(4) = 600
    MSFlexGrid1.Text = "Age "
End Sub

Private Sub ClearData()
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 5
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.FixedCols = 1
    Call ColumnHead
    lblCount.Caption = "Found : 0 Contacts"
End Sub

Private Sub ResetDate()
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1) 'DateSerial(Year(Date), 1, 1)
    dtpTo.Value = DateSerial(Year(Date), Month(Date) + 1, 1) 'DateSerial(Year(Date), 12, 31)
    lblFindWhat.Caption = "Find Contacts' Birthdays in year " & Year(Date)
End Sub

Private Sub RefreshData()
Dim strSQL As String
Dim R As Integer
Dim I As Integer
Dim mintDay As Integer
Dim mintMonth As Integer
Dim mintYear As Integer
Dim mintAge As Integer

On Error GoTo errMessage

    Set rs = New ADODB.Recordset
    strSQL = "SELECT ContactID, EnglishName, ChineseName, DOBDay, DOBMonth, DOBYear, Birthday FROM Contact" & _
        " WHERE Birthday BETWEEN" & _
        " DateSerial(" & Year(Date) & "," & Month(dtpFrom.Value) & "," & Day(dtpFrom.Value) & ") AND" & _
        " DateSerial(" & Year(Date) & "," & Month(dtpTo.Value) & "," & Day(dtpTo.Value) & ")" & _
        " ORDER BY Birthday"
    rs.Open strSQL, cn

    If Not rs.EOF Then
        While Not rs.EOF
            R = R + 1
            rs.MoveNext
        Wend
        lblCount.Caption = "Found : " & R & " Contacts"
        MSFlexGrid1.Rows = R + 1
        Call ColumnHead
        MSFlexGrid1.ColAlignment(3) = 6
        If R > 15 Then
            MSFlexGrid1.ColWidth(1) = 2810
        End If
        rs.MoveFirst
        While Not rs.EOF
            I = I + 1
            MSFlexGrid1.TextMatrix(I, 0) = rs!ContactID
            MSFlexGrid1.TextMatrix(I, 1) = rs!EnglishName
            MSFlexGrid1.TextMatrix(I, 2) = rs!ChineseName
            'MSFlexGrid1.CellFontName = gstrFontName '"NSimSun"
            'MSFlexGrid1.CellFontSize = 9.75
    
            'rs!DOBDay & " " & MonthName(rs!DOBMonth, True)
            'Format(rs!Birthday, "dd MMM")
            If rs!DOBYear <> "" Then
                If rs!DOBMonth <> "" Then
                    If rs!DOBDay <> "" Then
                        MSFlexGrid1.TextMatrix(I, 3) = Format(DateSerial(rs!DOBYear, rs!DOBMonth, rs!DOBDay), "dd MMM yyyy")
                    Else 'only has Month
                        MSFlexGrid1.TextMatrix(I, 3) = Format(DateSerial(rs!DOBYear, rs!DOBMonth, 1), "MMM yyyy") 'MonthName(rs!DOBMonth, True)
                    End If
                Else
                    MSFlexGrid1.TextMatrix(I, 3) = ""
                End If
            Else
                If rs!DOBMonth <> "" Then
                    If rs!DOBDay <> "" Then
                        MSFlexGrid1.TextMatrix(I, 3) = Format(DateSerial(1980, rs!DOBMonth, rs!DOBDay), "dd MMM")
                    Else 'only has Month
                        MSFlexGrid1.TextMatrix(I, 3) = MonthName(rs!DOBMonth, True)
                    End If
                Else
                    MSFlexGrid1.TextMatrix(I, 3) = ""
                End If
            End If
            
            'If IsNull(rs!Mobile1) Then
            '    MSFlexGrid1.TextMatrix(I, 4) = ""
            'Else
            '    MSFlexGrid1.TextMatrix(I, 4) = rs!Mobile1
            'End If
            
            mintYear = 0
            mintMonth = 0
            mintDay = 0
            mintAge = 0
            If rs!DOBYear <> "" Then
                If IsNumeric(rs!DOBYear) = True Then mintYear = CInt(rs!DOBYear)
            End If
            If rs!DOBMonth <> "" Then
                If IsNumeric(rs!DOBMonth) = True Then mintMonth = CInt(rs!DOBMonth)
                If rs!DOBDay <> "" Then
                    If IsNumeric(rs!DOBDay) = True Then mintDay = CInt(rs!DOBDay)
                Else
                    mintDay = 1
                End If
            End If
            mintAge = ComputeAge(mintDay, mintMonth, mintYear)
            If mintAge = 0 Then
                MSFlexGrid1.TextMatrix(I, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(I, 4) = mintAge
            End If
            rs.MoveNext
        Wend
    End If
    Exit Sub
errMessage:
    MsgBox Error, vbCritical, "RefreshData"
End Sub

Private Sub cmdReset_Click()
    ClearData
    ResetDate
End Sub

Private Sub dtpFrom_Change()
    dtpFrom.ToolTipText = "Year = " & Year(dtpFrom.Value)
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFind.SetFocus
    End If
End Sub

Private Sub dtpTo_Change()
    dtpTo.ToolTipText = "Year = " & Year(dtpTo.Value)
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFind.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ClearData
    ResetDate
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
