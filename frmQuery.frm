VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQuery 
   Caption         =   "Ad Hoc Query"
   ClientHeight    =   10275
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14325
   ControlBox      =   0   'False
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   14325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   12000
      Picture         =   "frmQuery.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   9120
      Width           =   1935
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   9
      Left            =   11160
      TabIndex        =   42
      Top             =   6000
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   8
      Left            =   11160
      TabIndex        =   41
      Top             =   5640
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   7
      Left            =   11160
      TabIndex        =   40
      Top             =   5280
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   6
      Left            =   11160
      TabIndex        =   39
      Top             =   4920
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   5
      Left            =   11160
      TabIndex        =   38
      Top             =   4560
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   4
      Left            =   11160
      TabIndex        =   37
      Top             =   4200
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   3
      Left            =   11160
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   2
      Left            =   11160
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Index           =   1
      Left            =   11160
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   9
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   6000
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   8
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   5640
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   7
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   5280
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   6
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   5
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   4
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   3
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   2
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cmboOperator 
      Height          =   315
      Index           =   1
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox txtCriteriaField 
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      Height          =   10185
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8520
         Top             =   9480
      End
      Begin VB.ListBox lstfld 
         Height          =   2790
         Left            =   8640
         TabIndex        =   61
         Top             =   6600
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Frame Frame10 
         Height          =   1875
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   6585
         Begin VB.ListBox lstTables 
            Appearance      =   0  'Flat
            DragIcon        =   "frmQuery.frx":0884
            Height          =   1200
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   510
            Width           =   2955
         End
         Begin VB.ListBox lstTableSelected 
            Appearance      =   0  'Flat
            DragIcon        =   "frmQuery.frx":0CC6
            Height          =   1200
            Left            =   3240
            TabIndex        =   57
            Top             =   510
            Width           =   3195
         End
         Begin VB.Label Label15 
            Caption         =   "Available Tables"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   210
            Width           =   1635
         End
         Begin VB.Label Label16 
            Caption         =   "Selected Table"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3270
            TabIndex        =   59
            Top             =   180
            Width           =   1305
         End
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   9
         Left            =   5520
         TabIndex        =   24
         Top             =   6000
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   8
         Left            =   5520
         TabIndex        =   23
         Top             =   5640
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   7
         Left            =   5520
         TabIndex        =   22
         Top             =   5280
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   6
         Left            =   5520
         TabIndex        =   21
         Top             =   4920
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   5
         Left            =   5520
         TabIndex        =   20
         Top             =   4560
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   4
         Left            =   5520
         TabIndex        =   19
         Top             =   4200
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   3
         Left            =   5520
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   3795
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   7320
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   13770
         Top             =   1860
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Choose the access database"
         Filter          =   "Access Database(*.mdb)|*.mdb"
         InitDir         =   "c:\"
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "RUN QUERY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4680
         TabIndex        =   10
         Top             =   6720
         Width           =   2535
      End
      Begin VB.TextBox txtCriteria 
         Height          =   285
         Index           =   0
         Left            =   11220
         TabIndex        =   9
         Top             =   2760
         Width           =   2385
      End
      Begin VB.ComboBox cmboOperator 
         Height          =   315
         Index           =   0
         Left            =   9540
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2760
         Width           =   1545
      End
      Begin VB.TextBox txtCriteriaField 
         Height          =   285
         Index           =   0
         Left            =   5460
         TabIndex        =   7
         Top             =   2760
         Width           =   3795
      End
      Begin VB.ListBox lstCriteria 
         Appearance      =   0  'Flat
         DragIcon        =   "frmQuery.frx":1108
         Height          =   3345
         Left            =   240
         TabIndex        =   6
         Top             =   2880
         Width           =   4155
      End
      Begin VB.Frame Frame9 
         Height          =   1875
         Left            =   6840
         TabIndex        =   1
         Top             =   120
         Width           =   7335
         Begin VB.ListBox lstFields 
            Appearance      =   0  'Flat
            DragIcon        =   "frmQuery.frx":154A
            Height          =   1200
            Left            =   120
            TabIndex        =   3
            Top             =   510
            Width           =   3315
         End
         Begin VB.ListBox lstSelected 
            Appearance      =   0  'Flat
            DragIcon        =   "frmQuery.frx":198C
            Height          =   1200
            Left            =   3480
            TabIndex        =   2
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label8 
            Caption         =   "Available fields"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Selected fields"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3510
            TabIndex        =   4
            Top             =   210
            Width           =   1305
         End
      End
      Begin MSDataGridLib.DataGrid grdQuery 
         Bindings        =   "frmQuery.frx":1DCE
         Height          =   1905
         Left            =   120
         TabIndex        =   11
         Top             =   7740
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   3360
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   735
         Left            =   600
         TabIndex        =   55
         Top             =   4920
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Select Multiple Tables"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   54
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   43
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Criteria Value"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11340
         TabIndex        =   15
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Selection Criteria"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9540
         TabIndex        =   14
         Top             =   2460
         Width           =   1485
      End
      Begin VB.Label Label17 
         Caption         =   "Selected Fields"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Top             =   2550
         Width           =   1725
      End
      Begin VB.Label lblDrag 
         Height          =   285
         Left            =   3060
         TabIndex        =   12
         Top             =   2400
         Width           =   1275
      End
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   10
      Left            =   4440
      TabIndex        =   62
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   52
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   51
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   50
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   49
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   47
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   46
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   45
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   44
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New Query"
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSQLQuery As String
Private WithEvents objADODC As Adodc
Attribute objADODC.VB_VarHelpID = -1
Dim lstcnt As Integer
Dim qcnt As Integer
Dim actv As Integer
Dim rsp

Sub modReformatForm()
Dim cc As Integer
Dim critcnt As Integer

'sub to clear the appropriate controls/
lstFields.Clear
lstfld.Clear
lstTables.Clear
lstTableSelected.Clear
lstSelected.Clear
grdQuery.ClearFields
grdQuery.Visible = False
Label3.Caption = vbNullString
Label1(10).Caption = vbNullString
Check1.Value = vbUnchecked
With frmJoin
  .Label1.Caption = vbNullString
  .Label2.Caption = vbNullString
  .List1.Clear
  .List2.Clear
  .Check1.Value = vbUnchecked
End With
For critcnt = 1 To 9
  txtCriteriaField(critcnt).Visible = False
  cmboOperator(critcnt).Visible = False
  txtCriteria(critcnt).Visible = False
Next critcnt
lstcnt = 0
strfile = ""
strSQLQuery = ""
 lstTables.Enabled = True
DoEvents
For cc = 0 To qcnt
  txtCriteriaField(cc).Text = ""
  cmboOperator(cc).Clear
  txtCriteria(cc).Text = ""
Next cc
grdQuery.Enabled = True
lstCriteria.Clear
lstCriteria.Enabled = True

On Error Resume Next
X = Controls.Remove(objADODC) 'remove control added onthe fly
grdQuery.ClearFields
On Error GoTo 0
DoEvents
End Sub


Private Sub modcreateQuery()
Dim strstrSqlString As String
Dim strTableName As String
Dim strTableQuery As String
Dim strConditions As String
Dim Criteria As String
Dim i As Integer
Dim strfieldtype As String
Dim q As Integer

strConditions = ""
Criteria = ""

For q = 0 To qcnt
  strfieldtype = Label1(q).Caption
    
    If strfieldtype = "DAT" Or strfieldtype = "STR" Then
      Criteria = "'" & txtCriteria(q).Text & "'"
    Else
      Criteria = txtCriteria(q).Text
    End If
    
    If q = 0 Then
      strConditions = lstTableSelected.Text & "." & txtCriteriaField(q).Text & Space(1) & cmboOperator(q) & Space(1) & Criteria
    Else
      strConditions = strConditions & " AND " & lstTableSelected.Text & "." & txtCriteriaField(q).Text & Space(1) & cmboOperator(q) & Space(1) & Criteria
    End If
Next q
strTableQuery = ""
strSQLQuery = ""

'loop through list box to get selected fields
For i = 0 To lstSelected.ListCount - 1
        strTableQuery = strTableQuery & lstSelected.List(i) & ","
Next i

If Left$(strTableQuery, 1) = "," Then strTableQuery = Right$(strTableQuery, Len(strTableQuery) - 1)
If Right$(strTableQuery, 1) = "," Then strTableQuery = Left$(strTableQuery, Len(strTableQuery) - 1)
'strConditions = txtCriteriaField(0).Text & " " & cmboOperator(0).Text & " " & "'" & txtCriteria(0).Text & "'"
If txtCriteriaField(0).Text = "" And lstTableSelected.ListCount = 1 Then
  strSQLQuery = "SELECT " & strTableQuery & " FROM " & lstTableSelected.Text
ElseIf txtCriteriaField(0).Text = "" And lstTableSelected.ListCount = 2 Then
  strSQLQuery = "SELECT " & strTableQuery & " FROM " & lstTableSelected.List(0) & ", " & lstTableSelected.List(1) & " WHERE " & Label3.Caption
ElseIf txtCriteriaField(0).Text = "" And lstTableSelected.ListCount = 3 Then
  strSQLQuery = "SELECT " & strTableQuery & " FROM " & lstTableSelected.List(0) & ", " & lstTableSelected.List(1) & ", " & lstTableSelected.List(2) & " WHERE " & Label3.Caption & " AND " & Label1(10).Caption
ElseIf txtCriteriaField(0).Text <> "" And lstTableSelected.ListCount = 1 Then
  strSQLQuery = "SELECT " & strTableQuery & " FROM " & lstTableSelected.Text & " WHERE" & Space(1) & strConditions
ElseIf txtCriteriaField(0).Text <> "" And lstTableSelected.ListCount = 2 Then
  strSQLQuery = "SELECT " & strTableQuery & " FROM " & lstTableSelected.List(0) & ", " & lstTableSelected.List(1) & " WHERE" & Space(1) & Label3.Caption & " AND " & strConditions
ElseIf txtCriteriaField(0).Text <> "" And lstTableSelected.ListCount = 3 Then
  strSQLQuery = "SELECT " & strTableQuery & " FROM " & lstTableSelected.List(0) & ", " & lstTableSelected.List(1) & "' " & lstTableSelected.List(2) & " WHERE" & Space(1) & Label3.Caption & " AND " & Label1(10).Caption & " AND " & strConditions
End If
'MsgBox strSQLQuery
End Sub

Private Sub Check1_Click()
If lstTables.ListCount = 0 Then
  Exit Sub
Else
rsp = InputBox("Enter the number of tables to usewith this query (maximium of 3): ", "Table Joins")
  If Int(rsp) > 3 Or Int(rsp) < 1 Then
    MsgBox "Please enter either 2 or 3 for the number of tables to use.", vbInformation, "Table Join Error"
    rsp = vbNullString
  Else
  End If
End If
End Sub

Private Sub cmdRun_Click()
Call modcreateQuery
DoEvents

Select Case actv
  Case 0
    Set objADODC = Controls.Add("MSAdodcLib.adodc", "adodcQuery1", Frame1)
  Case 1
    Set objADODC = Controls.Add("MSAdodcLib.adodc", "adodcQuery2", Frame1)
  Case 2
    Set objADODC = Controls.Add("MSAdodcLib.adodc", "adodcQuery3", Frame1)
  Case 3
    Set objADODC = Controls.Add("MSAdodcLib.adodc", "adodcQuery4", Frame1)
  Case 4
    Set objADODC = Controls.Add("MSAdodcLib.adodc", "adodcQuery5", Frame1)
  Case Else
    'MsgBox "Maximum allowed changes has been reached, please start again."
End Select
actv = actv + 1

With objADODC     'format the instantiated command button object
            .Width = 7785
            .Height = 330
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & strfile
            .Left = 120
            .Caption = "Data"
            .Top = 5100
            .Visible = True
            .Enabled = True
            .RecordSource = strSQLQuery
            .Visible = False
            DoEvents
            .Refresh
End With

Set grdQuery.DataSource = objADODC
grdQuery.Refresh
grdQuery.Visible = True
grdQuery.Enabled = True
End Sub

Private Sub Command1_Click()
  frmQuery.Cls
  frmQuery.Timer1.Enabled = True
End Sub

Private Sub Form_Load()

Call loadDAOData(frmQuery)
lstcnt = 0
actv = 0
Check1.Value = vbUnchecked
Label3.Caption = vbNullString
Label1(10).Caption = vbNullString

End Sub


Private Sub lstCriteria_DblClick()
'txtCriteriaField = lstCriteria.Text
'lstCriteria.Enabled = False

Dim dbsFIELDNAMES As DAO.Database
Dim tdfTest As DAO.TableDef
Dim fldloop As DAO.Field
Dim X As Integer
Dim scnt As Integer
Dim rsp

lstcnt = lstcnt + 1

Select Case lstcnt
  Case 1
    txtCriteriaField(0).Text = lstCriteria.Text
  Case 2
    txtCriteriaField(1).Text = lstCriteria.Text
    txtCriteriaField(1).Visible = True
    txtCriteria(1).Visible = True
    cmboOperator(1).Visible = True
  Case 3
    txtCriteriaField(2).Text = lstCriteria.Text
    txtCriteriaField(2).Visible = True
    txtCriteria(2).Visible = True
    cmboOperator(2).Visible = True
  Case 4
    txtCriteriaField(3).Text = lstCriteria.Text
    txtCriteriaField(3).Visible = True
    txtCriteria(3).Visible = True
    cmboOperator(3).Visible = True
  Case 5
    txtCriteriaField(4).Text = lstCriteria.Text
    txtCriteriaField(4).Visible = True
    txtCriteria(4).Visible = True
    cmboOperator(4).Visible = True
  Case 6
    txtCriteriaField(5).Text = lstCriteria.Text
    txtCriteriaField(5).Visible = True
    txtCriteria(5).Visible = True
    cmboOperator(5).Visible = True
  Case 7
    txtCriteriaField(6).Text = lstCriteria.Text
    txtCriteriaField(6).Visible = True
    txtCriteria(6).Visible = True
    cmboOperator(6).Visible = True
  Case 8
    txtCriteriaField(7).Text = lstCriteria.Text
    txtCriteriaField(7).Visible = True
    txtCriteria(7).Visible = True
    cmboOperator(7).Visible = True
  Case 9
    txtCriteriaField(8).Text = lstCriteria.Text
    txtCriteriaField(8).Visible = True
    txtCriteria(8).Visible = True
    cmboOperator(8).Visible = True
  Case 10
    txtCriteriaField(9).Text = lstCriteria.Text
    txtCriteriaField(9).Visible = True
    txtCriteria(9).Visible = True
    cmboOperator(9).Visible = True
  Case 11
    MsgBox "Sorry, 10 is the maximum query parameters."
    lstcnt = 10
End Select
'sub to add a new datagrid control, add a ado data control, bind the grid
'to the data control and then format the grid
Set dbsFIELDNAMES = OpenDatabase(strfile)
If lstTableSelected.ListCount = 1 Then
  Set tdfTest = dbsFIELDNAMES.TableDefs(lstTableSelected.Text)
ElseIf lstTableSelected.ListCount = 2 Then
  rsp = InputBox("Please select the number corresponding to the table to use: #1:  " & lstTableSelected.List(0) & " OR #2:  " & lstTableSelected.List(1), "Select Table")
    If rsp = "1" Then
      lstTableSelected.ListIndex = 0
    ElseIf rsp = "2" Then
      lstTableSelected.ListIndex = 1
    End If
  Set tdfTest = dbsFIELDNAMES.TableDefs(lstTableSelected.Text)
End If
For Each fldloop In tdfTest.Fields
  For scnt = (lstcnt - 1) To 9
    If fldloop.Name = Trim(txtCriteriaField(scnt)) Then
        Select Case fldloop.Type
                Case 3
                    strfieldtype = "INT"
                    Label1(scnt).Caption = "INT"
                    With cmboOperator(scnt)
                      .Clear
                      .AddItem " < "
                      .AddItem " > "
                      .AddItem " <> "
                      .AddItem " = "
                      .AddItem " <= "
                      .AddItem " >= "
                    End With
                Case 4
                    strfieldtype = "INT"
                    Label1(scnt).Caption = "INT"
                    With cmboOperator(scnt)
                      .Clear
                      .AddItem " < "
                      .AddItem " > "
                      .AddItem " <> "
                      .AddItem " = "
                      .AddItem " <= "
                      .AddItem " >= "
                    End With
                Case 5
                    strfieldtype = "CUR"
                    Label1(scnt).Caption = "CUR"
                    With cmboOperator(scnt)
                      .Clear
                      .AddItem " < "
                      .AddItem " > "
                      .AddItem " <> "
                      .AddItem " = "
                      .AddItem " <= "
                      .AddItem " >= "
                    End With
                Case 7
                    strfieldtype = "FLT"
                    Label1(scnt).Caption = "FLT"
                    With cmboOperator(scnt)
                      .Clear
                      .AddItem " < "
                      .AddItem " > "
                      .AddItem " <> "
                      .AddItem " = "
                      .AddItem " <= "
                      .AddItem " >= "
                    End With
                Case 8
                    strfieldtype = "DAT"
                    Label1(scnt).Caption = "DAT"
                    With cmboOperator(scnt)
                      .Clear
                      .AddItem " < "
                      .AddItem " > "
                      .AddItem " <> "
                      .AddItem " = "
                      .AddItem " <= "
                      .AddItem " >= "
                    End With
                Case 10
                    strfieldtype = "STR"
                    Label1(scnt).Caption = "STR"
                    With cmboOperator(scnt)
                      .Clear
                      .AddItem " <> "
                      .AddItem " = "
                      .AddItem " LIKE "
                    End With
                Case 12
                    strfieldtype = "MEM"
                    Label1(scnt).Caption = "MEM"
                    With cmboOperator(scnt)
                      .Clear
                      .AddItem " <> "
                      .AddItem " = "
                      .AddItem " LIKE "
                    End With
        End Select
    End If
 Next scnt
Next fldloop
qcnt = lstcnt - 1

End Sub


Private Sub lstFields_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DY   ' Declare variable.
   DY = TextHeight("A")   ' Get height of one line.
   lblDrag.Move (Frame9.Left + lstFields.Left), lstFields.Top + Y + DY, lstFields.Width, DY
   lblDrag.Drag   ' Drag label outline.
End Sub


Private Sub lstSelected_DragDrop(Source As Control, X As Single, Y As Single)
  If lstTableSelected.ListCount = 1 Then
    lstSelected.AddItem lstTableSelected.Text & "." & lstFields.Text
    lstCriteria.AddItem lstFields.Text
  ElseIf lstTableSelected.ListCount = 2 Or lstTableSelected.ListCount = 3 Then
    lstSelected.AddItem lstFields.Text
    lstfld.ListIndex = lstFields.ListIndex
    lstCriteria.AddItem lstfld.Text
    lstfld.RemoveItem lstfld.ListIndex
  End If
    lstFields.RemoveItem lstFields.ListIndex
End Sub


Private Sub lstTables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim DY
DY = TextHeight("A")
lblDrag.Move lstTables.Left, lstTables.Top + Y + DY, lstTables.Width, DY
lblDrag.Drag
End Sub


Private Sub lstTableSelected_Click()
Dim dbsFIELDNAMES As DAO.Database
Dim tdfTest As DAO.TableDef
Dim fldloop As DAO.Field
Dim X As Integer

If lstTableSelected.ListCount = 1 Then
  lstFields.Clear
  lstCriteria.Clear
Else
End If

Set dbsFIELDNAMES = OpenDatabase(strfile)
Set tdfTest = dbsFIELDNAMES.TableDefs(lstTableSelected.Text)

For Each fldloop In tdfTest.Fields
  If lstTableSelected.ListCount = 1 Then
    lstFields.AddItem fldloop.Name
  ElseIf lstTableSelected.ListCount = 2 Or lstTableSelected.ListCount = 3 Then
    lstFields.AddItem lstTableSelected.Text & "." & fldloop.Name
    lstfld.AddItem fldloop.Name
  End If
    
Next fldloop
End Sub


Private Sub lstTableSelected_DragDrop(Source As Control, X As Single, Y As Single)


 lstTableSelected.AddItem lstTables.Text
  If Check1.Value = vbUnchecked Then
   lstTables.Enabled = False
 ElseIf Check1.Value = vbChecked Then
   If lstTableSelected.ListCount = Int(rsp) Then
     lstTables.Enabled = False
   Else
   End If
End If

If lstTables.Enabled = False And lstTableSelected.ListCount > 1 Then
   frmJoin.Show
   Call frmJoin.tbljoin
Else
End If

End Sub


Private Sub mnuNew_Click()
modReformatForm
Call loadDAOData(frmQuery)
End Sub


Private Sub Timer1_Timer()
  frmQuery.Width = frmQuery.Width - 120 'make the form smaller
  frmQuery.Height = frmQuery.Height - 120
  If frmQuery.Width < 2800 Then End 'until it is small then exit the program
End Sub
