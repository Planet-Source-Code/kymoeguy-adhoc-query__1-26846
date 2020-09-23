VERSION 5.00
Begin VB.Form frmJoin 
   Caption         =   "Join"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Join"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim response

If List1.Text = List2.Text And Check1.Value = vbUnchecked Then
  frmQuery.Label3.Caption = Label1.Caption & "." & List1.Text & " = " & Label2.Caption & "." & List2.Text
ElseIf List1.Text = List2.Text And Check1.Value = vbChecked Then
  frmQuery.Label1(10).Caption = Label1.Caption & "." & List1.Text & " = " & Label2.Caption & "." & List2.Text
ElseIf List1.Text <> List2.Text Then
  response = MsgBox("The fields you have selected are not the same, do you wish to continue?", vbYesNo, "Verify Join")
    If response = vbNo Then
      Exit Sub
    ElseIf response = vbYes And Check1.Value = vbUnchecked Then
      frmQuery.Label3.Caption = Label1.Caption & "." & List1.Text & " = " & Label2.Caption & "." & List2.Text
    ElseIf response = vbYes And Check1.Value = vbChecked Then
      frmQuery.Label1(10).Caption = Label1.Caption & "." & List1.Text & " = " & Label2.Caption & "." & List2.Text
    End If
End If
If frmQuery.lstTableSelected.ListCount = 2 Then
  Me.Hide
ElseIf frmQuery.lstTableSelected.ListCount = 3 And Check1.Value = vbUnchecked Then
  Label1.Caption = vbNullString
  Label2.Caption = vbNullString
  List1.Clear
  List2.Clear
  Check1.Value = vbChecked
  Call tbljoin
ElseIf frmQuery.lstTableSelected.ListCount = 3 And Check1.Value = vbChecked Then
  Me.Hide
End If
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Public Function tbljoin()
Dim dbsFIELDNAMES As DAO.Database
Dim tdfTest As DAO.TableDef
Dim tdf2Test As DAO.TableDef
Dim fldloop As DAO.Field
Dim fldloop2 As DAO.Field
Dim rsp
Dim tlst, tlst2 As String

If frmQuery.lstTableSelected.ListCount = 2 Then
  Label1.Caption = frmQuery.lstTableSelected.List(0)
  Label2.Caption = frmQuery.lstTableSelected.List(1)
ElseIf frmQuery.lstTableSelected.ListCount = 3 Then
  rsp = InputBox("Please enter the number of the first table to use in the join:  1. " & frmQuery.lstTableSelected.List(0) & " OR 2. " & frmQuery.lstTableSelected.List(1) & " OR 3. " & frmQuery.lstTableSelected.List(2), "Select Table")
    If rsp = "1" Then
      Label1.Caption = frmQuery.lstTableSelected.List(0)
      tlst = frmQuery.lstTableSelected.List(1)
      tlst2 = frmQuery.lstTableSelected.List(2)
    ElseIf rsp = "2" Then
      Label1.Caption = frmQuery.lstTableSelected.List(1)
      tlst = frmQuery.lstTableSelected.List(0)
      tlst2 = frmQuery.lstTableSelected.List(2)
    ElseIf rsp = "3" Then
      Label1.Caption = frmQuery.lstTableSelected.List(2)
      tlst = frmQuery.lstTableSelected.List(0)
      tlst2 = frmQuery.lstTableSelected.List(1)
    End If
  rsp = InputBox("Please enter the number of the table to join " & Label1.Caption & ":  1. " & tlst & " OR 2. " & tlst2, "Select Table")
    If rsp = "1" Then
      Label2.Caption = tlst
    ElseIf rsp = "2" Then
      Label2.Caption = tlst2
    End If
End If

List1.Clear
List2.Clear

Set dbsFIELDNAMES = OpenDatabase(strfile)
Set tdfTest = dbsFIELDNAMES.TableDefs(Label1.Caption)

For Each fldloop In tdfTest.Fields
    
    List1.AddItem fldloop.Name
    
Next fldloop

Set tdf2Test = dbsFIELDNAMES.TableDefs(Label2.Caption)

For Each fldloop2 In tdf2Test.Fields

  List2.AddItem fldloop2.Name

Next fldloop2

End Function
