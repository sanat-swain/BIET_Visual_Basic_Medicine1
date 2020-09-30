VERSION 5.00
Begin VB.Form frmsupplier 
   BackColor       =   &H00400000&
   Caption         =   "SUPPLIER"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   10560
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Height          =   1575
      Left            =   2040
      TabIndex        =   7
      Top             =   4680
      Width           =   5775
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton CMDFIND 
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton CMDDEL 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdNEXT 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CMDADD 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Text            =   " "
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   4800
      TabIndex        =   5
      Text            =   " "
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Text            =   " "
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Caption         =   "SUPPLIER ADDRESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      Caption         =   "SUPPLIER NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   "SUPPLIER NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "SUPPLIER DETAILS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   4050
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim CMD As ADODB.Command

Private Sub CMDADD_Click()
clear
generatesno
Text2.SetFocus
cmdSAVE.Enabled = True
cmdADD.Enabled = False
End Sub

Private Sub CMDDEL_Click()
Dim i As String
i = "Are you sure to delete?"
If MsgBox(i, vbYesNo, "warning?") = vbYes Then
Text2.Text = ""
Text3.Text = ""
 Set CMD = New ADODB.Command
    'cmd.CommandText = "delete from Party_Master where Party_Code='" & Trim(txtICode.Text) & "' "
    CMD.CommandText = "Update dealer set sname='" & Trim(Text2.Text) & "',Address='" & Trim(Text3.Text) & "'  where sno like '" & Trim(Text1.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
MsgBox "Successfully deleted", vbExclamation, "Deletion"
clear
ElseIf MsgBox(i, vbYesNo, "Warning?") = vbNo Then
MsgBox "Do you want to exit", vbOKOnly, "Stop"
End If
Set rs = Nothing
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub CMDFIND_Click()
Dim i As String
i = InputBox("Enter The sno U want to find:")
clear
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select * from dealer where  sno like '" & i & "'"
rs.ActiveConnection = cn
rs.Open
If rs.EOF Then
MsgBox ("The PARTY With This Code is Not Exist")
CMDDEL.Enabled = False
Else
clear
load
End If
CMDDEL.Enabled = True
cmdUPDATE.Enabled = True
cmdADD.Enabled = True
rs.Close
End Sub

Private Sub CMDSAVE_Click()
If Text1.Text = "" Or Text2.Text = "" Then
  MsgBox "You Should Fill Code And Item NAME "
 Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "dealer"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       Call assign
       rs.Update
       rs.Close
       Set rs = Nothing
       Call clear
       End If
End If
Text2.SetFocus
cmdADD.Enabled = True
cmdSAVE.Enabled = False
End Sub

Private Sub CMDUPDATE_Click()
Set CMD = New ADODB.Command
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
    CMD.CommandText = "update dealer set sname='" & Trim(Text2.Text) & "', Address='" & Trim(Text3.Text) & "'where sno ='" & Trim(Text1.Text) & "'"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    MsgBox ("The Party is Successfully Modified")
    Set CMD = Nothing
    clear
    CMDDEL.Enabled = False
    cmdUPDATE.Enabled = False
    cmdSAVE.Enabled = False
    cmdADD.Enabled = True
    
End Sub

Private Sub DataGrid1_Click()
DisplayData
'RS.Refresh
cmdADD.Enabled = False
cmdSAVE.Enabled = True
CMDDEL.Enabled = True
cmdUPDATE.Enabled = True
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open "D:\Sanat\Medicine\Database.mdb"
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
frmsupplier.WindowState = 2
generatesno
cmdSAVE.Enabled = False
CMDDEL.Enabled = False
cmdADD.Enabled = True
cmdFIND.Enabled = True
cmdEXIT.Enabled = True
End Sub

Public Sub generatesno()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select  Count(*)  from SUPPLIER ;"
rs.ActiveConnection = cn
rs.Open
Text1.Text = "s" & (rs(0) + 1)
Text1.Enabled = False
Text1.BackColor = RGB(220, 220, 220)
End Sub

Public Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Public Sub assign()
rs(0) = Trim(Text1.Text)
rs(1) = Trim(Text2.Text)
rs(2) = Trim(Text3.Text)
End Sub

Public Sub load()
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
End Sub

Public Sub DisplayData()
    Text1.Text = DataGrid1.Columns(0)
    Text2.Text = DataGrid1.Columns(1)
    Text3.Text = DataGrid1.Columns(2)
End Sub


