VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstock 
   BackColor       =   &H00008000&
   Caption         =   "STOCK"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11820
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      Height          =   1575
      Left            =   2160
      TabIndex        =   19
      Top             =   4560
      Width           =   5775
      Begin VB.CommandButton cmdEXIT 
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
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdFIND 
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
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdDELETE 
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSAVE 
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
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdADD 
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
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7800
      TabIndex        =   10
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtpcode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox TXTSLNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   60817409
      CurrentDate     =   36527
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TOTAL PRICE Rs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   17
      Top             =   3360
      Width           =   2580
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "UNIT PRICE     Rs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   15
      Top             =   3360
      Width           =   2625
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "QUANTIT            "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   13
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "PURCHASE DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   11
      Top             =   2760
      Width           =   2670
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SUPPLIER NO        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "STOCK ENTRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   3600
   End
   Begin VB.Label LBLSNAME 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Height          =   360
      Left            =   5160
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label LBLPCODE 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "BATCH NO        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   2265
   End
   Begin VB.Label LBLPNAME 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "MEDICINE NAME "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   3
      Top             =   1560
      Width           =   2565
   End
   Begin VB.Label LBLSLNO 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "SL NO               "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "frmstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim CMD As ADODB.Command
Private Sub CMBPNAME_click()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select PCODE,uprice,unit from partmaster where PName like '" & Trim(cmbpname.Text) & "'; "
rs.ActiveConnection = cn
rs.Open
txtpcode.Text = rs(0)
txtpurprice.Text = rs(1)
TXTUNIT.Text = rs(2)
End Sub
Private Sub cmbsname_click()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select address from dealer  where sname like '" & Trim(cmbsname.Text) & "'; "
rs.ActiveConnection = cn
rs.Open
txtsadd.Text = rs(0)
End Sub
Private Sub CMDADD_Click()
clear
generatesno
cmbpname.SetFocus
cmdSAVE.Enabled = True
cmdADD.Enabled = False
End Sub

Private Sub cmddelete_Click()
Dim i As String
i = "Are you sure to delete?"
If MsgBox(i, vbYesNo, "warning?") = vbYes Then
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "delete sname,address from stock  where slno like'" & Trim(TXTSLNO.Text) & "';"
rs.ActiveConnection = cn
rs.Open
'Set CMD = New ADODB.Command
'CMD.CommandText = "update dealer set sname=" ",address=""  where sno like'" & Trim(Text1.Text) & "';"
'CMD.CommandText = "update dealer set sname="",address=""  where sno like'" & Trim(Text1.Text) & "';"
'CMD.CommandType = adCmdText
'Set CMD.ActiveConnection = CN
'CMD.Execute
MsgBox "Successfully deleted", vbExclamation, "Deletion"
clear2
ElseIf MsgBox(i, vbYesNo, "Warning?") = vbNo Then
MsgBox "Do you want to exit", vbOKOnly, "Stop"
End If
'End If
Set rs = Nothing
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub CMDFIND_Click()
Dim i As String
i = InputBox("Enter The slno U want to find:")
clear
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select * from stock where  slno like '" & i & "'"
rs.ActiveConnection = cn
rs.Open
If rs.EOF Then
MsgBox ("The PARTY With This Code is Not Exist")
cmdDELETE.Enabled = False
Else
clear
load
End If
cmdDELETE.Enabled = True
cmdUPDATE.Enabled = True
cmdADD.Enabled = False
rs.Close
End Sub

Private Sub CMDSAVE_Click()
Dim i As String
If TXTSLNO.Text = "" Or cmbpname.Text = "" Then
  MsgBox "You Should Fill SLNO And PNAME "
 Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "select pcode from stockist where pcode like '" & Trim(txtpcode.Text) & "';"
       rs.ActiveConnection = cn
       rs.Open
       If rs.EOF = True Then
       rs.Close
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "stockist"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       ASSIGN2
       rs.Update
       rs.Close
       Else
      Set CMD = New ADODB.Command
      x = Val(txtquantity.Text)
      MsgBox (x)
      CMD.CommandText = "update stockist set stock=stock+ " & Val(txtquantity.Text) & " where pcode like '" & Trim(txtpcode.Text) & "';"
      CMD.CommandType = adCmdText
      Set CMD.ActiveConnection = cn
      CMD.Execute
      MsgBox ("The Party is Successfully Modified")
      Set CMD = Nothing
      End If
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "stock"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       Call assign
       rs.Update
       rs.Close
       Set rs = Nothing
         
End If
End If
cmbpname.SetFocus
cmdADD.Enabled = True
cmdSAVE.Enabled = False
End Sub

Private Sub CMDUPDATE_Click()
Set CMD = New ADODB.Command
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
    CMD.CommandText = "update stock set purprice='" & Trim(txtpurprice.Text) & "'where slno like '" & Trim(TXTSLNO.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    MsgBox ("The Party is Successfully Modified")
    Set CMD = Nothing
    clear2
    cmdDELETE.Enabled = False
    cmdUPDATE.Enabled = False
    cmdSAVE.Enabled = False
    cmdADD.Enabled = True
End Sub

Private Sub CMDshowstock_Click()
frmstockshow.Show
End Sub


Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open ("D:\Sanat\Medicine\Database.mdb")
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "SELECT MEDICINENAME FROM MEDICINE"
rs.ActiveConnection = cn
rs.Open
While Not rs.EOF
cmbpname.AddItem rs(0)
rs.MoveNext
Wend
rs.Close
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "SELECT SUPPLIERNAME FROM SUPPLIER"
rs.ActiveConnection = cn
rs.Open
While Not rs.EOF
cmbsname.AddItem rs(0)
rs.MoveNext
Wend
frmstock.WindowState = 2
cmdSAVE.Enabled = False
cmdDELETE.Enabled = False
cmdADD.Enabled = True
cmdFIND.Enabled = True
cmdEXIT.Enabled = True
End Sub
Public Sub assign()
rs(0) = TXTSLNO.Text
rs(2) = cmbpname.Text
rs(1) = txtpcode.Text
rs(3) = cmbsname.Text
rs(4) = txtsadd.Text
rs(5) = cmbdate.Value
rs(6) = txtquantity.Text
rs(7) = TXTUNIT.Text
rs(8) = txtpurprice.Text
'RS(9) = txtstock.Text
End Sub
Public Sub clear()
TXTSLNO.Text = ""
cmbpname.Text = ""
txtpcode.Text = ""
cmbsname.Text = ""
txtsadd.Text = ""
'cmbdate.Value = ""
txtquantity.Text = ""
TXTUNIT.Text = ""
txtpurprice.Text = ""
txtstock.Text = ""
End Sub
Public Sub generatesno()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select  Count(*)  from stock ;"
rs.ActiveConnection = cn
rs.Open
TXTSLNO.Text = (rs(0) + 1)
End Sub

Public Sub load()
TXTSLNO.Text = rs(0)
cmbpname.Text = rs(2)
txtpcode.Text = rs(1)
cmbsname.Text = rs(3)
txtsadd.Text = rs(4)
cmbdate.Value = rs(5)
txtquantity.Text = rs(6)
TXTUNIT.Text = rs(7)
txtpurprice.Text = rs(8)
'txtstock.Text = RS(9)
End Sub

Public Sub clear2()
cmbpname.Text = ""
txtpcode.Text = ""
cmbsname.Text = ""
txtsadd.Text = ""
'cmbdate.Value = ""
txtquantity.Text = ""
TXTUNIT.Text = ""
txtpurprice.Text = ""
txtstock.Text = ""
End Sub
Private Sub txtstock_GotFocus()
'Set RS = New ADODB.Recordset
'RS.CursorType = adOpenKeyset
'RS.LockType = adLockOptimistic
'RS.Source = "select stock from stock"
'RS.ActiveConnection = CN
'RS.Open
End Sub

Public Sub ASSIGN2()
rs(0) = cmbpname.Text
rs(1) = txtpcode.Text
rs(2) = txtquantity.Text
End Sub
