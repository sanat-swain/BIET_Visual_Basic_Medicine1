VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsale 
   BackColor       =   &H00008000&
   Caption         =   "SALE"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   10545
   Begin VB.TextBox Text10 
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
      Left            =   6720
      TabIndex        =   31
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text9 
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
      Left            =   3480
      TabIndex        =   29
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text8 
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
      Left            =   6720
      TabIndex        =   27
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text7 
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
      Left            =   3480
      TabIndex        =   26
      Top             =   4680
      Width           =   735
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
      Left            =   6600
      TabIndex        =   25
      Top             =   3960
      Width           =   1095
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
      Left            =   3480
      TabIndex        =   24
      Top             =   3960
      Width           =   1575
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
      Left            =   3480
      TabIndex        =   23
      Top             =   3240
      Width           =   4575
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
      Left            =   3480
      TabIndex        =   22
      Top             =   2520
      Width           =   1575
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
      Left            =   3480
      TabIndex        =   21
      Top             =   1800
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
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
      Format          =   60882945
      CurrentDate     =   38211
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
      Left            =   3480
      TabIndex        =   19
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000040&
      Height          =   1575
      Left            =   1920
      TabIndex        =   11
      Top             =   6360
      Width           =   7455
      Begin VB.CommandButton cmdUPDATE 
         Caption         =   "UPDATE"
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
         Left            =   3840
         TabIndex        =   33
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdNEW 
         Caption         =   "NEW"
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
         Left            =   5640
         TabIndex        =   32
         Top             =   240
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
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1455
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
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTOTAL 
         Caption         =   "TOTAL"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1575
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
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
         Left            =   5640
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
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
      Left            =   4320
      TabIndex        =   30
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
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
      Left            =   7800
      TabIndex        =   28
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DISCOUNT"
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
      Left            =   600
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblstock 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TAX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5640
      TabIndex        =   10
      Top             =   3960
      Width           =   630
   End
   Begin VB.Label lblsprice 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "SELL PRICE Rs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   4320
      TabIndex        =   9
      Top             =   4680
      Width           =   2340
   End
   Begin VB.Label lbltotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TOTAL  Rs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5040
      TabIndex        =   8
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "SALE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbltcode 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TRANSACTION NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "DATE "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5280
      TabIndex        =   5
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label lblcname 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "CUSTOMER NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   2760
   End
   Begin VB.Label lblpcode 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "MEDICINE NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   2475
   End
   Begin VB.Label lblpname 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "BATCH NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1620
   End
   Begin VB.Label lblquantity 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   1
      Top             =   4680
      Width           =   1545
   End
   Begin VB.Label lblunit 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "RETAIL PRICE   Rs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   0
      Top             =   3960
      Width           =   2835
   End
End
Attribute VB_Name = "frmsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim CMD As ADODB.Command
Dim sql As String
Private Sub CMBPNAME_click()
 Set rs = New ADODB.Recordset
 rs.CursorType = adOpenKeyset
 rs.LockType = adLockOptimistic
 rs.Source = "select pcode,unit,sprice,stockist.stock from partmaster,stockist where pname like '" & Trim(cmbpname.Text) & "';"
 rs.ActiveConnection = cn
 rs.Open
 txtpcode.Text = rs(0)
 TXTUNIT.Text = rs(1)
 txtsprice = rs(2)
 rs.Close
 txtpcode.Enabled = False
 txtstock.Enabled = False
End Sub

Private Sub CMDADD_Click()
clear
cmdSAVE.Enabled = True
txtcname.SetFocus
generatetcode
txttcode.Enabled = False
cmbpname.Refresh
End Sub

Private Sub cmddelete_Click()
Dim i As String
i = MsgBox("Do You Want to Delete   ", vbYesNo, "Save")
    If i = vbYes Then
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select Count(tcode) from transaction2 where tcode like '" & Trim(txttcode.Text) & "';"
    rs.ActiveConnection = cn
    rs.Open
    If (rs(0) > 1) Then
    Set CMD = New ADODB.Command
    CMD.CommandText = "delete from transaction2 where tcode like '" & Trim(txttcode.Text) & "';"
    CMD.CommandType = adCmdText
     Set CMD.ActiveConnection = cn
     CMD.Execute
     Set CMD = Nothing
     Else
    Set CMD = New ADODB.Command
    'CMD.CommandText = "Update partmaster set pname='" & Trim(TXTPNAME.Text) & "',pcatagory='" & Trim(TXTPCATAGORY.Text) & "',uprice='" & Trim(TXTUNITPRICE.Text) & "',sprice='" & Trim(TXTSELLPRICE.Text) & "',unit='" & Trim(txtunit.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    Set CMD = New ADODB.Command
    CMD.CommandText = "Delete from transaction where tcode like '" & Trim(txttcode.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    MsgBox ("The Record is successfully deleted")
    'Else
    
    'End If
    End If
    clear
    'cmbSoName.SetFocus
    cmdDELETE.Enabled = True
    cmdUPDATE.Enabled = False
    'End If
    End If
End Sub

Private Sub CMDEXIT_Click()
End
End Sub

Private Sub CMDFIND_Click()
Dim i As String
i = InputBox("Enter The Party Code U want to find:")
clear
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "Select * from transaction2 where  tcode like '" & i & "';"
rs.ActiveConnection = cn
rs.Open
If rs.EOF Then
MsgBox ("TRANSACTION WITH THIS CODE DOSENOT EXIST")
cmdDELETE.Enabled = False
cmdModify.Enabled = False
Else
clear
load
End If
rs.Close
End Sub

Private Sub CMDSAVE_Click()
Dim i As String
' If TXTSLNO.Text = "" Or cmbpname.Text = "" Then
  'MsgBox "You Should Fill SLNO And PNAME "
 'Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "TRANSACTION2"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       assign
       rs.Update
       rs.Close
       MsgBox "SUCCESSFULLY SAVED"
    End If
    Set CMD = New ADODB.Command
    CMD.CommandText = "Update STOCKIST set STOCK=" & Val(txtstock.Text) - Val(txtquantity.Text) & " where pname like '" & Trim(cmbpname.Text) & "'; "
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    clear
    Set CMD = New ADODB.Command
    CMD.CommandText = "delete from stockist where stock=0; "
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
  
cmbpname.SetFocus
cmdADD.Enabled = True
cmdSAVE.Enabled = False
End Sub

Private Sub cmdstock_Click()
frmstockshow.Show
End Sub

Private Sub Form_Load()
frmsale.WindowState = 2
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open ("D:\Sanat\Medicine\Database.mdb")
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "SELECT MEDICINENAME From MEDICINE"
rs.ActiveConnection = cn
rs.Open
While Not rs.EOF
cmbpname.AddItem rs(0)
rs.MoveNext
Wend
rs.Close
cmdSAVE.Enabled = False
cmdDELETE.Enabled = False
cmdADD.Enabled = True
cmdFIND.Enabled = True
cmdEXIT.Enabled = True
End Sub
Public Sub ASSIGN2()
rs(0) = txttcode.Text
rs(1) = cmbdate.Value
rs(2) = txtcname.Text
End Sub


Public Sub generatetcode()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Source = "select  count(*)  from transaction2 ;"
rs.ActiveConnection = cn
rs.Open
txttcode.Text = (rs(0) + 1)
End Sub

Public Sub clear()
txtcname.Text = ""
cmbpname.Text = ""
txtpcode.Text = ""
txtquantity.Text = ""
TXTUNIT.Text = ""
txttotal.Text = ""
txtstock.Text = ""
txtsprice.Text = ""


End Sub

Private Sub txttotal_GotFocus()
txttotal.Text = Val(txtquantity.Text) * Val(txtsprice.Text)
End Sub

Public Sub assign()
rs(0) = Trim(txttcode.Text)
rs(1) = cmbdate.Value
rs(2) = Trim(txtcname.Text)
rs(3) = cmbpname.Text
rs(4) = Trim(txtpcode.Text)
rs(5) = Trim(txtquantity.Text)
rs(6) = Trim(txtsprice.Text)
rs(7) = Trim(TXTUNIT.Text)
rs(8) = Trim(txttotal.Text)
End Sub

Public Sub load()
txttcode.Text = rs(0)
cmbdate.Value = rs(1)
txtcname.Text = rs(2)
cmbpname.Text = rs(3)
txtpcode.Text = rs(4)
txtquantity.Text = rs(5)
txtsprice.Text = rs(6)
TXTUNIT.Text = rs(7)
txttotal.Text = rs(8)
'txtstock.Text = ""
'txtsprice.Text = ""

End Sub
