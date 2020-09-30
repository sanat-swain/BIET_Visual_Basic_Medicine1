VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00400000&
   Caption         =   "SONOSCAN MEDICINES"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10140
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnumaster 
      Caption         =   "&MASTER"
      Begin VB.Menu mnupmaster 
         Caption         =   "&MEDICINES"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnusupplier 
         Caption         =   "&SUPPLIER"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuentry 
      Caption         =   "&ENTRY"
      Begin VB.Menu MNUSTOCK 
         Caption         =   "&STOCK"
         Shortcut        =   ^S
      End
      Begin VB.Menu MNUSALE 
         Caption         =   "S&ALE"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MNUREPORT 
      Caption         =   "&REPORT"
      Begin VB.Menu RPTSTOCK 
         Caption         =   "ST&OCK"
         Shortcut        =   ^O
      End
      Begin VB.Menu RPTTRANSACT 
         Caption         =   "SA&LES"
         Shortcut        =   ^L
      End
      Begin VB.Menu RPTSUPPLIER 
         Caption         =   "SU&PPLIER"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu MNUBILLS 
      Caption         =   "&BILLS"
      Begin VB.Menu mnuSALES 
         Caption         =   "SAL&ES"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuPURCHASE 
         Caption         =   "PUR&CHASE"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
MDIForm1.WindowState = 2
End Sub

Private Sub MNUBILLS_Click()
'frmTBill.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnupmaster_Click()
frmmedicine.Show
End Sub

Private Sub mnustock_Click()
frmstock.Show
End Sub

Private Sub mnusupplier_Click()
frmsupplier.Show
End Sub

Private Sub mnuSALE_Click()
frmsale.Show
End Sub

Private Sub RPTSTOCK_Click()
stockreport.Show
End Sub

Private Sub RPTSUPPLIER_Click()
supplierReport.Show
End Sub

Private Sub RPTTRANSACT_Click()
TRANSACTIONREPORT.Show
End Sub
