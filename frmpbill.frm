VERSION 5.00
Begin VB.Form frmpbill 
   BackColor       =   &H00808000&
   Caption         =   "PURCHASE BILL"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   7920
   Begin VB.TextBox txtcname 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Text            =   " "
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cmbtcode 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Text            =   " "
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtpname 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Text            =   " "
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtquantity 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Text            =   " "
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtuprice 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Text            =   " "
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Text            =   " "
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TRANSACTION NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SUPPLIER Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "MEDICINE Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Unit  Price "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   555
   End
End
Attribute VB_Name = "frmpbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
