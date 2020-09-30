VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmedicine 
   BackColor       =   &H00400000&
   Caption         =   "MEDICINE"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14505
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   14505
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Caption         =   "COMPOSITION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   9120
      TabIndex        =   24
      Top             =   1800
      Width           =   5295
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0000
         Left            =   3720
         List            =   "MEDICINE.frx":0002
         TabIndex        =   54
         Text            =   "100 mg"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox Text11 
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
         Left            =   840
         TabIndex        =   53
         Top             =   4800
         Width           =   2775
      End
      Begin VB.ComboBox Combo10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0004
         Left            =   3720
         List            =   "MEDICINE.frx":0006
         TabIndex        =   51
         Text            =   "100 mg"
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox Text10 
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
         Left            =   840
         TabIndex        =   50
         Top             =   4320
         Width           =   2775
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0008
         Left            =   3720
         List            =   "MEDICINE.frx":000A
         TabIndex        =   48
         Text            =   "100 mg"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text9 
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
         Left            =   840
         TabIndex        =   47
         Top             =   3840
         Width           =   2775
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":000C
         Left            =   3720
         List            =   "MEDICINE.frx":000E
         TabIndex        =   45
         Text            =   "100 mg"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text8 
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
         Left            =   840
         TabIndex        =   44
         Top             =   3360
         Width           =   2775
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0010
         Left            =   3720
         List            =   "MEDICINE.frx":0012
         TabIndex        =   42
         Text            =   "100 mg"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text7 
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
         Left            =   840
         TabIndex        =   41
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0014
         Left            =   3720
         List            =   "MEDICINE.frx":0016
         TabIndex        =   39
         Text            =   "100 mg"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text6 
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
         Left            =   840
         TabIndex        =   38
         Top             =   2400
         Width           =   2775
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0018
         Left            =   3720
         List            =   "MEDICINE.frx":001A
         TabIndex        =   36
         Text            =   "100 mg"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text5 
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
         Left            =   840
         TabIndex        =   35
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":001C
         Left            =   3720
         List            =   "MEDICINE.frx":001E
         TabIndex        =   33
         Text            =   "100 mg"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text4 
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
         Left            =   840
         TabIndex        =   32
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0020
         Left            =   3720
         List            =   "MEDICINE.frx":0022
         TabIndex        =   30
         Text            =   "100 mg"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text3 
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
         Left            =   840
         TabIndex        =   29
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "MEDICINE.frx":0024
         Left            =   3720
         List            =   "MEDICINE.frx":0026
         TabIndex        =   27
         Text            =   "100 mg"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text2 
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
         Left            =   840
         TabIndex        =   26
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000A&
         Caption         =   "C10"
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
         Left            =   120
         TabIndex        =   52
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000A&
         Caption         =   "C9"
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
         Left            =   240
         TabIndex        =   49
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000A&
         Caption         =   "C8"
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
         Left            =   240
         TabIndex        =   46
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000A&
         Caption         =   "C7"
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
         Left            =   240
         TabIndex        =   43
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000A&
         Caption         =   "C6"
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
         Left            =   240
         TabIndex        =   40
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000A&
         Caption         =   "C5"
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
         Left            =   240
         TabIndex        =   37
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000A&
         Caption         =   "C4"
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
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000A&
         Caption         =   "C3"
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
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000A&
         Caption         =   "C2"
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
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "C1"
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
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Height          =   1575
      Left            =   1200
      TabIndex        =   17
      Top             =   7080
      Width           =   7695
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
         TabIndex        =   23
         Top             =   240
         Width           =   1455
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
         TabIndex        =   22
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
         Left            =   3960
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CMDDELETE 
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
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
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   60817409
      CurrentDate     =   38211
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   60817409
      CurrentDate     =   38211
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "MEDICINE.frx":0028
      Left            =   2520
      List            =   "MEDICINE.frx":006E
      TabIndex        =   14
      Text            =   "10 Tablets"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4800
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox TXTUNIT 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4800
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox TXTPCATAGORY 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4800
      TabIndex        =   9
      Top             =   6120
      Width           =   3855
   End
   Begin VB.TextBox TXTPNAME 
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
      Left            =   7440
      TabIndex        =   8
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox TXTPCODE 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
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
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404080&
      Caption         =   "RETAIL PRICE"
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
      Left            =   720
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "MEDICINE DETAILS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   4065
   End
   Begin VB.Label LBLUNIT 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      Caption         =   "PACKAGE  "
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
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label LBLSELLPRICE 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      Caption         =   "EXP  DATE "
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
      Left            =   4800
      TabIndex        =   4
      Top             =   1800
      Width           =   1755
   End
   Begin VB.Label LBLUNITPRICE 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      Caption         =   "MFG DATE "
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
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label LBLPCATAGORY 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      Caption         =   "MEDICINE CATEGORY"
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
      Left            =   720
      TabIndex        =   2
      Top             =   6120
      Width           =   3330
   End
   Begin VB.Label LBLPNAME 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
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
      Left            =   4800
      TabIndex        =   1
      Top             =   1200
      Width           =   2565
   End
   Begin VB.Label LBLPCODE 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      Caption         =   "BATCH NO "
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
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   1710
   End
End
Attribute VB_Name = "frmmedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim CMD As ADODB.Command
Dim sql As String

Private Sub CMDADD_Click()
generatepartycode
clear
txtpname.SetFocus
cmdSAVE.Enabled = True
cmdADD.Enabled = False
End Sub

Private Sub cmddelete_Click()
Dim i As String
i = MsgBox("Do You Want to Delete   ", vbYesNo, "Save")
    If i = vbYes Then
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select Count(pcode) from partmaster where pcode like '" & Trim(txtpcode.Text) & "';"
    rs.ActiveConnection = cn
    rs.Open
    If (rs(0) > 1) Then
    Set CMD = New ADODB.Command
    CMD.CommandText = "update partmaster set txtpname.text="", txtpcatagory.text="" ,txtunitprice.text=,TXTSELLPRICE.Text="",txtunit.Text="" where pcode like '" & Trim(txtpcode.Text) & "';"
    'CMD.CommandText = "delete from partmaster where pcode like '" & Trim(txtpcode.Text) & "';"
    CMD.CommandType = adCmdText
     Set CMD.ActiveConnection = cn
     CMD.Execute
     Set CMD = Nothing
     Else
    Set CMD = New ADODB.Command
    CMD.CommandText = "Update partmaster set pname='" & Trim(txtpname.Text) & "',pcatagory='" & Trim(TXTPCATAGORY.Text) & "',uprice='" & Trim(TXTUNITPRICE.Text) & "',sprice='" & Trim(TXTSELLPRICE.Text) & "',unit='" & Trim(TXTUNIT.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    Set CMD = Nothing
    Set CMD = New ADODB.Command
    CMD.CommandText = "Delete from partmaster where pcode like '" & Trim(txtpcode.Text) & "';"
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
    cmdDELETE.Enabled = False
    cmdUPDATE.Enabled = False
    'End If
    End If
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
rs.Source = "Select * from partmaster where  pcode like '" & i & "'"
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
cmdSAVE.Enabled = False
cmdADD.Enabled = True
rs.Close
End Sub

Private Sub CMDSAVE_Click()
If txtpcode.Text = "" Or txtpname.Text = "" Or TXTPCATAGORY.Text = "" Or TXTUNITPRICE.Text = "" Then
  MsgBox "You Should Fill all the data in fields "
 Else
  i = MsgBox("Do You Want to Save   ", vbYesNo, "Save")
    If i = vbYes Then
       Set rs = New ADODB.Recordset
       rs.CursorType = adOpenKeyset
       rs.LockType = adLockOptimistic
       rs.Source = "partmaster"
       rs.ActiveConnection = cn
       rs.Open
       rs.AddNew
       Call assign
       rs.Update
       rs.Close
       Set rs = Nothing
       Call clear
       'Else
       'MsgBox "sss"
       End If
End If
'End If
txtpname.SetFocus
'Exit Sub
cmdADD.Enabled = True
cmdSAVE.Enabled = False
End Sub

Private Sub CMDUPDATE_Click()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
    Set CMD = New ADODB.Command
    CMD.CommandText = "update partmaster set pname='" & Trim(txtpname.Text) & "', pcatagory='" & Trim(TXTPCATAGORY.Text) & "',uprice='" & Trim(TXTUNITPRICE.Text) & "',sprice='" & Trim(TXTSELLPRICE.Text) & "',unit='" & Trim(TXTUNIT.Text) & "'where pcode like '" & Trim(txtpcode.Text) & "';"
    CMD.CommandType = adCmdText
    Set CMD.ActiveConnection = cn
    CMD.Execute
    MsgBox ("The Party is Successfully Modified")
    Set CMD = Nothing
    clear
    cmdDELETE.Enabled = False
    cmdUPDATE.Enabled = False
    cmdSAVE.Enabled = False
    cmdADD.Enabled = True
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.Open "d:\sanat\Medicine\Database.mdb"
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
'rs.Source = "Select distinct Party_Name from Master_Ledger"
'rs.ActiveConnection = cn
'rs.Open
'While Not rs.EOF
'cmbIName.AddParty rs(0)
'rs.MoveNext
'Wend
'rs.Close
frmmedicine.WindowState = 2
cmdNEXT.Enabled = False
cmdSAVE.Enabled = False
cmdDELETE.Enabled = False
cmdADD.Enabled = True
End Sub

Public Sub generatepartycode()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
'rs.Source = "Select Society_Code from Society_Ledger"
rs.Source = "select  Count(*)  from MEDICINE ;"
'Transaction_Ledger where Messer_No having (select distinct(Messer_No) from Transaction_Ledger)  ;"
rs.ActiveConnection = cn
rs.Open
'While Not rs.EOF
'MsgBox (rs(0, 1))
'rs (0) + 1
'Wend
txtpcode.Text = "P" & (rs(0) + 1)
txtpcode.Enabled = False
txtpcode.BackColor = RGB(220, 220, 220)
End Sub

Public Sub clear()
txtpname = ""
TXTPCATAGORY = ""
TXTUNITPRICE = ""
TXTSELLPRICE = ""
TXTUNIT = ""
End Sub

Public Sub assign()
rs(0) = txtpcode.Text
rs(1) = txtpname.Text
rs(2) = TXTPCATAGORY.Text
rs(3) = TXTUNITPRICE.Text
rs(4) = TXTSELLPRICE.Text
rs(5) = TXTUNIT.Text

End Sub

Public Sub load()
txtpcode.Text = rs(0)
txtpname.Text = rs(1)
TXTPCATAGORY.Text = rs(2)
TXTUNITPRICE.Text = rs(3)
TXTSELLPRICE.Text = rs(4)
TXTUNIT.Text = rs(5)
End Sub
