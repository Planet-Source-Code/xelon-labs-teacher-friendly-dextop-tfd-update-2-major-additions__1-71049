VERSION 5.00
Begin VB.Form frmcalc 
   BackColor       =   &H00925C0C&
   BorderStyle     =   0  'None
   Caption         =   "Functions Calculator"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form7"
   ScaleHeight     =   4800
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   23
      Top             =   2880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Abs(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Fix(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Atn(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Sgn(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Sin(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Cos(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   9
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Tan(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   10
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Log(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   8
      Left            =   1800
      TabIndex        =   11
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Rndm"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   9
      Left            =   1800
      TabIndex        =   12
      Top             =   1680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Rnd(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   10
      Left            =   1800
      TabIndex        =   13
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Exp(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   11
      Left            =   1800
      TabIndex        =   14
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Sqr(x)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5844267
      FCOL            =   16777215
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   12
      Left            =   2880
      TabIndex        =   15
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   13
      Left            =   2880
      TabIndex        =   16
      Top             =   1680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   14
      Left            =   2880
      TabIndex        =   17
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "x"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton Func 
      Height          =   375
      Index           =   15
      Left            =   2880
      TabIndex        =   18
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "/"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   22
      Top             =   3600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   24
      Top             =   2880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "2"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   25
      Top             =   2880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "3"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   26
      Top             =   3240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   27
      Top             =   3240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "5"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   6
      Left            =   1800
      TabIndex        =   28
      Top             =   3240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "6"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   3135
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   29
      Top             =   3600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "7"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   8
      Left            =   1080
      TabIndex        =   30
      Top             =   3600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "8"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.MacButton Num 
      Height          =   375
      Index           =   9
      Left            =   1800
      TabIndex        =   31
      Top             =   3600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "9"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15325907
      FCOL            =   0
   End
   Begin Project1.title titlebar 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
   End
   Begin Project1.MacButton Cl 
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   33
      Top             =   2880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "C"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   128
      FCOL            =   16777215
   End
   Begin Project1.MacButton Del 
      Height          =   375
      Left            =   2880
      TabIndex        =   34
      Top             =   3240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "DEL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   64
      FCOL            =   15325907
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "frmcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curtex As TextBox

Private Sub Cl_Click(Index As Integer)
Text1 = ""
Text2 = ""
End Sub

Private Sub Del_Click()
Dim ss As Integer
ss = curtex.SelStart
curtex.text = Left(curtex.text, ss - 1) & Right(curtex.text, Len(curtex.text) - ss)
curtex.SelStart = ss - 1
End Sub

Private Sub Form_Load()
titlebar.sett Me
Set curtex = Text1
End Sub

Private Sub Func_Click(Index As Integer)
On Error GoTo y
Dim result As Double
If Index = 0 Then
result = Abs(Text1)
ElseIf Index = 1 Then
result = Fix(Text1)
ElseIf Index = 2 Then
result = Atn(Text1)
ElseIf Index = 3 Then
result = Sgn(Text1)
ElseIf Index = 4 Then
result = Sin(Text1)
ElseIf Index = 5 Then
result = Cos(Text1)
ElseIf Index = 6 Then
result = Tan(Text1)
ElseIf Index = 7 Then
result = Log(Text1)
ElseIf Index = 8 Then
Randomize
result = Rnd
ElseIf Index = 9 Then
result = Round(Text1)
ElseIf Index = 10 Then
result = VBA.Math.exp(Text1)
ElseIf Index = 11 Then
result = Sqr(Text1)
ElseIf Index = 12 Then
result = Val(Text1) + Val(Text2)
ElseIf Index = 13 Then
result = Val(Text1) - Val(Text2)
ElseIf Index = 14 Then
result = Val(Text1) * Val(Text2)
ElseIf Index = 15 Then
result = CStr(Val(Text1) / Val(Text2))
End If
If Index <= 11 And Index <> 8 Then
Text3 = Left(Func(Index).Caption, InStr(1, Func(Index).Caption, "(")) & Text1 & ") = " & result
ElseIf Index = 8 Then
Text3 = "Randomized Val, " & vbCrLf & "=" & result
ElseIf Index >= 12 And Index <= 15 Then
Text3 = Text1 & Func(Index).Caption & Text2 & vbCrLf & "=" & result
End If

Exit Sub
y:
MsgBox err.Description, vbCritical, "Math Error"
End Sub

Private Sub Num_Click(Index As Integer)
Dim ss As Integer
ss = curtex.SelStart
curtex.text = Left(curtex.text, ss) & CStr(Index) & Right(curtex.text, Len(curtex.text) - ss)
curtex.SelStart = ss + 1
End Sub

Private Sub Text1_GotFocus()
Set curtex = Text1
End Sub

Private Sub Text2_GotFocus()
Set curtex = Text2
End Sub
