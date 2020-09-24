VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmcst 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Dextop Settings"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14220
   LinkTopic       =   "Form5"
   ScaleHeight     =   4815
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MacButton MacButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "About"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   15592683
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton3 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   "Desktop"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton4 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   "Colors"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton5 
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   "Clock Properties"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton2 
      Height          =   255
      Left            =   3600
      TabIndex        =   32
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   "Skin"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton24 
      Height          =   375
      Left            =   2040
      TabIndex        =   30
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Apply"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton19 
      Height          =   375
      Left            =   3360
      TabIndex        =   29
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Cancel"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   9480
      ScaleHeight     =   3825
      ScaleWidth      =   4545
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
      Begin VB.Frame Frame2 
         BackColor       =   &H0059341C&
         Caption         =   "Change Graphics Profile"
         ForeColor       =   &H00F2E2D9&
         Height          =   3495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4335
         Begin Project1.MacButton MacButton22 
            Height          =   375
            Left            =   1560
            TabIndex        =   26
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "South"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin Project1.MacButton MacButton17 
            Height          =   375
            Left            =   3000
            TabIndex        =   22
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "N-E"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin Project1.MacButton MacButton16 
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "North"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin Project1.MacButton MacButton15 
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "N-W"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2C9C0&
            BorderStyle     =   0  'None
            Caption         =   "Clock Images"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   19
            Top             =   2040
            Width           =   4095
            Begin VB.Line Line3 
               X1              =   3000
               X2              =   3000
               Y1              =   240
               Y2              =   1200
            End
            Begin VB.Line Line2 
               X1              =   2040
               X2              =   2040
               Y1              =   240
               Y2              =   1200
            End
            Begin VB.Line Line1 
               X1              =   1080
               X2              =   1080
               Y1              =   240
               Y2              =   1200
            End
            Begin Project1.aicAlphaImage aicAlphaImage4 
               Height          =   975
               Left            =   120
               ToolTipText     =   "Back Image"
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1720
               Image           =   "frmcln.frx":0000
               Scaler          =   1
            End
            Begin Project1.aicAlphaImage aicAlphaImage3 
               Height          =   975
               Left            =   1080
               ToolTipText     =   "Hour Image"
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1720
               Image           =   "frmcln.frx":0018
               Scaler          =   1
            End
            Begin Project1.aicAlphaImage aicAlphaImage2 
               Height          =   975
               Left            =   2040
               ToolTipText     =   "Minute Image"
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1720
               Image           =   "frmcln.frx":0030
               Scaler          =   1
            End
            Begin Project1.aicAlphaImage aicAlphaImage1 
               Height          =   975
               Left            =   3000
               ToolTipText     =   "Second Image"
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1720
               Image           =   "frmcln.frx":0048
               Scaler          =   1
            End
         End
         Begin Project1.MacButton MacButton18 
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "West"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin Project1.MacButton MacButton20 
            Height          =   375
            Left            =   3000
            TabIndex        =   24
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "East"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin Project1.MacButton MacButton21 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "S-W"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin Project1.MacButton MacButton23 
            Height          =   375
            Left            =   3000
            TabIndex        =   27
            Top             =   1320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "S-E"
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            FCOL            =   0
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Set Position of Clock"
            ForeColor       =   &H00F2E2D9&
            Height          =   375
            Left            =   1560
            TabIndex        =   28
            Top             =   840
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   4800
      ScaleHeight     =   3825
      ScaleWidth      =   4545
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
      Begin Project1.MacButton MacButton14 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Select Color"
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
      Begin Project1.MacButton MacButton13 
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Menu ForeColor"
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
      Begin Project1.MacButton MacButton12 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Menu BackColor"
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
      Begin Project1.MacButton MacButton11 
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Label Color"
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
      Begin Project1.MacButton MacButton10 
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Border"
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
      Begin Project1.MacButton MacButton9 
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Label ForeColor"
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
      Begin Project1.MacButton MacButton8 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Opaque Label"
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
      Begin Project1.MacButton MacButton7 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Set Label BackColor"
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
      Begin VB.Frame Frame1 
         BackColor       =   &H0059341C&
         Caption         =   "Set Color Scheme"
         ForeColor       =   &H00F2E2D9&
         Height          =   3495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4335
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Following is a collection of Color Schemes Apply them to your current Registry Profile"
            ForeColor       =   &H00F2E2D9&
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   3855
         End
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4800
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3825
      ScaleWidth      =   4545
      TabIndex        =   3
      Top             =   600
      Width           =   4575
      Begin VB.PictureBox image1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   480
         ScaleHeight     =   2985
         ScaleWidth      =   3705
         TabIndex        =   34
         Top             =   240
         Width           =   3735
      End
      Begin Project1.MacButton MacButton6 
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         ToolTipText     =   "select Dextop color"
         Top             =   3360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BTYPE           =   4
         TX              =   ".../\"
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
      Begin Project1.LabelText LabelText1 
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   3360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         Caption         =   "Select"
         Text            =   "Image Path"
      End
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   529
   End
End
Attribute VB_Name = "frmcst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Dim flw As String
Dim Bdr As String
Dim Opq As String
Dim Posit As String

Private Sub aicAlphaImage1_Click()
On Error Resume Next
On Error Resume Next
cd.ShowOpen
aicAlphaImage1.LoadImage_FromFile cd.filename
aicAlphaImage1.Tag = cd.filename
End Sub

Private Sub aicAlphaImage2_Click()
On Error Resume Next
cd.ShowOpen
aicAlphaImage2.LoadImage_FromFile cd.filename
aicAlphaImage2.Tag = cd.filename
End Sub

Private Sub aicAlphaImage3_Click()
On Error Resume Next
cd.ShowOpen
aicAlphaImage3.LoadImage_FromFile cd.filename
aicAlphaImage3.Tag = cd.filename
End Sub

Private Sub aicAlphaImage4_Click()
On Error Resume Next
cd.ShowOpen
aicAlphaImage4.LoadImage_FromFile cd.filename
aicAlphaImage4.Tag = cd.filename
End Sub

Private Sub Drive1_Change()
Dir1 = Drive1
End Sub

Private Sub Form_Load()
On Error Resume Next
titlebar.sett Me
LabelText1.Set_Browse
flw = App.path & "\config.ini"
Me.Width = 4575
aicAlphaImage4.Tag = Form1.aicAlphaImage1.Tag
aicAlphaImage3.Tag = Form1.aicHour.Tag
aicAlphaImage4.Tag = Form1.aicMinute.Tag
aicAlphaImage4.Tag = Form1.aicSecond.Tag
GetIni
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
Title.blink
End Sub

Private Sub Form_LostFocus()
On Error Resume Next
Title.unblink
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
aicAlphaImage1.ClearImage
Set aimage = Nothing
aicAlphaImage2.ClearImage
Set aimage = Nothing
aicAlphaImage3.ClearImage
Set aimage = Nothing
aicAlphaImage4.ClearImage
Set aimage = Nothing

End Sub

Private Sub LabelText1_Browsed()
On Error Resume Next
cd.ShowOpen
Set Form1.Picture = LoadPicture(cd.filename)
Form1.LoadBkg cd.filename
LoadBkg cd.filename
LabelText1.text = cd.filename
LabelText1.Apply
End Sub

Sub LoadBkg(pth As String)
Dim c32 As New c32bppDIB
Image1.AutoRedraw = True
c32.InitializeDIB Image1.Width / 15, Image1.Height / 15
c32.LoadPicture_File pth
Set Image1.Picture = Nothing
c32.Render Image1.hdc, 0, 0, Image1.Width / 15, Image1.Height / 15
Image1.Picture = Image1.Image
Image1.Refresh
Image1.AutoRedraw = False
Set c32 = Nothing
End Sub

Private Sub MacButton1_Click()
frmabout.Show
End Sub

Private Sub MacButton10_Click()
On Error Resume Next
If Bdr = "True" Then
Bdr = "False"
MacButton10.BackColor = &HC0C0C0
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).BorderStyle = 0
Next
Else
Bdr = "True"
MacButton10.BackColor = &H808080
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).BorderStyle = 1
Next
End If
End Sub

Private Sub MacButton11_Click()
On Error Resume Next
cd.ShowColor
MacButton11.BackColor = cd.color
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).ForeColor = cd.color
Next
End Sub

Private Sub MacButton12_Click()
On Error Resume Next
cd.ShowColor
MacButton12.BackColor = cd.color
Form1.MENU1.Set_BackColor MacButton12.BackColor
Form1.MENU2.Set_BackColor MacButton12.BackColor
End Sub

Private Sub MacButton13_Click()
On Error Resume Next
cd.ShowColor
MacButton13.BackColor = cd.color
Form1.MENU1.Set_ForeColor MacButton13.BackColor
Form1.MENU2.Set_ForeColor MacButton13.BackColor
End Sub

Private Sub MacButton14_Click()
On Error Resume Next
cd.ShowColor
MacButton14.BackColor = cd.color
Form1.Line1.BorderColor = cd.color
Form1.Line2.BorderColor = cd.color
Form1.Line3.BorderColor = cd.color
Form1.Line4.BorderColor = cd.color
'Form1.MENU1.Set_Roller_BackColor MacButton14.BackColor
'Form1.MENU2.Set_Roller_BackColor MacButton14.BackColor
End Sub

Private Sub MacButton15_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = 0
Form1.aicAlphaImage1.Top = 0
Form1.aicHour.Left = 0
Form1.aicMinute.Left = 0
Form1.aicSecond.Left = 0
Form1.aicHour.Top = 0
Form1.aicMinute.Top = 0
Form1.aicSecond.Top = 0
Posit = "N-W"
End Sub

Private Sub MacButton16_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = Form1.Width / 15 / 2
Form1.aicAlphaImage1.Top = 0
Form1.aicHour.Left = Form1.Width / 15 / 2
Form1.aicMinute.Left = Form1.Width / 15 / 2
Form1.aicSecond.Left = Form1.Width / 15 / 2
Form1.aicHour.Top = 0
Form1.aicMinute.Top = 0
Form1.aicSecond.Top = 0
Posit = "North"
End Sub

Private Sub MacButton17_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = Form1.Width / 15 - Form1.aicAlphaImage1.Width
Form1.aicAlphaImage1.Top = 0
Form1.aicHour.Left = Form1.Width / 15 - Form1.aicHour.Width
Form1.aicMinute.Left = Form1.Width / 15 - Form1.aicMinute.Width
Form1.aicSecond.Left = Form1.Width / 15 - Form1.aicSecond.Width
Form1.aicHour.Top = 0
Form1.aicMinute.Top = 0
Form1.aicSecond.Top = 0
Posit = "N-E"
End Sub

Private Sub MacButton18_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = 0
Form1.aicHour.Left = 0
Form1.aicMinute.Left = 0
Form1.aicSecond.Left = 0
Form1.aicAlphaImage1.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Form1.aicHour.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Form1.aicMinute.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Form1.aicSecond.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Posit = "West"
End Sub

Private Sub MacButton19_Click()
On Error Resume Next
If MsgBox("Do you want to Apply & Save Changes", vbYesNo, "Confirmation") = vbYes Then
MacButton24_Click
End If
Unload Me
End Sub

Private Sub MacButton2_Click()
Sknr.Show
End Sub

Private Sub MacButton20_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = Form1.Width / 15 - Form1.aicAlphaImage1.Width
Form1.aicHour.Left = Form1.Width / 15 - Form1.aicAlphaImage1.Width
Form1.aicMinute.Left = Form1.Width / 15 - Form1.aicAlphaImage1.Width
Form1.aicSecond.Left = Form1.Width / 15 - Form1.aicAlphaImage1.Width
Form1.aicAlphaImage1.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Form1.aicHour.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Form1.aicMinute.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Form1.aicSecond.Top = Form1.Height / 15 / 2 - 1 / Form1.aicAlphaImage1.Height
Posit = "East"
End Sub

Private Sub MacButton21_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = 0
Form1.aicAlphaImage1.Top = Form1.Height / 15 - Form1.aicAlphaImage1.Height
Form1.aicHour.Left = 0
Form1.aicMinute.Left = 0
Form1.aicSecond.Left = 0
Form1.aicHour.Top = Form1.Height / 15 - Form1.aicHour.Height
Form1.aicMinute.Top = Form1.Height / 15 - Form1.aicMinute.Height
Form1.aicSecond.Top = Form1.Height / 15 - Form1.aicSecond.Height
Posit = "S-W"
End Sub

Private Sub MacButton22_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = Form1.Width / 15 / 2
Form1.aicAlphaImage1.Top = Form1.Height / 15 - Form1.aicAlphaImage1.Height
Form1.aicHour.Left = Form1.Width / 15 / 2
Form1.aicMinute.Left = Form1.Width / 15 / 2
Form1.aicSecond.Left = Form1.Width / 15 / 2
Form1.aicHour.Top = Form1.Height / 15 - Form1.aicHour.Height
Form1.aicMinute.Top = Form1.Height / 15 - Form1.aicMinute.Height
Form1.aicSecond.Top = Form1.Height / 15 - Form1.aicSecond.Height
Posit = "South"
End Sub

Private Sub MacButton23_Click()
On Error Resume Next
Form1.aicAlphaImage1.Left = Form1.Width / 15 - Form1.aicAlphaImage1.Width
Form1.aicAlphaImage1.Top = Form1.Height / 15 - Form1.aicAlphaImage1.Height
Form1.aicHour.Left = Form1.Width / 15 - Form1.aicHour.Width
Form1.aicMinute.Left = Form1.Width / 15 - Form1.aicMinute.Width
Form1.aicSecond.Left = Form1.Width / 15 - Form1.aicSecond.Width
Form1.aicHour.Top = Form1.Height / 15 - Form1.aicHour.Height
Form1.aicMinute.Top = Form1.Height / 15 - Form1.aicMinute.Height
Form1.aicSecond.Top = Form1.Height / 15 - Form1.aicSecond.Height
Posit = "S-E"
End Sub

Private Sub MacButton24_Click()
On Error Resume Next
Dim flw As String
flw = App.path & "\Config.ini"
Call WriteIni("BackGround", "Wallpaper", LabelText1.text, flw)
Call WriteIni("BackGround", "BackColor", Picture1.BackColor, flw)
Call WriteIni("Color", "Opaque label", Opq, flw)
Call WriteIni("Color", "Label BackColor", MacButton7.BackColor, flw)
Call WriteIni("Color", "Label ForeColor", MacButton9.BackColor, flw)
Call WriteIni("Color", "Border", Bdr, flw)
Call WriteIni("Color", "Border Color", MacButton11.BackColor, flw)
Call WriteIni("Color", "Menu BackColor", MacButton12.BackColor, flw)
Call WriteIni("Color", "Menu ForeColor", MacButton13.BackColor, flw)
Call WriteIni("Color", "Menu Border Color", MacButton14.BackColor, flw)
Call WriteIni("Clock", "Back", aicAlphaImage4.Tag, flw)
Call WriteIni("Clock", "Hour", aicAlphaImage3.Tag, flw)
Call WriteIni("Clock", "Minute", aicAlphaImage2.Tag, flw)
Call WriteIni("Clock", "Second", aicAlphaImage1.Tag, flw)
Call WriteIni("Clock", "Position", Posit, flw)
End Sub


Private Sub MacButton3_Click()
On Error Resume Next
On Error Resume Next
Picture1.Left = 0
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
MacButton3.BackColor = &H404040
MacButton4.BackColor = &HC0C0C0
MacButton5.BackColor = &HC0C0C0
End Sub

Private Sub MacButton4_Click()
On Error Resume Next
Picture2.Left = 0
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
MacButton4.BackColor = &H404040
MacButton3.BackColor = &HC0C0C0
MacButton5.BackColor = &HC0C0C0
End Sub

Private Sub MacButton5_Click()
On Error Resume Next
Picture3.Left = 0
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = True
MacButton5.BackColor = &H404040
MacButton4.BackColor = &HC0C0C0
MacButton3.BackColor = &HC0C0C0
End Sub

Private Sub MacButton6_Click()
On Error Resume Next
cd.ShowColor
Picture1.BackColor = cd.color
Form1.BackColor = cd.color
Form1.LoadBkg LabelText1.text
End Sub

Private Sub MacButton7_Click()
On Error Resume Next
cd.ShowColor
MacButton7.BackColor = cd.color
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).BackColor = cd.color
Next
End Sub

Private Sub MacButton8_Click()
On Error Resume Next
If Opq = "True" Then
Opq = "False"
MacButton8.BackColor = &HC0C0C0
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).BackStyle = 0
Next
Else
Opq = "True"
MacButton8.BackColor = &H808080
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).BackStyle = 1
Next
End If
End Sub

Private Sub MacButton9_Click()
On Error Resume Next
cd.ShowColor
MacButton9.BackColor = cd.color
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).ForeColor = cd.color
Next
For x = 0 To tasks.UBound
Form1.tasks(i).ForeColor = cd.color
Next
For x = 0 To Bag.UBound
Form1.Bag(i).ForeColor = cd.color
Next
For x = 0 To Cafe.UBound
Form1.Cafe(i).ForeColor = cd.color
Next
End Sub
Public Sub GetIni()
On Error Resume Next
Dim img As String
Dim str As String
flw = App.path & "\config.ini"
Set Form1.Picture = Nothing
img = GetFromIni("BackGround", "Wallpaper", flw)
Form1.BackColor = GetFromIni("BackGround", "BackColor", flw)
                    If Right$(img, 10) = " <AppPath>" Then
                    img = App.path & "\" & Left(img, Len(img) - 10)
                    End If
LabelText1.text = img
Set Form1.Picture = LoadPicture(img)
LoadBkg img
frmcst.Image1.Picture = LoadPicture(GetFromIni("BackGround", "Wallpaper", flw))
frmcst.Picture1.BackColor = GetFromIni("BackGround", "BackColor", flw)
Opq = GetFromIni("Color", "Opaque label", flw)
If Opq = "True" Then
MacButton8.BackColor = &H808080
For i = 0 To Form1.imgicon.UBound
Form1.lblcaption(i).BackStyle = 1
Next
Else
MacButton8.BackColor = &HC0C0C0
For i = 0 To Form1.imgicon.UBound
Form1.lblcaption(i).BackStyle = 0
Next
End If
MacButton7.BackColor = GetFromIni("Color", "Label BackColor", flw)
For i = 0 To Form1.imgicon.UBound
Form1.lblcaption(i).BackColor = MacButton7.BackColor
Next
MacButton9.BackColor = GetFromIni("Color", "Label ForeColor", flw)
For i = 0 To Form1.imgicon.UBound
Form1.lblcaption(i).ForeColor = MacButton9.BackColor
Next
For x = 0 To tasks.UBound
Form1.tasks(x).ForeColor = MacButton9.BackColor
Next
For x = 0 To Bag.UBound
Form1.Bag(x).ForeColor = MacButton9.BackColor
Next
For x = 0 To Cafe.UBound
Form1.Cafe(x).ForeColor = MacButton9.BackColor
Next
Bdr = GetFromIni("Color", "Border", flw)
If Bdr = "true" Then
Bdr = "False"
MacButton10_Click
Else
Bdr = "true"
MacButton10_Click
End If
MacButton11.BackColor = GetFromIni("Color", "Border Color", flw)
For i = 1 To Form1.imgicon.UBound
Form1.lblcaption(i).ForeColor = MacButton11.BackColor
Next
MacButton12.BackColor = GetFromIni("Color", "Menu BackColor", flw)
Form1.MENU1.Set_BackColor MacButton12.BackColor
Form1.MENU2.Set_BackColor MacButton12.BackColor

MacButton13.BackColor = GetFromIni("Color", "Menu ForeColor", flw)
Form1.MENU1.Set_ForeColor MacButton13.BackColor
Form1.MENU2.Set_ForeColor MacButton13.BackColor

MacButton14.BackColor = GetFromIni("Color", "Menu Border Color", flw)
Form1.Line1.BorderColor = MacButton14.BackColor
Form1.Line2.BorderColor = MacButton14.BackColor
Form1.Line3.BorderColor = MacButton14.BackColor
Form1.Line4.BorderColor = MacButton14.BackColor
str = GetFromIni("Clock", "Back", flw)
If Right(str, 10) = " <AppPath>" Then
str = App.path & "\" & Left(str, Len(str) - 10)
End If

Call aicAlphaImage4.LoadImage_FromFile(str)
Call Form1.aicAlphaImage1.LoadImage_FromFile(str)
aicAlphaImage4.Tag = str

str = GetFromIni("Clock", "Hour", flw)
If Right(str, 10) = " <AppPath>" Then
str = App.path & "\" & Left(str, Len(str) - 10)
End If

Call aicAlphaImage3.LoadImage_FromFile(str)
Call Form1.aicHour.LoadImage_FromFile(str)
aicAlphaImage3.Tag = str

str = GetFromIni("Clock", "Minute", flw)
If Right(str, 10) = " <AppPath>" Then
str = App.path & "\" & Left(str, Len(str) - 10)
End If

Call aicAlphaImage2.LoadImage_FromFile(str)
Call Form1.aicMinute.LoadImage_FromFile(str)
aicAlphaImage2.Tag = str

str = GetFromIni("Clock", "Second", flw)
If Right(str, 10) = " <AppPath>" Then
str = App.path & "\" & Left(str, Len(str) - 10)
End If

Call aicAlphaImage1.LoadImage_FromFile(str)
Call Form1.aicSecond.LoadImage_FromFile(str)
aicAlphaImage1.Tag = str

Posit = GetFromIni("Clock", "Position", flw)
If Posit = "N-W" Then
MacButton15_Click
ElseIf Posit = "North" Then
MacButton16_Click
ElseIf Posit = "N-E" Then
MacButton17_Click
ElseIf Posit = "West" Then
MacButton18_Click
ElseIf Posit = "East" Then
MacButton20_Click
ElseIf Posit = "S-W" Then
MacButton21_Click
ElseIf Posit = "South" Then
MacButton22_Click
ElseIf Posit = "S-E" Then
MacButton23_Click
End If
End Sub


