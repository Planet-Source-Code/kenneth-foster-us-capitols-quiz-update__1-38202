VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "US Capitols Quiz Box"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   570
      TabIndex        =   64
      Top             =   1995
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.PictureBox Picture1 
      Height          =   4980
      Left            =   3480
      Picture         =   "US Caps QuizBox.frx":0000
      ScaleHeight     =   4920
      ScaleWidth      =   7500
      TabIndex        =   13
      Top             =   -255
      Width           =   7560
      Begin VB.ListBox lstCapsLIst 
         Columns         =   3
         Height          =   3570
         Left            =   -15
         Sorted          =   -1  'True
         TabIndex        =   69
         Top             =   240
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.CommandButton cmdSpell 
         Caption         =   "Capitols List"
         Height          =   300
         Left            =   75
         TabIndex        =   68
         Top             =   4545
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Easy"
         Height          =   300
         Left            =   1935
         TabIndex        =   66
         Top             =   4545
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hard"
         Height          =   300
         Left            =   1215
         TabIndex        =   65
         Top             =   4545
         Width           =   645
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Levels"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1605
         TabIndex        =   67
         Top             =   4290
         Width           =   630
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Wyoming"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   49
         Left            =   2010
         TabIndex        =   63
         Top             =   1755
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Winconsin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   48
         Left            =   4095
         TabIndex        =   62
         Top             =   1275
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Washington"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   47
         Left            =   285
         TabIndex        =   61
         Top             =   570
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "W.VIr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   46
         Left            =   5835
         TabIndex        =   60
         Top             =   2175
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Vermont"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   45
         Left            =   6075
         TabIndex        =   59
         Top             =   705
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Virginia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   44
         Left            =   6015
         TabIndex        =   58
         Top             =   2415
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Utah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   43
         Left            =   1605
         TabIndex        =   57
         Top             =   2340
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Texas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   42
         Left            =   3255
         TabIndex        =   56
         Top             =   3750
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Tennessee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   41
         Left            =   4815
         TabIndex        =   55
         Top             =   2835
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "S.Dakota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   40
         Left            =   2925
         TabIndex        =   54
         Top             =   1455
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "S.Carolina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   39
         Left            =   6075
         TabIndex        =   53
         Top             =   3090
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "R.I."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   38
         Left            =   7170
         TabIndex        =   52
         Top             =   1455
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Penn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   37
         Left            =   5970
         TabIndex        =   51
         Top             =   1740
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Oregon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   36
         Left            =   465
         TabIndex        =   50
         Top             =   1305
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Oklahoma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   35
         Left            =   3240
         TabIndex        =   49
         Top             =   2970
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Ohio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   34
         Left            =   5400
         TabIndex        =   48
         Top             =   2070
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "N.Dakota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   33
         Left            =   2910
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "N.Carolina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   32
         Left            =   6060
         TabIndex        =   46
         Top             =   2685
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "New York"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   31
         Left            =   6195
         TabIndex        =   45
         Top             =   1140
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "New Mexico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   30
         Left            =   2205
         TabIndex        =   44
         Top             =   3000
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "N.Jer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   29
         Left            =   6975
         TabIndex        =   43
         Top             =   1890
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "N.Ham"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   28
         Left            =   6810
         TabIndex        =   42
         Top             =   990
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Nevado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   735
         TabIndex        =   41
         Top             =   2115
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Nebraska"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   26
         Left            =   2940
         TabIndex        =   40
         Top             =   1935
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Montana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   25
         Left            =   1800
         TabIndex        =   39
         Top             =   930
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Missouri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   24
         Left            =   4110
         TabIndex        =   38
         Top             =   2490
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Miss"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   23
         Left            =   4665
         TabIndex        =   37
         Top             =   3495
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Minnesoda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   22
         Left            =   3750
         TabIndex        =   36
         Top             =   780
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Michigan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   21
         Left            =   5025
         TabIndex        =   35
         Top             =   1530
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Mass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   6975
         TabIndex        =   34
         Top             =   1230
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Maryland"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   19
         Left            =   6690
         TabIndex        =   33
         Top             =   2295
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Maine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   6870
         TabIndex        =   32
         Top             =   510
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Louisiana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   17
         Left            =   4320
         TabIndex        =   31
         Top             =   4065
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Kentucky"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   16
         Left            =   4980
         TabIndex        =   30
         Top             =   2535
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Kansas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   15
         Left            =   3270
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Iowa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   14
         Left            =   3990
         TabIndex        =   28
         Top             =   1815
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Indiana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   13
         Left            =   4860
         TabIndex        =   27
         Top             =   1860
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Illinios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   4470
         TabIndex        =   26
         Top             =   2160
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Idaho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   11
         Left            =   1290
         TabIndex        =   25
         Top             =   1500
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Hawaii"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   285
         TabIndex        =   24
         Top             =   4080
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Georgia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   9
         Left            =   5535
         TabIndex        =   23
         Top             =   3450
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Florida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   5820
         TabIndex        =   22
         Top             =   3960
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Dela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   7
         Left            =   7005
         TabIndex        =   21
         Top             =   2100
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Conn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   7005
         TabIndex        =   20
         Top             =   1665
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Colorado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2250
         TabIndex        =   19
         Top             =   2385
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Californa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   285
         TabIndex        =   18
         Top             =   2820
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Arkansas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   4020
         TabIndex        =   17
         Top             =   3165
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Arizonia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1335
         TabIndex        =   16
         Top             =   3075
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Alaska"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   285
         TabIndex        =   15
         Top             =   3780
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Alabama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   4920
         TabIndex        =   14
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3330
      Width           =   720
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0080FFFF&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3330
      Width           =   1080
   End
   Begin VB.CommandButton cmdNewQuiz 
      BackColor       =   &H0080FF80&
      Caption         =   "New Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3330
      Width           =   765
   End
   Begin VB.TextBox txtScoreInc 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2175
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   435
      Width           =   525
   End
   Begin VB.TextBox txtScoreCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   705
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   435
      Width           =   525
   End
   Begin VB.TextBox txtCorrect 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1035
      Width           =   1995
   End
   Begin VB.ListBox lstCaps 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2010
      Width           =   2310
   End
   Begin VB.Shape Shape2 
      Height          =   765
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   3930
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      Height          =   885
      Left            =   285
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   2970
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1710
      TabIndex        =   12
      Top             =   4275
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Capitols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1920
      TabIndex        =   11
      Top             =   3930
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "US"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1320
      TabIndex        =   10
      Top             =   3930
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   105
      Picture         =   "US Caps QuizBox.frx":C2AE
      Top             =   3930
      Width           =   1020
   End
   Begin VB.Label lblInc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Incorrect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1755
      TabIndex        =   6
      Top             =   60
      Width           =   1425
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Correct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   390
      TabIndex        =   5
      Top             =   60
      Width           =   1110
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   225
      TabIndex        =   1
      Top             =   1455
      Width           =   3045
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(49)
Dim B(49)
Dim p As Integer
Dim w As Integer
Dim MaxNumber As Integer
Dim seq As Integer
Dim ChosenNumber As Integer
Dim MainLoop As Integer
Dim X As Integer
Dim Y As Integer
Public n As Integer
Const s As Integer = 49
Dim sStates(s) As String
Dim sCaps(s) As String

Private Sub cmdNewQuiz_Click()
 
Dim z As Integer
For z = 0 To 49
   lblState(z).Visible = False
   lblState(z).ForeColor = vbBlack
Next z

 cmdStart.Enabled = True
    w = 0
    X = 0
    Y = 0
    Call Array_Fill
    txtScoreCor.Text = ""
    txtScoreInc.Text = ""
    txtCorrect.Text = ""
    lblText.Caption = ""
    lstCaps.Clear
    cmdStart.Visible = True
End Sub

Private Sub cmdQuit_Click()
txtCorrect.Text = ""
    Unload Me
End Sub

Private Sub cmdSpell_Click()
Dim e As Integer
lstCapsLIst.Clear
lstCapsLIst.Visible = True
For e = 0 To 49
  lstCapsLIst.AddItem sCaps(e)
Next e

End Sub

Private Sub cmdStart_Click()

Call Make_Quiz
cmdStart.Visible = False
cmdNewQuiz.Visible = True
If Text1.Visible = True Then
Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
lstCaps.Visible = False
cmdSpell.Visible = True
Text1.Visible = True
Text1.SetFocus

End Sub

Private Sub Command2_Click()
lstCaps.Visible = True
Text1.Visible = False
cmdSpell.Visible = False

End Sub

Private Sub Form_Load()
cmdNewQuiz.Visible = False
MakeWindow Me, True
    Call Array_Fill
    cmdStart.Caption = "Start"
    w = 0
    sStates(0) = "Alabama"  ' US states-------------
    sStates(1) = "Alaska"
    sStates(2) = "Arizonia"
    sStates(3) = "Arkansas"
    sStates(4) = "California"
    sStates(5) = "Colorado"
    sStates(6) = "Connecticut"
    sStates(7) = "Delaware"
    sStates(8) = "Florida"
    sStates(9) = "Georgia"
    sStates(10) = "Hawaii"
    sStates(11) = "Idaho"
    sStates(12) = "Illinois"
    sStates(13) = "Indiana"
    sStates(14) = "Iowa"
    sStates(15) = "Kansas"
    sStates(16) = "Kentucky"
    sStates(17) = "Louisiana"
    sStates(18) = "Maine"
    sStates(19) = "Maryland"
    sStates(20) = "Massachusetts"
    sStates(21) = "Michigan"
    sStates(22) = "Minnesota"
    sStates(23) = "Mississippi"
    sStates(24) = "Missouri"
    sStates(25) = "Montana"
    sStates(26) = "Nebraska"
    sStates(27) = "Nevada"
    sStates(28) = "New Hampshire"
    sStates(29) = "New Jersy"
    sStates(30) = "New Mexico"
    sStates(31) = "New York"
    sStates(32) = "No. Carolina"
    sStates(33) = "No. Dakota"
    sStates(34) = "Ohio"
    sStates(35) = "Oklahoma"
    sStates(36) = "Oregon"
    sStates(37) = "Pennsylvania"
    sStates(38) = "Rhode Island"
    sStates(39) = "So. Carolina"
    sStates(40) = "So. Dakota"
    sStates(41) = "Tennessee"
    sStates(42) = "Texas"
    sStates(43) = "Utah"
    sStates(44) = "Virginia"
    sStates(45) = "Vermont"
    sStates(46) = "West Virginia"
    sStates(47) = "Washington"
    sStates(48) = "Wisconsin"
    sStates(49) = "Wyoming"

    sCaps(0) = "Montgomery"   'US capitols-----------
    sCaps(1) = "Juneau"
    sCaps(2) = "Phoenix"
    sCaps(3) = "Little Rock"
    sCaps(4) = "Sacramento"
    sCaps(5) = "Denver"
    sCaps(6) = "Hartford"
    sCaps(7) = "Dover"
    sCaps(8) = "Tallahassee"
    sCaps(9) = "Atlanta"
    sCaps(10) = "Honolulu"
    sCaps(11) = "Boise"
    sCaps(12) = "Springfield"
    sCaps(13) = "Indianapolis"
    sCaps(14) = "Des Moines"
    sCaps(15) = "Topeka"
    sCaps(16) = "Frankfort"
    sCaps(17) = "Baton Rouge"
    sCaps(18) = "Augusta"
    sCaps(19) = "Annopolis"
    sCaps(20) = "Boston"
    sCaps(21) = "Lansing"
    sCaps(22) = "St. Paul"
    sCaps(23) = "Jackson"
    sCaps(24) = "Jefferson City"
    sCaps(25) = "Helena"
    sCaps(26) = "Lincoln"
    sCaps(27) = "Carson City"
    sCaps(28) = "Concord"
    sCaps(29) = "Trenton"
    sCaps(30) = "Santa Fe"
    sCaps(31) = "Albany"
    sCaps(32) = "Raleigh"
    sCaps(33) = "Bismarck"
    sCaps(34) = "Columbus"
    sCaps(35) = "Oklahoma City"
    sCaps(36) = "Salem"
    sCaps(37) = "Harrisburg"
    sCaps(38) = "Providence"
    sCaps(39) = "Columbia"
    sCaps(40) = "Pierre"
    sCaps(41) = "Nashville"
    sCaps(42) = "Austin"
    sCaps(43) = "Salt Lake City"
    sCaps(44) = "Richmond"
    sCaps(45) = "Montelier"
    sCaps(46) = "Charleston"
    sCaps(47) = "Olympia"
    sCaps(48) = "Madison"
    sCaps(49) = "Cheyenne"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub lstCaps_Click()

    If lstCaps.Text = sCaps(p) Then
        txtCorrect.Text = "Correct"
        lblState(p).ForeColor = vbBlue
        txtCorrect.BackColor = vbWhite
        lblText.Caption = ""
        lstCaps.Clear
        X = X + 1
        txtScoreCor.Text = X
    Else
        txtCorrect.Text = sCaps(p)
        txtCorrect.BackColor = vbRed
        lblState(p).ForeColor = vbRed
        lblText.Caption = ""
        lstCaps.Clear
        Y = Y + 1
        txtScoreInc.Text = Y
    End If
        Call Make_Quiz
  End Sub
Private Sub Array_Fill()
    ' random non repeat numbers array

    'Set the original array
    MaxNumber = 49
    For seq = 0 To MaxNumber
        a(seq) = seq
    Next seq
    'Main Loop (mix the numbers all up)
    Randomize (Timer)
    For MainLoop = MaxNumber To 0 Step -1
        ChosenNumber = Int(MainLoop * Rnd)
        B(MaxNumber - MainLoop) = a(ChosenNumber)
        a(ChosenNumber) = a(MainLoop)
    Next MainLoop
   
End Sub
Private Sub Make_Quiz()
Dim m As Integer
    Dim t As Integer
    Dim r As Integer
    lstCaps.Clear
    
    Randomize
    m = Int(s * Rnd) 'picks the three other capitols
    t = Int(s * Rnd)
    r = Int(s * Rnd)
    '========================================================
    If m = t Or m = r Then m = m + 1
    If t = m Or t = r Then t = t + 1
    If r = m Or r = t Then r = r + 1
    '=====================================================
    p = B(w)

    lblText.Caption = sStates(p)
    lblState(p).Visible = True
    lstCaps.AddItem sCaps(p) 'capitol that matches state
    Select Case p
      Case m: m = m + 2
      Case t: t = t + 2
      Case r: r = r + 2
    End Select
    ' if a choice is out of range then correct
    If m > s Then m = m - 6
    If t > s Then t = t - 20
    If r > s Then r = r - 45
    '======================================================
        
    lstCaps.AddItem sCaps(m) 'The three other capitol choices
    lstCaps.AddItem sCaps(t)
    lstCaps.AddItem sCaps(r)
    If w = 49 Then GoTo Here:
    w = w + 1 ' increment to get next state
Here:
    If X + Y = 50 Then
        lblText.Caption = "End of Quiz"
        cmdStart.Enabled = False
        lstCaps.Clear
    End If

End Sub

Private Sub lstCapsLIst_Click()
Text1.Text = lstCapsLIst.Text
lstCapsLIst.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Text_Box
End If

End Sub
Private Sub Text_Box()
If Text1.Text = sCaps(p) Then
        txtCorrect.Text = "Correct"
        lblState(p).ForeColor = vbBlue
        txtCorrect.BackColor = vbWhite
        lblText.Caption = ""
        Text1.Text = ""
        X = X + 1
        txtScoreCor.Text = X
    Else
        txtCorrect.Text = sCaps(p)
        txtCorrect.BackColor = vbRed
        lblState(p).ForeColor = vbRed
        lblText.Caption = ""
        Text1.Text = ""
        Y = Y + 1
        txtScoreInc.Text = Y
    End If
   Call Make_Quiz
End Sub
