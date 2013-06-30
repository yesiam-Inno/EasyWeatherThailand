VERSION 5.00
Object = "{EB3B8C42-F8B9-4E48-8EE0-77E81D7B37FB}#1.0#0"; "lvButtons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_Config 
   Caption         =   "Form1"
   ClientHeight    =   10215
   ClientLeft      =   1080
   ClientTop       =   450
   ClientWidth     =   17655
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   17655
   Begin VB.Frame Frame1 
      Caption         =   "Config Data Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10005
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   17370
      Begin lvButtons.lvButtons_H cmd_OpenInput 
         Height          =   450
         Left            =   11235
         TabIndex        =   91
         Top             =   2520
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   794
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MSComDlg.CommonDialog Clg_File 
         Left            =   16545
         Top             =   2730
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Txt_Outputfile 
         Height          =   420
         Left            =   6210
         TabIndex        =   90
         Text            =   "EasyWeatherThai.txt"
         Top             =   3030
         Width           =   5010
      End
      Begin VB.TextBox Txt_Inputfile 
         Height          =   405
         Left            =   6195
         TabIndex        =   89
         Text            =   "EasyWeather.txt"
         Top             =   2550
         Width           =   4995
      End
      Begin lvButtons.lvButtons_H Cmd_Time 
         Height          =   345
         Left            =   16350
         TabIndex        =   88
         Top             =   3270
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   609
         Caption         =   "ทิศลม"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Timer Tmr_Direction 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   15930
         Top             =   3195
      End
      Begin MSComctlLib.ImageList ImgList_Direction 
         Left            =   15375
         Top             =   3075
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   275
         ImageHeight     =   275
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":0000
               Key             =   "N"
               Object.Tag             =   "N"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":98AD
               Key             =   "NNE"
               Object.Tag             =   "NNE"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":13092
               Key             =   "NE"
               Object.Tag             =   "NE"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":1C8A6
               Key             =   "ENE"
               Object.Tag             =   "ENE"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":2609F
               Key             =   "E"
               Object.Tag             =   "E"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":2F882
               Key             =   "ESE"
               Object.Tag             =   "ESE"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":39057
               Key             =   "SE"
               Object.Tag             =   "SE"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":4283E
               Key             =   "SSE"
               Object.Tag             =   "SSE"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":4C05B
               Key             =   "S"
               Object.Tag             =   "S"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":55865
               Key             =   "SSW"
               Object.Tag             =   "SSW"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":5F045
               Key             =   "SW"
               Object.Tag             =   "SW"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":68821
               Key             =   "WSW"
               Object.Tag             =   "WSW"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":72022
               Key             =   "W"
               Object.Tag             =   "W"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":7B91B
               Key             =   "WNW"
               Object.Tag             =   "WNW"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":850CE
               Key             =   "NW"
               Object.Tag             =   "NW"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Config.frx":8E91D
               Key             =   "NNW"
               Object.Tag             =   "NNW"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   6105
         ScaleHeight     =   3000
         ScaleWidth      =   6735
         TabIndex        =   74
         Top             =   6765
         Width           =   6735
         Begin VB.TextBox Txt_Press_Y 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1650
            TabIndex        =   75
            Text            =   "00000.00"
            Top             =   2370
            Width           =   3390
         End
         Begin VB.TextBox Txt_Press_X 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   1890
            Width           =   3390
         End
         Begin VB.TextBox Txt_Press_B 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   1455
            Width           =   3390
         End
         Begin VB.TextBox Txt_Press_A 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   975
            Width           =   3330
         End
         Begin lvButtons.lvButtons_H Cmd_Press 
            Height          =   1845
            Left            =   5130
            TabIndex        =   20
            Top             =   990
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   3254
            Caption         =   "คำนวณ Y "
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Lbl_Press_Add 
            Alignment       =   2  'Center
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5055
            TabIndex        =   87
            Top             =   450
            Width           =   495
         End
         Begin VB.Label Lbl_Press_B 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "  B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5535
            TabIndex        =   76
            Top             =   435
            Width           =   1155
         End
         Begin VB.Label Lbl_Press_val_y 
            BackColor       =   &H008080FF&
            Caption         =   "      Y ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   86
            Top             =   2355
            Width           =   4590
         End
         Begin VB.Label Lbl_Press_val_x 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Val. X ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   85
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Lbl_Press_X 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "( X )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3750
            TabIndex        =   84
            Top             =   465
            Width           =   1320
         End
         Begin VB.Label Lbl_Rain_ln 
            Alignment       =   2  'Center
            Caption         =   "ln"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3390
            TabIndex        =   83
            Top             =   465
            Width           =   360
         End
         Begin VB.Label Lbl_Press_A 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2175
            TabIndex        =   82
            Top             =   465
            Width           =   1215
         End
         Begin VB.Label Lbl_Press_Equal 
            Alignment       =   2  'Center
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            TabIndex        =   81
            Top             =   465
            Width           =   495
         End
         Begin VB.Label Lbl_Press_Y 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   80
            Top             =   465
            Width           =   1485
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pressure"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   90
            TabIndex        =   79
            Top             =   45
            Width           =   2295
         End
         Begin VB.Label Lbl_Press_val_b 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Val. B ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   78
            Top             =   1470
            Width           =   3375
         End
         Begin VB.Label Lbl_Press_val_a 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Val. A ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   77
            Top             =   975
            Width           =   3300
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   15
            Top             =   30
            Width           =   2475
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   6135
         ScaleHeight     =   3000
         ScaleWidth      =   6705
         TabIndex        =   61
         Top             =   3555
         Width           =   6705
         Begin VB.TextBox Txt_Rain_A 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   975
            Width           =   3270
         End
         Begin VB.TextBox Txt_Rain_B 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   1455
            Width           =   3270
         End
         Begin VB.TextBox Txt_Rain_X 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   1890
            Width           =   3270
         End
         Begin VB.TextBox Txt_Rain_Y 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1650
            TabIndex        =   62
            Text            =   "00000.00"
            Top             =   2370
            Width           =   3270
         End
         Begin lvButtons.lvButtons_H Cmd_Rain 
            Height          =   1845
            Left            =   5085
            TabIndex        =   16
            Top             =   975
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   3254
            Caption         =   "คำนวณ Y "
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Lbl_Rain_B 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "(  B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3990
            TabIndex        =   66
            Top             =   225
            Width           =   1275
         End
         Begin VB.Label Lbl_Rain_val_a 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Val. A ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   73
            Top             =   975
            Width           =   3300
         End
         Begin VB.Label Lbl_Rain_val_b 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Val. B ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   72
            Top             =   1470
            Width           =   3375
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Caption         =   "Rain"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   90
            TabIndex        =   71
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label Lbl_Rain_Y 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   165
            TabIndex        =   70
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Lbl_Rain_Equal 
            Alignment       =   2  'Center
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1635
            TabIndex        =   69
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Lbl_Rain_A 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2130
            TabIndex        =   68
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label Lbl_Rain_e 
            Alignment       =   2  'Center
            Caption         =   "e"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3555
            TabIndex        =   67
            Top             =   465
            Width           =   630
         End
         Begin VB.Label Lbl_Rain_X 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "X  )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5235
            TabIndex        =   65
            Top             =   225
            Width           =   1440
         End
         Begin VB.Label Lbl_Rain_val_x 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Val. X ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   64
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Lbl_Rain_val_y 
            BackColor       =   &H008080FF&
            Caption         =   "      Y ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   63
            Top             =   2355
            Width           =   4500
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   15
            Top             =   30
            Width           =   2475
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   210
         ScaleHeight     =   3000
         ScaleWidth      =   5715
         TabIndex        =   48
         Top             =   6720
         Width           =   5715
         Begin VB.TextBox Txt_Wind_Y 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1650
            TabIndex        =   49
            Text            =   "00000.00"
            Top             =   2370
            Width           =   2145
         End
         Begin VB.TextBox Txt_Wind_X 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   1890
            Width           =   2190
         End
         Begin VB.TextBox Txt_Wind_C 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   1455
            Width           =   2190
         End
         Begin VB.TextBox Txt_Wind_M 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   975
            Width           =   2190
         End
         Begin lvButtons.lvButtons_H Cmd_Wind 
            Height          =   1845
            Left            =   3930
            TabIndex        =   12
            Top             =   975
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   3254
            Caption         =   "คำนวณ Y "
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Lbl_Wind_val_y 
            BackColor       =   &H008080FF&
            Caption         =   "      Y ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   60
            Top             =   2355
            Width           =   3420
         End
         Begin VB.Label Lbl_Wind_val_x 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Val. X ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   59
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Lbl_Wind_C 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4245
            TabIndex        =   58
            Top             =   465
            Width           =   1215
         End
         Begin VB.Label Lbl_Wind_Add 
            Alignment       =   2  'Center
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3765
            TabIndex        =   57
            Top             =   465
            Width           =   495
         End
         Begin VB.Label Lbl_wind_X 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "(X)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2805
            TabIndex        =   56
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_Wind_M 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1845
            TabIndex        =   55
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_Wind_Equal 
            Alignment       =   2  'Center
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1365
            TabIndex        =   54
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Lbl_Wind_Y 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   165
            TabIndex        =   53
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Wind"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   90
            TabIndex        =   52
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label Lbl_Wind_val_c 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Val. C ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   51
            Top             =   1470
            Width           =   3375
         End
         Begin VB.Label Lbl_Wind_val_m 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Val. M ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   50
            Top             =   975
            Width           =   3300
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   15
            Top             =   30
            Width           =   2520
         End
      End
      Begin VB.PictureBox Pic_Hum 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   210
         ScaleHeight     =   3000
         ScaleWidth      =   5715
         TabIndex        =   35
         Top             =   3540
         Width           =   5715
         Begin VB.TextBox Txt_Hum_M 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   975
            Width           =   2190
         End
         Begin VB.TextBox Txt_Hum_C 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   1455
            Width           =   2190
         End
         Begin VB.TextBox Txt_Hum_X 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   1890
            Width           =   2190
         End
         Begin VB.TextBox Txt_Hum_Y 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1650
            TabIndex        =   36
            Text            =   "00000.00"
            Top             =   2370
            Width           =   2145
         End
         Begin lvButtons.lvButtons_H Cmd_Hum 
            Height          =   1845
            Left            =   3930
            TabIndex        =   8
            Top             =   975
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   3254
            Caption         =   "คำนวณ Y "
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Lbl_Hum_val_m 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Val. M ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   47
            Top             =   960
            Width           =   3300
         End
         Begin VB.Label Lbl_Hum_val_c 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Val. C ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   46
            Top             =   1470
            Width           =   3375
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Caption         =   "Humidity"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   90
            TabIndex        =   45
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label Lbl_Hum_Y 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   165
            TabIndex        =   44
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Lbl_Hum_Equal 
            Alignment       =   2  'Center
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1365
            TabIndex        =   43
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Lbl_Hum_M 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1845
            TabIndex        =   42
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_Hum_X 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "(X)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2805
            TabIndex        =   41
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_Hum_Add 
            Alignment       =   2  'Center
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3765
            TabIndex        =   40
            Top             =   465
            Width           =   495
         End
         Begin VB.Label Lbl_Hum_C 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4245
            TabIndex        =   39
            Top             =   465
            Width           =   1215
         End
         Begin VB.Label Lbl_Hum_val_x 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Val. X ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   38
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Lbl_Hum_val_y 
            BackColor       =   &H008080FF&
            Caption         =   "      Y ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   37
            Top             =   2355
            Width           =   3420
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   15
            Top             =   30
            Width           =   2520
         End
      End
      Begin lvButtons.lvButtons_H Cmd_SaveConfig 
         Height          =   555
         Left            =   15360
         TabIndex        =   26
         Top             =   615
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         Caption         =   "บันทึกค่ามาตรฐาน"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.PictureBox Pic_Temp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   210
         ScaleHeight     =   3000
         ScaleWidth      =   5715
         TabIndex        =   22
         Top             =   435
         Width           =   5715
         Begin VB.TextBox Txt_Temp_Y 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1665
            TabIndex        =   21
            Text            =   "00000.00"
            Top             =   2385
            Width           =   2190
         End
         Begin VB.TextBox Txt_Temp_X 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   1890
            Width           =   2190
         End
         Begin lvButtons.lvButtons_H Cmd_Temp 
            Height          =   1845
            Left            =   3930
            TabIndex        =   4
            Top             =   975
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   3254
            Caption         =   "คำนวณ Y "
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.TextBox Txt_Temp_C 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   2
            Text            =   "0.00"
            Top             =   1455
            Width           =   2190
         End
         Begin VB.TextBox Txt_Temp_M 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1650
            TabIndex        =   1
            Text            =   "0.00"
            Top             =   960
            Width           =   2190
         End
         Begin VB.Label Lbl_Temp_val_y 
            BackColor       =   &H008080FF&
            Caption         =   "      Y ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   450
            TabIndex        =   34
            Top             =   2370
            Width           =   3420
         End
         Begin VB.Label Lbl_Temp_val_x 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Val. X ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   33
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Lbl_Temp_C 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4245
            TabIndex        =   32
            Top             =   465
            Width           =   1215
         End
         Begin VB.Label Lbl_Temp_Add 
            Alignment       =   2  'Center
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3765
            TabIndex        =   31
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Lbl_Temp_X 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            Caption         =   "(X)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2805
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_Temp_M 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1845
            TabIndex        =   29
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_Temp_Equal 
            Alignment       =   2  'Center
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1365
            TabIndex        =   28
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Lbl_Temp_Y 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   165
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Caption         =   "Temperature"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   60
            Width           =   2295
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   15
            Top             =   30
            Width           =   2475
         End
         Begin VB.Label Lbl_Temp_val_c 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Val. C ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   24
            Top             =   1470
            Width           =   3375
         End
         Begin VB.Label Lbl_Temp_val_m 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Val. M ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   480
            TabIndex        =   23
            Top             =   975
            Width           =   3300
         End
      End
      Begin lvButtons.lvButtons_H lvButtons_H1 
         Height          =   450
         Left            =   11220
         TabIndex        =   92
         Top             =   3030
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   794
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Date / Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6330
         TabIndex        =   94
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label Lab_DateTime 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6270
         TabIndex        =   93
         Top             =   825
         Width           =   4980
      End
      Begin VB.Image Img_Direction 
         Height          =   4125
         Left            =   13035
         Picture         =   "Frm_Config.frx":9818E
         Top             =   3615
         Width           =   4125
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   6270
         Top             =   375
         Width           =   2475
      End
   End
End
Attribute VB_Name = "Frm_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cur_i As Integer

Private Sub Cmd_Time_Click()
         Tmr_Direction.Enabled = True
         Cur_i = 1
End Sub

Private Sub Tmr_Direction_Timer()
      If Cur_i = 17 Then
             Tmr_Direction.Enabled = False
             Cur_i = 1
      Else
              Img_Direction.Picture = ImgList_Direction.ListImages(Cur_i).Picture
              Cur_i = Cur_i + 1
      End If
End Sub

Private Sub Cmd_SaveConfig_Click()
' ########## Save Data for Temperature  #########
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_TEMP_M", Txt_Temp_M.Text
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_TEMP_C", Txt_Temp_C.Text
' ########## Save Data for Humidity  #########
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_HUM_M", Txt_Hum_M.Text
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_HUM_C", Txt_Hum_C.Text
' ########## Save Data for Wind  #########
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_WIND_M", Txt_Wind_M.Text
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_WIND_C", Txt_Wind_C.Text
' ########## Save Data for Rain  #########
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_RAIN_A", Txt_Rain_A.Text
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_RAIN_B", Txt_Rain_B.Text
' ########## Save Data for Pressure  #########
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_PRESS_A", Txt_Press_A.Text
     WriteRegFile App.Path & "\" & CONFIG_FILENAME, "SETTING", "P_PRESS_B", Txt_Press_B.Text
     MsgBox "ระบบได้ทำการบันทึกเรียบร้อยแล้ว", vbOKOnly
End Sub

' #####################################  Temperature  ##########################################
Private Sub Cmd_Temp_Click()
         Txt_Temp_Y.Text = Cal_Linear(CDbl(Txt_Temp_M.Text), CDbl(Txt_Temp_C.Text), CDbl(Txt_Temp_X.Text))
          Lbl_Temp_Y.Caption = Format(Txt_Temp_Y.Text, "#####0.00000")
End Sub

Private Sub Form_Load()
Call Main
' ########## Show Data for Temperature  #########
Txt_Temp_M.Text = Format(p_Temp_M, "#####0.0000")
Txt_Temp_C.Text = Format(p_Temp_C, "#####0.0000")
' ########## Show Data for Humidity  #########
Txt_Hum_M.Text = Format(p_Hum_M, "#####0.0000")
Txt_Hum_C.Text = Format(p_Hum_C, "#####0.0000")
' ########## Show Data for Wind  #########
Txt_Wind_M.Text = Format(p_Wind_M, "#####0.0000")
Txt_Wind_C.Text = Format(p_Wind_C, "#####0.0000")
' ########## Show Data for Rain  #########
Txt_Rain_A.Text = Format(p_Rain_A, "#####0.0000")
Txt_Rain_B.Text = Format(p_Rain_B, "#####0.0000")
' ########## Show Data for Pressure  #########
Txt_Press_A.Text = Format(p_Press_A, "#####0.0000")
Txt_Press_B.Text = Format(p_Press_B, "#####0.0000")

Dim tmpstr As String
    
    Open App.Path & "\" & Txt_Inputfile.Text For Input As #1
        If Err = 0 Then
            Do Until EOF(1)
                Line Input #1, tmpstr
                                MsgBox "ข้อมูล" & tmpstr, vbOKOnly
            Loop
        End If
    Close #1
End Sub

Private Sub Txt_Temp_C_Validate(Cancel As Boolean)
            Lbl_Temp_C.Caption = Format(Txt_Temp_C.Text, "#####0.0000")
            Txt_Temp_C.Text = Format(Txt_Temp_C.Text, "#####0.0000")
End Sub

Private Sub Txt_Temp_M_Validate(Cancel As Boolean)
            Lbl_Temp_M.Caption = Format(Txt_Temp_M.Text, "#####0.0000")
            Txt_Temp_M.Text = Format(Txt_Temp_M.Text, "#####0.0000")
End Sub

Private Sub Txt_Temp_X_Validate(Cancel As Boolean)
            Lbl_Temp_X.Caption = "(" & Format(Txt_Temp_X.Text, "#####0.0000") & ")"
            Txt_Temp_X.Text = Format(Txt_Temp_X.Text, "#####0.0000")
End Sub

Private Sub Txt_Temp_M_GotFocus()
            Txt_Temp_M.SelStart = 0
            Txt_Temp_M.SelLength = Len(Txt_Temp_M.Text)
End Sub

Private Sub Txt_Temp_C_GotFocus()
            Txt_Temp_C.SelStart = 0
            Txt_Temp_C.SelLength = Len(Txt_Temp_C.Text)
End Sub

Private Sub Txt_Temp_X_GotFocus()
            Txt_Temp_X.SelStart = 0
            Txt_Temp_X.SelLength = Len(Txt_Temp_X.Text)
End Sub
' #####################################  Humidity  ##########################################
Private Sub Cmd_Hum_Click()
         Txt_Hum_Y.Text = Cal_Linear(CDbl(Txt_Hum_M.Text), CDbl(Txt_Hum_C.Text), CDbl(Txt_Hum_X.Text))
          Lbl_Hum_Y.Caption = Format(Txt_Hum_Y.Text, "#####0.0000")
End Sub

Private Sub Txt_Hum_C_Validate(Cancel As Boolean)
            Lbl_Hum_C.Caption = Format(Txt_Hum_C.Text, "#####0.0000")
            Txt_Hum_C.Text = Format(Txt_Hum_C.Text, "#####0.0000")
End Sub

Private Sub Txt_Hum_M_Validate(Cancel As Boolean)
            Lbl_Hum_M.Caption = Format(Txt_Hum_M.Text, "#####0.0000")
            Txt_Hum_M.Text = Format(Txt_Hum_M.Text, "#####0.0000")
End Sub

Private Sub Txt_Hum_X_Validate(Cancel As Boolean)
            Lbl_Hum_X.Caption = "(" & Format(Txt_Hum_X.Text, "#####0.0000") & ")"
            Txt_Hum_X.Text = Format(Txt_Hum_X.Text, "#####0.0000")
End Sub

Private Sub Txt_Hum_M_GotFocus()
            Txt_Hum_M.SelStart = 0
            Txt_Hum_M.SelLength = Len(Txt_Hum_M.Text)
End Sub

Private Sub Txt_Hum_C_GotFocus()
            Txt_Hum_C.SelStart = 0
            Txt_Hum_C.SelLength = Len(Txt_Hum_C.Text)
End Sub

Private Sub Txt_Hum_X_GotFocus()
            Txt_Hum_X.SelStart = 0
            Txt_Hum_X.SelLength = Len(Txt_Hum_X.Text)
End Sub

' #####################################  Wind  ##########################################
Private Sub Cmd_Wind_Click()
         Txt_Wind_Y.Text = Cal_Linear(CDbl(Txt_Wind_M.Text), CDbl(Txt_Wind_C.Text), CDbl(Txt_Wind_X.Text))
          Lbl_Wind_Y.Caption = Format(Txt_Wind_Y.Text, "#####0.0000")
End Sub

Private Sub Txt_Wind_C_Validate(Cancel As Boolean)
            Lbl_Wind_C.Caption = Format(Txt_Wind_C.Text, "#####0.0000")
            Txt_Wind_C.Text = Format(Txt_Wind_C.Text, "#####0.0000")
End Sub

Private Sub Txt_Wind_M_Validate(Cancel As Boolean)
            Lbl_Wind_M.Caption = Format(Txt_Wind_M.Text, "#####0.0000")
            Txt_Wind_M.Text = Format(Txt_Wind_M.Text, "#####0.0000")
End Sub

Private Sub Txt_Wind_X_Validate(Cancel As Boolean)
            Lbl_wind_X.Caption = "(" & Format(Txt_Wind_X.Text, "#####0.0000") & ")"
            Txt_Wind_X.Text = Format(Txt_Wind_X.Text, "#####0.0000")
End Sub

Private Sub Txt_Wind_M_GotFocus()
            Txt_Wind_M.SelStart = 0
            Txt_Wind_M.SelLength = Len(Txt_Wind_M.Text)
End Sub

Private Sub Txt_Wind_C_GotFocus()
            Txt_Wind_C.SelStart = 0
            Txt_Wind_C.SelLength = Len(Txt_Wind_C.Text)
End Sub

Private Sub Txt_Wind_X_GotFocus()
            Txt_Wind_X.SelStart = 0
            Txt_Wind_X.SelLength = Len(Txt_Wind_X.Text)
End Sub


' #####################################  Rain  ##########################################
Private Sub Cmd_Rain_Click()
         Txt_Rain_Y.Text = Cal_Exponential(CDbl(Txt_Rain_A.Text), CDbl(Txt_Rain_B.Text), CDbl(Txt_Rain_X.Text))
          Lbl_Rain_Y.Caption = Format(Txt_Rain_Y.Text, "#####0.0000")
End Sub

Private Sub Txt_Rain_B_Validate(Cancel As Boolean)
            Lbl_Rain_B.Caption = Format(Txt_Rain_B.Text, "#####0.0000")
            Txt_Rain_B.Text = Format(Txt_Rain_B.Text, "#####0.0000")
End Sub

Private Sub Txt_Rain_A_Validate(Cancel As Boolean)
            Lbl_Rain_A.Caption = Format(Txt_Rain_A.Text, "#####0.0000")
            Txt_Rain_A.Text = Format(Txt_Rain_A.Text, "#####0.0000")
End Sub

Private Sub Txt_Rain_X_Validate(Cancel As Boolean)
            Lbl_Rain_X.Caption = "(" & Format(Txt_Rain_X.Text, "#####0.0000") & ")"
            Txt_Rain_X.Text = Format(Txt_Rain_X.Text, "#####0.0000")
End Sub

Private Sub Txt_Rain_A_GotFocus()
            Txt_Rain_A.SelStart = 0
            Txt_Rain_A.SelLength = Len(Txt_Rain_A.Text)
End Sub

Private Sub Txt_Rain_B_GotFocus()
            Txt_Rain_B.SelStart = 0
            Txt_Rain_B.SelLength = Len(Txt_Rain_B.Text)
End Sub

Private Sub Txt_Rain_X_GotFocus()
            Txt_Rain_X.SelStart = 0
            Txt_Rain_X.SelLength = Len(Txt_Rain_X.Text)
End Sub

' #####################################  Pressure  ##########################################
Private Sub Cmd_Press_Click()
         Txt_Press_Y.Text = Cal_Logarithm(CDbl(Txt_Press_A.Text), CDbl(Txt_Press_B.Text), CDbl(Txt_Press_X.Text))
          Lbl_Press_Y.Caption = Format(Txt_Press_Y.Text, "#####0.0000")
End Sub

Private Sub Txt_Press_B_Validate(Cancel As Boolean)
            Lbl_Press_B.Caption = Format(Txt_Press_B.Text, "#####0.0000")
            Txt_Press_B.Text = Format(Txt_Press_B.Text, "#####0.0000")
End Sub

Private Sub Txt_Press_A_Validate(Cancel As Boolean)
            Lbl_Press_A.Caption = Format(Txt_Press_A.Text, "#####0.0000")
            Txt_Press_A.Text = Format(Txt_Press_A.Text, "#####0.0000")
End Sub

Private Sub Txt_Press_X_Validate(Cancel As Boolean)
            Lbl_Press_X.Caption = "(" & Format(Txt_Press_X.Text, "#####0.0000") & ")"
            Txt_Press_X.Text = Format(Txt_Press_X.Text, "#####0.0000")
End Sub

Private Sub Txt_Press_A_GotFocus()
            Txt_Press_A.SelStart = 0
            Txt_Press_A.SelLength = Len(Txt_Press_A.Text)
End Sub

Private Sub Txt_Press_B_GotFocus()
            Txt_Press_B.SelStart = 0
            Txt_Press_B.SelLength = Len(Txt_Press_B.Text)
End Sub

Private Sub Txt_Press_X_GotFocus()
            Txt_Press_X.SelStart = 0
            Txt_Press_X.SelLength = Len(Txt_Press_X.Text)
End Sub
