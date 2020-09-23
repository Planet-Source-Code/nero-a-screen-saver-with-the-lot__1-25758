VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Settings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charlie Set Up"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTAB 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Index           =   1
      Left            =   330
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   356
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1245
      Width           =   5340
      Begin VB.Frame fraTextPosition 
         Caption         =   "Text Position"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1650
         TabIndex        =   6
         Top             =   480
         Width           =   2040
         Begin VB.OptionButton optBelow 
            Caption         =   "Below Picture"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   150
            TabIndex        =   8
            Top             =   540
            Width           =   1740
         End
         Begin VB.OptionButton optAbove 
            Caption         =   "Above Picture"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1740
         End
      End
      Begin VB.TextBox txtDISPLAY 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   750
         TabIndex        =   5
         Text            =   "Charlie II Screen Saver"
         Top             =   0
         Width           =   4590
      End
      Begin VB.Frame fraSPEED 
         Caption         =   "Animation Speed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   4815
         Begin MSComctlLib.Slider sldSPEED 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   210
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   10
            Max             =   100
            SelStart        =   50
            TickFrequency   =   5
            Value           =   50
         End
         Begin VB.Label lblSLOW 
            Caption         =   "Slow"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   11
            Top             =   675
            Width           =   615
         End
         Begin VB.Label lblFAST 
            Alignment       =   1  'Right Justify
            Caption         =   "Fast"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   12
            Top             =   675
            Width           =   615
         End
      End
      Begin VB.Label lblTEXT 
         BackStyle       =   0  'Transparent
         Caption         =   "Text :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   45
         Width           =   675
      End
   End
   Begin VB.PictureBox picTAB 
      BorderStyle     =   0  'None
      Height          =   2925
      Index           =   2
      Left            =   330
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   356
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   5340
      Begin VB.TextBox txtFONTS 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   765
         TabIndex        =   14
         Text            =   "txtFONTS"
         Top             =   120
         Width           =   3810
      End
      Begin MSComctlLib.ListView lsvFONTS 
         Height          =   2100
         Left            =   750
         TabIndex        =   16
         Top             =   435
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   3704
         View            =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         TextBackground  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlFONTS"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.ListBox lstFONTS 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox chkUNDERLINE 
         Caption         =   "Underline"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3390
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1200
      End
      Begin VB.CheckBox chkITALIC 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2175
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1200
      End
      Begin VB.CheckBox chkBOLD 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   750
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1200
      End
      Begin MSComctlLib.ImageList imlFONTS 
         Left            =   4440
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Settings.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Settings.frx":11A6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTAB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2925
      Index           =   3
      Left            =   330
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   356
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   5340
      Begin VB.PictureBox picTEXT_CLR 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2400
         Left            =   75
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   346
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   525
         Width           =   5190
         Begin VB.PictureBox picTD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   6
            Left            =   4365
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1770
            Width           =   780
         End
         Begin VB.PictureBox picTD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   5
            Left            =   4365
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1035
            Width           =   780
         End
         Begin VB.PictureBox picTD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   4
            Left            =   3435
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1770
            Width           =   780
         End
         Begin VB.PictureBox picTD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   3
            Left            =   3435
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1035
            Width           =   780
         End
         Begin VB.PictureBox picTD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   2
            Left            =   2505
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1770
            Width           =   780
         End
         Begin VB.PictureBox picTD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   1
            Left            =   2505
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1035
            Width           =   780
         End
         Begin VB.PictureBox picTC_Colour 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   3930
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   81
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   315
            Width           =   1215
         End
         Begin VB.PictureBox picTC_Colour 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   2505
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   81
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   315
            Width           =   1215
         End
         Begin VB.Frame fraTC_STYLE 
            Caption         =   "Text Gradient Style"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2400
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   2250
            Begin VB.OptionButton optTCS 
               Caption         =   "&Square"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   30
               Top             =   1965
               Width           =   1875
            End
            Begin VB.OptionButton optTCS 
               Caption         =   "Diagonal &Down"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   29
               Top             =   1590
               Width           =   1875
            End
            Begin VB.OptionButton optTCS 
               Caption         =   "Diagonal &Up"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   28
               Top             =   1290
               Width           =   1875
            End
            Begin VB.OptionButton optTCS 
               Caption         =   "&Vertical"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   27
               Top             =   915
               Width           =   1875
            End
            Begin VB.OptionButton optTCS 
               Caption         =   "&Horizontal"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Top             =   615
               Width           =   1875
            End
            Begin VB.OptionButton optTCS 
               Caption         =   "&One Colour"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   1875
            End
         End
         Begin VB.Label lblTDrctn 
            Alignment       =   2  'Center
            Caption         =   "Text Gradient Direction"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2505
            TabIndex        =   35
            Top             =   720
            Width           =   2640
         End
         Begin VB.Label lblTC 
            Alignment       =   2  'Center
            Caption         =   "Text Colour 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   3930
            TabIndex        =   32
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblTC 
            Alignment       =   2  'Center
            Caption         =   "Text Colour 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   2505
            TabIndex        =   31
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox picTEXT_PIC 
         BorderStyle     =   0  'None
         Height          =   2400
         Left            =   75
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   346
         TabIndex        =   42
         Top             =   525
         Visible         =   0   'False
         Width           =   5190
         Begin VB.PictureBox picTP_Preview 
            AutoRedraw      =   -1  'True
            Height          =   1860
            Left            =   2730
            ScaleHeight     =   120
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   160
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   0
            Width           =   2460
         End
         Begin VB.Frame fraTP_Ptype 
            Caption         =   "Picture Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   2430
            Begin VB.OptionButton optTP_PicType 
               Caption         =   "Use &Desktop"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   45
               Top             =   600
               Width           =   2175
            End
            Begin VB.OptionButton optTP_PicType 
               Caption         =   "Use &Background"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   44
               Top             =   270
               Width           =   2175
            End
            Begin VB.OptionButton optTP_PicType 
               Caption         =   "&User Specified"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   46
               Top             =   930
               Width           =   2175
            End
         End
         Begin VB.Frame fraTP_Fit 
            Caption         =   "Picture Fit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   0
            TabIndex        =   47
            Top             =   1395
            Width           =   2430
            Begin VB.OptionButton optTP_PicFit 
               Caption         =   "Ce&ntre"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   48
               Top             =   270
               Value           =   -1  'True
               Width           =   1080
            End
            Begin VB.OptionButton optTP_PicFit 
               Caption         =   "&Tile"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1290
               TabIndex        =   49
               Top             =   270
               Width           =   1080
            End
            Begin VB.OptionButton optTP_PicFit 
               Caption         =   "&Stretch"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   50
               Top             =   600
               Width           =   1080
            End
            Begin VB.OptionButton optTP_PicFit 
               Caption         =   "&Zoom"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   1290
               TabIndex        =   51
               Top             =   600
               Width           =   1080
            End
         End
         Begin VB.CommandButton cmdTP_browse 
            Caption         =   "&Browse"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3345
            TabIndex        =   92
            Top             =   1935
            Width           =   1230
         End
      End
      Begin VB.OptionButton optPIC_TEXT 
         Caption         =   "&Picture TEXT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Render TEXT as a picture cutout."
         Top             =   0
         Width           =   1875
      End
      Begin VB.OptionButton optCLR_TEXT 
         Caption         =   "&Coloured TEXT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   795
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Render TEXT using predined colours and patterns."
         Top             =   0
         Value           =   -1  'True
         Width           =   1875
      End
   End
   Begin VB.PictureBox picTAB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2925
      Index           =   4
      Left            =   330
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   356
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   5340
      Begin VB.PictureBox picBACK_CLR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   75
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   346
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   525
         Width           =   5190
         Begin VB.PictureBox picBD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   4
            Left            =   3435
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1770
            Width           =   780
         End
         Begin VB.PictureBox picBD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   3
            Left            =   3435
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1035
            Width           =   780
         End
         Begin VB.PictureBox picBD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   2
            Left            =   2505
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   1770
            Width           =   780
         End
         Begin VB.PictureBox picBD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   1
            Left            =   2505
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   1035
            Width           =   780
         End
         Begin VB.PictureBox picBC_Colour 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   3930
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   81
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   315
            Width           =   1215
         End
         Begin VB.PictureBox picBC_Colour 
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   2505
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   81
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   315
            Width           =   1215
         End
         Begin VB.Frame fraBC_STYLE 
            Caption         =   "Back Gradient Style"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2400
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   2250
            Begin VB.OptionButton optBCS 
               Caption         =   "&Square"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   63
               Top             =   1965
               Width           =   1875
            End
            Begin VB.OptionButton optBCS 
               Caption         =   "Diagonal &Down"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   62
               Top             =   1590
               Width           =   1875
            End
            Begin VB.OptionButton optBCS 
               Caption         =   "Diagonal &Up"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   61
               Top             =   1290
               Width           =   1875
            End
            Begin VB.OptionButton optBCS 
               Caption         =   "&Vertical"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   60
               Top             =   915
               Width           =   1875
            End
            Begin VB.OptionButton optBCS 
               Caption         =   "&Horizontal"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   59
               Top             =   615
               Width           =   1875
            End
            Begin VB.OptionButton optBCS 
               Caption         =   "&One Colour"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   1875
            End
         End
         Begin VB.PictureBox picBD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   5
            Left            =   4365
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1035
            Width           =   780
         End
         Begin VB.PictureBox picBD 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   6
            Left            =   4365
            ScaleHeight     =   39
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   52
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   1770
            Width           =   780
         End
         Begin VB.Label lblBDrctn 
            Alignment       =   2  'Center
            Caption         =   "Back Gradient Direction"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2505
            TabIndex        =   68
            Top             =   720
            Width           =   2640
         End
         Begin VB.Label lblBC 
            Alignment       =   2  'Center
            Caption         =   "Back Colour 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   3915
            TabIndex        =   66
            Top             =   0
            Width           =   1245
         End
         Begin VB.Label lblBC 
            Alignment       =   2  'Center
            Caption         =   "Back Colour 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   2490
            TabIndex        =   64
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.PictureBox picBACK_PIC 
         BorderStyle     =   0  'None
         Height          =   2400
         Left            =   75
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   346
         TabIndex        =   75
         Top             =   525
         Visible         =   0   'False
         Width           =   5190
         Begin VB.CommandButton cmdBP_Browse 
            Caption         =   "&Browse"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3345
            TabIndex        =   86
            Top             =   1935
            Width           =   1230
         End
         Begin VB.Frame fraBP_Fit 
            Caption         =   "Picture Fit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   0
            TabIndex        =   80
            Top             =   1395
            Width           =   2430
            Begin VB.OptionButton optBP_PicFit 
               Caption         =   "&Zoom"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   1290
               TabIndex        =   84
               Top             =   600
               Width           =   1080
            End
            Begin VB.OptionButton optBP_PicFit 
               Caption         =   "&Stretch"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   83
               Top             =   600
               Width           =   1080
            End
            Begin VB.OptionButton optBP_PicFit 
               Caption         =   "&Tile"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1290
               TabIndex        =   82
               Top             =   270
               Width           =   1080
            End
            Begin VB.OptionButton optBP_PicFit 
               Caption         =   "Ce&ntre"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   81
               Top             =   270
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin VB.Frame fraBP_Ptype 
            Caption         =   "Picture Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   2430
            Begin VB.OptionButton optBP_PicType 
               Caption         =   "&User Specified"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   79
               Top             =   930
               Width           =   2175
            End
            Begin VB.OptionButton optBP_PicType 
               Caption         =   "Use &Background"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   77
               Top             =   270
               Width           =   2175
            End
            Begin VB.OptionButton optBP_PicType 
               Caption         =   "Use &Desktop"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   78
               Top             =   600
               Width           =   2175
            End
         End
         Begin VB.PictureBox picBP_Preview 
            AutoRedraw      =   -1  'True
            Height          =   1860
            Left            =   2730
            ScaleHeight     =   120
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   160
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   0
            Width           =   2460
         End
      End
      Begin VB.OptionButton optCLR_BACK 
         Caption         =   "&Coloured BACK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   795
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Render TEXT using predined colours and patterns."
         Top             =   0
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optPIC_BACK 
         Caption         =   "&Picture BACK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Render TEXT as a picture cutout."
         Top             =   0
         Width           =   1875
      End
   End
   Begin VB.PictureBox picTAB 
      BorderStyle     =   0  'None
      Height          =   2925
      Index           =   5
      Left            =   330
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   356
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   5340
      Begin VB.Label lblABOUT 
         Alignment       =   2  'Center
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         Left            =   225
         TabIndex        =   88
         Top             =   225
         Width           =   4890
      End
   End
   Begin VB.CommandButton cmdDEFAULT 
      Caption         =   "Use Defaults"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   89
      Top             =   4425
      Width           =   1650
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4575
      TabIndex        =   91
      Top             =   4425
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3075
      TabIndex        =   90
      Top             =   4425
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cdlSETTINGS 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   255
   End
   Begin MSComctlLib.TabStrip tabSETTINGS 
      Height          =   3675
      Left            =   150
      TabIndex        =   2
      Top             =   675
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   6482
      TabFixedWidth   =   1958
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Object.ToolTipText     =   "Alter general settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            Object.ToolTipText     =   "Alter Font Settings."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text Fill"
            Object.ToolTipText     =   "Alter Text Drawing settings."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Background Fill"
            Object.ToolTipText     =   "Alter settings for the Background."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Object.ToolTipText     =   "About this product."
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picHDR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   0
      Picture         =   "Settings.frx":1A82
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.PictureBox picWORK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5400
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim CrrntTabIndex As Integer

    Dim BttnFrame As RECT
    Dim FocusRect As RECT
    
    Dim FontsLoaded As Boolean

'---------------'
' Load the Form '
'---------------'
Private Sub Form_Load()

    'Get all the settings from the registry.
    Call GetValues
    
    txtDISPLAY.TEXT = CM.SavrText
    txtDISPLAY.FontName = CM.FontName
    If CM.BelowPic Then optBelow.Value = True
    If CM.AniSpeed >= 0 And CM.AniSpeed <= 100 Then sldSPEED.Value = CM.AniSpeed
    
    txtFONTS.TEXT = CM.FontName
    Call Load_Up_Fonts
    
    If CM.FontBold Then chkBOLD.Value = vbChecked
    If CM.FontItlc Then chkITALIC.Value = vbChecked
    If CM.FontUndr Then chkUNDERLINE.Value = vbChecked
    
    'Set up a focus rectangle for the colour directions
    FocusRect.LEFT = 3: FocusRect.TOP = 3: FocusRect.Right = 49: FocusRect.Bottom = 36
    
    optTCS(CM.TEXT.ColourStyle).Value = True
    picTC_Colour(1).BackColor = CM.TEXT.GradColour(1)
    picTC_Colour(2).BackColor = CM.TEXT.GradColour(2)
    Call optTCS_Click(CM.TEXT.ColourStyle)
    'Draw a frame around the First Text Colour.
    BttnFrame.LEFT = picTC_Colour(1).LEFT - 3: BttnFrame.Right = BttnFrame.LEFT + 87
    BttnFrame.TOP = picTC_Colour(1).TOP - 3:   BttnFrame.Bottom = BttnFrame.TOP + 25
    DrawFrameControl picTEXT_CLR.hDC, BttnFrame, 0, 0
    'Draw a frame around the Second Text Colour.
    BttnFrame.LEFT = picTC_Colour(2).LEFT - 3: BttnFrame.Right = BttnFrame.LEFT + 87
    BttnFrame.TOP = picTC_Colour(2).TOP - 3:   BttnFrame.Bottom = BttnFrame.TOP + 25
    DrawFrameControl picTEXT_CLR.hDC, BttnFrame, 0, 0
    'Draw frames around the Text Gradient Direction controls.
    For PNTR = 1 To 6
        BttnFrame.LEFT = picTD(PNTR).LEFT - 3: BttnFrame.Right = BttnFrame.LEFT + 58
        BttnFrame.TOP = picTD(PNTR).TOP - 3:   BttnFrame.Bottom = BttnFrame.TOP + 45
        DrawFrameControl picTEXT_CLR.hDC, BttnFrame, 0, 0
    Next PNTR
    
    optBCS(CM.BACK.ColourStyle).Value = True
    picBC_Colour(1).BackColor = CM.BACK.GradColour(1)
    picBC_Colour(2).BackColor = CM.BACK.GradColour(2)
    Call optBCS_Click(CM.BACK.ColourStyle)
    'Draw a frame around the First Background Colour.
    BttnFrame.LEFT = picBC_Colour(1).LEFT - 3: BttnFrame.Right = BttnFrame.LEFT + 87
    BttnFrame.TOP = picBC_Colour(1).TOP - 3:   BttnFrame.Bottom = BttnFrame.TOP + 25
    DrawFrameControl picBACK_CLR.hDC, BttnFrame, 0, 0
    'Draw a frame around the Second Background Colour.
    BttnFrame.LEFT = picBC_Colour(2).LEFT - 3: BttnFrame.Right = BttnFrame.LEFT + 87
    BttnFrame.TOP = picBC_Colour(2).TOP - 3:   BttnFrame.Bottom = BttnFrame.TOP + 25
    DrawFrameControl picBACK_CLR.hDC, BttnFrame, 0, 0
    'Draw frames around the Background Gradient Direction controls.
    For PNTR = 1 To 6
        BttnFrame.LEFT = picBD(PNTR).LEFT - 3: BttnFrame.Right = BttnFrame.LEFT + 58
        BttnFrame.TOP = picBD(PNTR).TOP - 3:   BttnFrame.Bottom = BttnFrame.TOP + 45
        DrawFrameControl picBACK_CLR.hDC, BttnFrame, 0, 0
    Next PNTR
    
    'Put information about this product in the "About" label
    lblABOUT.Caption = App.ProductName & vbCrLf & vbCrLf & _
                       App.Comments & vbCrLf & vbCrLf & _
                       "Version " & Str(App.Major) & " · " & _
                       Format(App.Minor, "00") & " · " & _
                       Format(App.Revision, "0000") & _
                       vbCrLf & vbCrLf & App.LegalCopyright
    
    'Set the first Tab as active
    CrrntTabIndex = tabSETTINGS.SelectedItem.Index

End Sub

'----------------------------------------'
' Display the Picture Title on the Form. '
'----------------------------------------'
Private Sub Form_Activate()

    'Make a cutout in the Form to accept the Picture Title.
    BitBlt Settings.hDC, (Settings.ScaleWidth - picHDR.WIDTH) / 2, 0, 288, 44, picHDR.hDC, 0, 0, vbSrcAnd
    'Slot the Picture Title into the cutout made in the Form.
    BitBlt Settings.hDC, (Settings.ScaleWidth - picHDR.WIDTH) / 2, 0, 288, 44, picHDR.hDC, 0, 44, vbSrcPaint

End Sub

'-----------------------------------'
' One of the Tabs has been Clicked. '
'-----------------------------------'
Private Sub tabSETTINGS_Click()

    'Only take action if this is a different Tab
    'to that which is currently selected.
    If tabSETTINGS.SelectedItem.Index <> CrrntTabIndex Then
        picTAB(tabSETTINGS.SelectedItem.Index).Visible = True
        picTAB(CrrntTabIndex).Visible = False
        CrrntTabIndex = tabSETTINGS.SelectedItem.Index
        Select Case CrrntTabIndex
            Case 2: lsvFONTS.SetFocus
            Case 3: optPIC_TEXT.Value = CM.TEXT.PictureFill
            Case 4: optPIC_BACK.Value = CM.BACK.PictureFill
        End Select
    End If

End Sub

'-----------------------------------------------'
' Load up all available Fonts into a Font List. '
'-----------------------------------------------'
Private Sub Load_Up_Fonts()

    Dim PNTR As Integer
    Dim Crnt As Integer
    'Dim itmX As ListItem, x
    Dim lstFONTShDC As Long
    
    FontsLoaded = False
    
    lstFONTShDC = GetDC(lstFONTS.hWnd)
    
    'Put all the TrueType's into lstFONTS
    lstFONTS.Clear
    ShowFontType = 4 'True Type
    EnumFontFamilies lstFONTShDC, vbNullString, AddressOf EnumFontFamTypeProc, lstFONTS
    'Add the TrueTypes font names with the TrueType icon to the Listview
    For PNTR = 0 To lstFONTS.ListCount - 1
        lsvFONTS.ListItems.Add , , lstFONTS.List(PNTR), , 1
    Next PNTR
    
    'Put the Fixed fonts into lstFONTS
    lstFONTS.Clear
    ShowFontType = 1 'Fixed Width
    EnumFontFamilies lstFONTShDC, vbNullString, AddressOf EnumFontFamTypeProc, lstFONTS
    'Add the FixedWidth font names with the FixedWidth icon to the Listview
    For PNTR = 0 To lstFONTS.ListCount - 1
        lsvFONTS.ListItems.Add , , lstFONTS.List(PNTR), , 2
    Next PNTR
    
    ReleaseDC lstFONTS.hWnd, lstFONTShDC
    
    FontsLoaded = True
    
    Call DisplayFontMatch

End Sub

'-------------------------------------------'
' Ensure that the display Text isn't blank. '
'-------------------------------------------'
Private Sub txtDISPLAY_LostFocus()

    If txtDISPLAY.TEXT = "" Then txtDISPLAY.TEXT = "Charlie II Screen Saver"
    CM.SavrText = txtDISPLAY.TEXT

End Sub

'----------------------------------------------'
' Text Above Picture option has been selected. '
'----------------------------------------------'
Private Sub optAbove_Click()
    CM.BelowPic = False
End Sub

'----------------------------------------------'
' Text Below Picture option has been selected. '
'----------------------------------------------'
Private Sub optBelow_Click()
    CM.BelowPic = True
End Sub

'--------------------------------------------------'
' Save the new Animation Speed into the Settings   '
' Area when the value of the Speed slider changes. '
'--------------------------------------------------'
Private Sub sldSPEED_Change()
    CM.AniSpeed = sldSPEED.Value
End Sub

'--------------------------------------'
' When the Font Name gets the focus we '
' select all of the Text.              '
'--------------------------------------'
Private Sub txtFONTS_GotFocus()

    txtFONTS.SelStart = 0
    txtFONTS.SelLength = Len(txtFONTS.TEXT)

End Sub

'----------------------------------------------'
' When the typed in Font Name is altered we'll '
' find the closest match to the current name.  '
'----------------------------------------------'
Private Sub txtFONTS_Change()

    Call DisplayFontMatch

End Sub

'-------------------------------------'
' A different Font has been selected. '
'-------------------------------------'
Private Sub lsvFonts_ItemClick(ByVal Item As MSComctlLib.ListItem)

    'Font name display is updated.
    txtFONTS.TEXT = lsvFONTS.ListItems(Item.Index)
    'Text Display's Font Name attribute is set to the selected Font.
    txtDISPLAY.FontName = txtFONTS.TEXT
    txtDISPLAY.FontSize = 12
    'The new Font Name is moved into the Settings area.
    CM.FontName = txtFONTS.TEXT

End Sub

'--------------------------------------------------'
' Find the closest match of the typed in Font Name '
' to a Font found in the list of loaded Fonts.     '
'--------------------------------------------------'
Private Sub DisplayFontMatch()
    
    'Do not do this if the Fonts haven't been loaded yet.
    If Not FontsLoaded Then Exit Sub
    
    'Find and display the Font Name found in the txtFONTS text box.
    Dim itmFound As ListItem
    Set itmFound = lsvFONTS.FindItem(txtFONTS.TEXT, lvwText, , lvwPartial)
    
    'If no match, select and scroll to the first item in the list.
    If itmFound Is Nothing Then
        lsvFONTS.ListItems(1).Selected = True
        lsvFONTS.ListItems(1).EnsureVisible
        CM.FontName = lsvFONTS.ListItems(1).TEXT
    Else
    'Scroll to the found item.
        itmFound.EnsureVisible
        itmFound.Selected = True
    End If

End Sub

'-------------------------------------'
' The "Bold" option has been altered. '
'-------------------------------------'
Private Sub chkBOLD_Click()

    If chkBOLD.Value = vbChecked Then
        txtDISPLAY.Font.Bold = True
    Else
        txtDISPLAY.Font.Bold = False
    End If
    CM.FontBold = txtDISPLAY.Font.Bold

End Sub

'---------------------------------------'
' The "Italic" option has been altered. '
'---------------------------------------'
Private Sub chkITALIC_Click()

    If chkITALIC.Value = vbChecked Then
        txtDISPLAY.Font.Italic = True
    Else
        txtDISPLAY.Font.Italic = False
    End If
    CM.FontItlc = txtDISPLAY.Font.Italic

End Sub

'------------------------------------------'
' The "Underline" option has been altered. '
'------------------------------------------'
Private Sub chkUNDERLINE_Click()

    If chkUNDERLINE.Value = vbChecked Then
        txtDISPLAY.Font.Underline = True
    Else
        txtDISPLAY.Font.Underline = False
    End If
    CM.FontUndr = txtDISPLAY.Font.Underline

End Sub

'----------------------------------------------'
' The Text "Colours" option has been selected. '
'----------------------------------------------'
Private Sub optCLR_TEXT_Click()

    If optCLR_TEXT.Value Then
        picTEXT_PIC.Visible = False
        picTEXT_CLR.Visible = True
        CM.TEXT.PictureFill = False
    End If

End Sub

'--------------------------------------------------'
' The Text Colour "Style" option has been altered. '
'--------------------------------------------------'
Private Sub optTCS_Click(Index As Integer)

    CM.TEXT.ColourStyle = Index
    If Index = 0 Then
        'When "One Colour" is selected we'll hide the second
        'Colour control and the Gradient Direction images.
        picTC_Colour(2).Visible = False
        For PNTR = 1 To 6: picTD(PNTR).Visible = False: Next PNTR
    Else
        'If other than "One Colour" has been selected we'll
        'make all of the Gradient controls visible.
        picTC_Colour(2).Visible = True
        For PNTR = 1 To 6: picTD(PNTR).Visible = True: Next PNTR
        Call Redraw_Text_Directions
    End If

End Sub

'-------------------------------------'
' Invoke the Common Dialog control to '
' browse for Text Gradient Colours.   '
'-------------------------------------'
Private Sub picTC_Colour_Click(Index As Integer)

    'Set Cancel to True
    cdlSETTINGS.CancelError = True
    On Error GoTo picTC_Colour_Click_Error
    'Set the Flags property
    cdlSETTINGS.Flags = cdlCCRGBInit Or cdlCCFullOpen Or cdlCCRGBInit
    cdlSETTINGS.Color = picTC_Colour(Index).BackColor
    'Display the Colour Dialog box
    cdlSETTINGS.ShowColor
    
    'Use the returned Colour
    picTC_Colour(Index).BackColor = cdlSETTINGS.Color
    CM.TEXT.GradColour(Index) = cdlSETTINGS.Color
    Call Redraw_Text_Directions
    
picTC_Colour_Click_Error:
    'User pressed the Cancel button
    Exit Sub
    
End Sub

'---------------------------------------------------'
' A different Gradient Direction has been selected. '
'---------------------------------------------------'
Private Sub picTD_Click(Index As Integer)
    CM.TEXT.ColourDrctn = Index
    Call Redraw_Text_Directions
End Sub

'--------------------------------------------'
' The Text Gradient Directions are re-drawn. '
'--------------------------------------------'
Private Sub Redraw_Text_Directions()

    For PNTR = 1 To 6
        Call GradFill(picTD(PNTR), CM.TEXT.GradColour(1), CM.TEXT.GradColour(2), _
                      CM.TEXT.ColourStyle, PNTR)
    Next PNTR
    
    'A Focus Rectangle is drawn onto the selected Gradient Direction.
    DrawFocusRect picTD(CM.TEXT.ColourDrctn).hDC, FocusRect

End Sub

'----------------------------------------------'
' The Text "Picture" option has been selected. '
'----------------------------------------------'
Private Sub optPIC_TEXT_Click()

    If optPIC_TEXT.Value Then
        picTEXT_CLR.Visible = False
        picTEXT_PIC.Visible = True
        CM.TEXT.PictureFill = True
        optTP_PicType(CM.TEXT.PictureType) = True
    End If

End Sub

'--------------------------------------------'
' The Type of Text Picture has been altered. '
'--------------------------------------------'
Private Sub optTP_PicType_Click(Index As Integer)

    CM.TEXT.PictureType = Index
    Select Case Index
        Case 0 To 1
            'The "Picture Fit" and "Browse" options are not available
            'for non-user defined pictures and are therefore disabled.
            fraTP_Fit.Enabled = False
            optTP_PicFit(0).Enabled = False
            optTP_PicFit(1).Enabled = False
            optTP_PicFit(2).Enabled = False
            optTP_PicFit(3).Enabled = False
            cmdTP_browse.Enabled = False
            Call LoadTextPic
        Case 2
            'The "Picture Fit" and "Browse" options
            'are enabled for user defined pictures.
            fraTP_Fit.Enabled = True
            optTP_PicFit(0).Enabled = True
            optTP_PicFit(1).Enabled = True
            optTP_PicFit(2).Enabled = True
            optTP_PicFit(3).Enabled = True
            cmdTP_browse.Enabled = True
            If optTP_PicFit(CM.TEXT.PictureFit) = True Then
                Call LoadTextPic
            Else
                optTP_PicFit(CM.TEXT.PictureFit) = True
            End If
    End Select

End Sub

'--------------------------------------------'
' The Text Picture's "Fit" has been altered. '
'--------------------------------------------'
Private Sub optTP_PicFit_Click(Index As Integer)

    CM.TEXT.PictureFit = Index
    Call LoadTextPic

End Sub

'-------------------------------------'
' Invoke the Common Dialog control to '
' browse for a Text Picture.          '
'-------------------------------------'
Private Sub cmdTP_browse_Click()

    'Set Cancel to True
    cdlSETTINGS.CancelError = True
    On Error GoTo cmdTP_browse_Click_Error
    'Setup the flags
    cdlSETTINGS.Flags = cdlOFNFileMustExist Or _
                        cdlOFNHideReadOnly Or _
                        cdlOFNLongNames Or _
                        cdlOFNNoChangeDir
    'Set the filter to open only Bitmaps or Jpegs.
    cdlSETTINGS.Filter = "Picture Files (*.bmp;*.jpg)|*.bmp;*.jpg"
    'Set up the existing File Name
    cdlSETTINGS.filename = CM.TEXT.PictureAddr
    'Display the OPEN dialog box.
    cdlSETTINGS.ShowOpen
    
    'Apply the returned Filename.
    CM.TEXT.PictureAddr = cdlSETTINGS.filename
    Call LoadTextPic
    
cmdTP_browse_Click_Error:
    'User pressed the Cancel button
    Exit Sub

End Sub

'--------------------------------'
' Load the Text Picture into the '
' Text Picture Picture Box.      '
'--------------------------------'
Private Sub LoadTextPic()

    Dim SaveScan_PicText As Long
    
    'If the Picture is Stretched in this Sub then save it's
    'Stretch Mode before altering it to "HalfTone".
    If CM.TEXT.PictureType < 2 Then
        SaveScan_PicText = GetStretchBltMode(picTP_Preview.hDC)
        RtnVal = SetStretchBltMode(picTP_Preview.hDC, STRETCH_HALFTONE)
    End If
    
    Select Case CM.TEXT.PictureType
        Case 0
            'Set the work Picture's size to the Screen size
            picWORK.WIDTH = Screen.WIDTH / Screen.TwipsPerPixelX
            picWORK.HEIGHT = Screen.HEIGHT / Screen.TwipsPerPixelY
            'Paint the background into the Work picture
            RtnVal = PaintDesktop(picWORK.hDC)
            'Stretch this image into the Preveiw Picture Box
            RtnVal = StretchBlt(picTP_Preview.hDC, 0, 0, 160, 120, _
                                picWORK.hDC, 0, 0, picWORK.WIDTH, picWORK.HEIGHT, _
                                vbSrcCopy)
        Case 1
            'Get the Handle for the desktop
            DeskTop_hWnd = GetDesktopWindow
            'Get the Device Context for the textop from it's Handle
            DeskTop_DC = GetDC(DeskTop_hWnd)
            'Stretch an image of the desktop into the Preview Picture Box
            RtnVal = StretchBlt(picTP_Preview.hDC, 0, 0, 160, 120, _
                                DeskTop_DC, 0, 0, _
                                Screen.WIDTH / Screen.TwipsPerPixelX, _
                                Screen.HEIGHT / Screen.TwipsPerPixelY, vbSrcCopy)
        Case 2
            Call PictFill(picTP_Preview, picWORK, CM.TEXT.PictureAddr, CM.TEXT.PictureFit)
    End Select
    picTP_Preview.Refresh
    
    'If the Picture's Stretch Mode has been altered then
    'restore it back to it's original configuration.
    If CM.TEXT.PictureType < 2 Then
        RtnVal = SetStretchBltMode(picTP_Preview.hDC, SaveScan_PicText)
    End If

End Sub

'----------------------------------------------------'
' The Background "Colours" option has been selected. '
'----------------------------------------------------'
Private Sub optCLR_BACK_Click()

    If optCLR_BACK.Value Then
        picBACK_PIC.Visible = False
        picBACK_CLR.Visible = True
        CM.BACK.PictureFill = False
    End If

End Sub

'--------------------------------------------------------'
' The Background Colour "Style" option has been altered. '
'--------------------------------------------------------'
Private Sub optBCS_Click(Index As Integer)

    CM.BACK.ColourStyle = Index
    If Index = 0 Then
        'When "One Colour" is selected we'll hide the second
        'Colour control and the Gradient Direction images.
        picBC_Colour(2).Visible = False
        For PNTR = 1 To 6: picBD(PNTR).Visible = False: Next PNTR
    Else
        'If other than "One Colour" has been selected we'll
        'make all of the Gradient controls visible.
        picBC_Colour(2).Visible = True
        For PNTR = 1 To 6: picBD(PNTR).Visible = True: Next PNTR
        Call Redraw_Back_Directions
    End If

End Sub

'-----------------------------------------'
' Invoke the Common Dialog control to     '
' browse for Background Gradient Colours. '
'-----------------------------------------'
Private Sub picBC_Colour_Click(Index As Integer)

    'Set Cancel to True
    cdlSETTINGS.CancelError = True
    On Error GoTo picBC_Colour_Click_Error
    'Set the Flags property
    cdlSETTINGS.Flags = cdlCCRGBInit Or cdlCCFullOpen Or cdlCCRGBInit
    cdlSETTINGS.Color = picBC_Colour(Index).BackColor
    'Display the Colour Dialog box
    cdlSETTINGS.ShowColor
    
    'Use the returned Colour
    picBC_Colour(Index).BackColor = cdlSETTINGS.Color
    CM.BACK.GradColour(Index) = cdlSETTINGS.Color
    Call Redraw_Back_Directions
    
picBC_Colour_Click_Error:
    'User pressed the Cancel button
    Exit Sub
    
End Sub

'---------------------------------------------------'
' A different Gradient Direction has been selected. '
'---------------------------------------------------'
Private Sub picBD_Click(Index As Integer)
    CM.BACK.ColourDrctn = Index
    Call Redraw_Back_Directions
End Sub

'--------------------------------------------------'
' The Background Gradient Directions are re-drawn. '
'--------------------------------------------------'
Private Sub Redraw_Back_Directions()

    For PNTR = 1 To 6
        Call GradFill(picBD(PNTR), CM.BACK.GradColour(1), CM.BACK.GradColour(2), _
                      CM.BACK.ColourStyle, PNTR)
    Next PNTR
    
    'A Focus Rectangle is drawn onto the selected Gradient Direction.
    DrawFocusRect picBD(CM.BACK.ColourDrctn).hDC, FocusRect

End Sub

'----------------------------------------------------'
' The Background "Picture" option has been selected. '
'----------------------------------------------------'
Private Sub optPIC_BACK_Click()

    If optPIC_BACK.Value Then
        picBACK_CLR.Visible = False
        picBACK_PIC.Visible = True
        CM.BACK.PictureFill = True
        optBP_PicType(CM.BACK.PictureType) = True
    End If

End Sub

'--------------------------------------------------'
' The Type of Background Picture has been altered. '
'--------------------------------------------------'
Private Sub optBP_PicType_Click(Index As Integer)

    CM.BACK.PictureType = Index
    Select Case Index
        Case 0 To 1
            'The "Picture Fit" and "Browse" options are not available
            'for non-user defined pictures and are therefore disabled.
            fraBP_Fit.Enabled = False
            optBP_PicFit(0).Enabled = False
            optBP_PicFit(1).Enabled = False
            optBP_PicFit(2).Enabled = False
            optBP_PicFit(3).Enabled = False
            cmdBP_Browse.Enabled = False
            Call LoadBackPic
        Case 2
            'The "Picture Fit" and "Browse" options
            'are enabled for user defined pictures.
            fraBP_Fit.Enabled = True
            optBP_PicFit(0).Enabled = True
            optBP_PicFit(1).Enabled = True
            optBP_PicFit(2).Enabled = True
            optBP_PicFit(3).Enabled = True
            cmdBP_Browse.Enabled = True
            If optBP_PicFit(CM.BACK.PictureFit) = True Then
                Call LoadBackPic
            Else
                optBP_PicFit(CM.BACK.PictureFit) = True
            End If
    End Select

End Sub

'--------------------------------------------------'
' The Background Picture's "Fit" has been altered. '
'--------------------------------------------------'
Private Sub optBP_PicFit_Click(Index As Integer)

    CM.BACK.PictureFit = Index
    Call LoadBackPic

End Sub

'-------------------------------------'
' Invoke the Common Dialog control to '
' browse for a Background Picture.    '
'-------------------------------------'
Private Sub cmdBP_Browse_Click()

    'Set Cancel to True
    cdlSETTINGS.CancelError = True
    On Error GoTo cmdBP_Browse_Click_Error
    'Setup the flags
    cdlSETTINGS.Flags = cdlOFNFileMustExist Or _
                        cdlOFNHideReadOnly Or _
                        cdlOFNLongNames Or _
                        cdlOFNNoChangeDir
    'Set the filter to open only Bitmaps or Jpegs.
    cdlSETTINGS.Filter = "Picture Files (*.bmp;*.jpg)|*.bmp;*.jpg"
    'Set up the existing File Name
    cdlSETTINGS.filename = CM.BACK.PictureAddr
    'Display the OPEN dialog box.
    cdlSETTINGS.ShowOpen
    
    'Apply the returned Filename.
    CM.BACK.PictureAddr = cdlSETTINGS.filename
    Call LoadBackPic
    
cmdBP_Browse_Click_Error:
    'User pressed the Cancel button
    Exit Sub

End Sub

'--------------------------------------'
' Load the Background Picture into the '
' Background Picture Picture Box.      '
'--------------------------------------'
Private Sub LoadBackPic()

    Dim SaveScan_PicBack As Long
    
    'If the Picture is Stretched in this Sub then save it's
    'Stretch Mode before altering it to "HalfTone".
    If CM.BACK.PictureType < 2 Then
        SaveScan_PicBack = GetStretchBltMode(picBP_Preview.hDC)
        RtnVal = SetStretchBltMode(picBP_Preview.hDC, STRETCH_HALFTONE)
    End If
    
    Select Case CM.BACK.PictureType
        Case 0
            'Set the work Picture's size to the Screen size
            picWORK.WIDTH = Screen.WIDTH / Screen.TwipsPerPixelX
            picWORK.HEIGHT = Screen.HEIGHT / Screen.TwipsPerPixelY
            'Paint the background into the Work picture
            RtnVal = PaintDesktop(picWORK.hDC)
            'Stretch this image into the Preveiw Picture Box
            RtnVal = StretchBlt(picBP_Preview.hDC, 0, 0, 160, 120, _
                                picWORK.hDC, 0, 0, picWORK.WIDTH, picWORK.HEIGHT, _
                                vbSrcCopy)
        Case 1
            'Get the Handle for the desktop
            DeskTop_hWnd = GetDesktopWindow
            'Get the Device Context for the Desktop from it's Handle
            DeskTop_DC = GetDC(DeskTop_hWnd)
            'Stretch an image of the desktop into the Preview Picture Box
            RtnVal = StretchBlt(picBP_Preview.hDC, 0, 0, 160, 120, _
                                DeskTop_DC, 0, 0, _
                                Screen.WIDTH / Screen.TwipsPerPixelX, _
                                Screen.HEIGHT / Screen.TwipsPerPixelY, vbSrcCopy)
        Case 2
            Call PictFill(picBP_Preview, picWORK, CM.BACK.PictureAddr, CM.BACK.PictureFit)
    End Select
    picBP_Preview.Refresh
    
    'If the Picture's Stretch Mode has been altered then
    'restore it back to it's original configuration.
    If CM.BACK.PictureType < 2 Then
        RtnVal = SetStretchBltMode(picBP_Preview.hDC, SaveScan_PicBack)
    End If

End Sub

'--------------------------------------------'
' DEFAULT button was pressed so we'll set    '
' all settings back to their default values  '
' then save them before exiting the program. '
'--------------------------------------------'
Private Sub cmdDEFAULT_Click()

    CM.SavrText = "Charlie II Screen Saver"
    CM.BelowPic = True
    CM.AniSpeed = 50
    
    CM.FontName = "Times New Roman"
    CM.FontBold = True
    CM.FontItlc = False
    CM.FontUndr = False
    
    CM.TEXT.PictureFill = False
    CM.TEXT.FillStyle = 0
    CM.TEXT.ColourStyle = 1
    CM.TEXT.GradColour(1) = vbYellow
    CM.TEXT.GradColour(2) = vbRed
    CM.TEXT.ColourDrctn = 1
    CM.TEXT.PictureType = 0
    CM.TEXT.PictureFit = 0
    CM.TEXT.PictureAddr = "C:\Windows\Red Blocks.bmp"
    
    CM.BACK.PictureFill = False
    CM.BACK.FillStyle = 0
    CM.BACK.ColourStyle = 3
    CM.BACK.GradColour(1) = vbCyan
    CM.BACK.GradColour(2) = vbBlue
    CM.BACK.ColourDrctn = 1
    CM.BACK.PictureType = 0
    CM.BACK.PictureFit = 0
    CM.BACK.PictureAddr = "C:\Windows\Clouds.bmp"
    
    Call PutValues
    Unload Settings

End Sub

'---------------------------------------------------------'
' OK button was pressed - Save all the Settings then Exit '
'---------------------------------------------------------'
Private Sub cmdOK_Click()

    Call PutValues
    Unload Settings

End Sub

'--------------------------------------------------------------'
' CANCEL button was pressed - Exit without saving the Settings '
'--------------------------------------------------------------'
Private Sub cmdCANCEL_Click()

    Unload Settings

End Sub

