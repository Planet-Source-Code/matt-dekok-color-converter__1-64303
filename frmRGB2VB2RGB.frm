VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmRGB2VB2RGB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Converter"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8895
   ForeColor       =   &H80000008&
   Icon            =   "frmRGB2VB2RGB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCursor 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   240
      Picture         =   "frmRGB2VB2RGB.frx":08CA
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   87
      Top             =   4020
      Width           =   225
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Color Dialog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4800
      Picture         =   "frmRGB2VB2RGB.frx":0E40
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "Pick Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4800
      Picture         =   "frmRGB2VB2RGB.frx":1182
      Style           =   1  'Graphical
      TabIndex        =   84
      TabStop         =   0   'False
      ToolTipText     =   "Pick the Background Color from Screen"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Copy Delphi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   80
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Copy Java"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   79
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copy Photoshop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   78
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Copy C++"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   73
      Top             =   3000
      Width           =   2655
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3135
      ScaleHeight     =   195
      ScaleWidth      =   2715
      TabIndex        =   69
      Top             =   1575
      Width           =   2745
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8040
      Top             =   4200
   End
   Begin VB.Timer tmrPick 
      Left            =   7680
      Top             =   4200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   68
      Top             =   4575
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Copy Web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   67
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy VB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   66
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy RGB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   65
      Top             =   1920
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmRGB2VB2RGB.frx":12CC
      Left            =   6120
      List            =   "frmRGB2VB2RGB.frx":131B
      TabIndex        =   64
      Text            =   "System Colors"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Frame Frame6 
      Caption         =   "QB Colors"
      Height          =   1335
      Left            =   6120
      TabIndex        =   47
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton QB 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   720
         TabIndex        =   62
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1320
         TabIndex        =   61
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1920
         TabIndex        =   60
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   720
         TabIndex        =   58
         Top             =   480
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   1320
         TabIndex        =   57
         Top             =   480
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   1920
         TabIndex        =   56
         Top             =   480
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   720
         TabIndex        =   54
         Top             =   720
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   1320
         TabIndex        =   53
         Top             =   720
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   1920
         TabIndex        =   52
         Top             =   720
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   720
         TabIndex        =   50
         Top             =   960
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   1320
         TabIndex        =   49
         Top             =   960
         Width           =   600
      End
      Begin VB.CommandButton QB 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   1920
         TabIndex        =   48
         Top             =   960
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2520
      TabIndex        =   8
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton cmdConvert2 
         BackColor       =   &H80000003&
         Caption         =   "Calculate Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaskColor       =   &H8000000F&
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   840
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         Caption         =   "VB or OLE Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1935
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "The format for this textbox is ""&H00FFFFFF&"""
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox opB 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox opG 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox opR 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Web Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   2055
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1815
         End
      End
   End
   Begin VB.TextBox txtVBColor 
      Height          =   285
      Left            =   600
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame7 
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   5895
      Begin VB.CommandButton Command14 
         Caption         =   "Grey Scale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   89
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Random Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   86
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Web Safe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   83
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "000000"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "0x0"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "$00000000"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "0x00000000"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Invert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   480
         Max             =   255
         TabIndex        =   46
         Top             =   1320
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   480
         Max             =   255
         TabIndex        =   45
         Top             =   960
         Width           =   1935
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   480
         Max             =   255
         TabIndex        =   44
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2520
         TabIndex        =   33
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2520
         TabIndex        =   32
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2520
         TabIndex        =   31
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         TabIndex        =   30
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2520
         TabIndex        =   29
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2520
         TabIndex        =   28
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Display in Hexidecimal?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "#000000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "&H00000000&"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Color Spy by Chetan Sarva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   88
         Top             =   2700
         Width           =   2295
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Photoshop:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   77
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Java:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   76
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Delphi:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   74
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "C++:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   71
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "0"
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "0"
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "160"
         Height          =   255
         Left            =   3480
         TabIndex        =   39
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Lum:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Sat:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Hue:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Web:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "VB:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "RGB:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   195
      Left            =   2640
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H80000003&
      Caption         =   "&Calculate Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtB 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtG 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtR 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set R, G, and B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label Label3 
         Caption         =   "B:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "G:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "R:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Timer timSpy 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7320
      Top             =   4200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuMixer 
      Caption         =   "Color Mixer"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmRGB2VB2RGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As String, i As Integer, varRGB As Long
Dim nRet As Long
' Starting Here: Code by Dan Redding - Blue Knot Software
' I take absolutely no credit for it
' I just wanted a screen color picker and his was the
' easiest to implement.
Private Type PointAPI
    X As Long
    Y As Long
End Type
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const WM_NCACTIVATE = &H86
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

' Made by dreamvb in HIS Project - DM Color Picker Version 3
Private Sub Command11_Click()
    Text2.Text = WebSafe(Text2.Text)
    Text3.Text = WebSafe(Text3.Text)
    Text4.Text = WebSafe(Text4.Text)
End Sub

Private Sub cmdPick_Click()
'start a single point sampling
Dim lReturn As Long
    lReturn = SetCapture(picColor.hwnd)
    tmrPick.Interval = 50
End Sub

Private Sub Command14_Click()
Dim r As Long, g As Long, b As Long
r = Val(Text2.Text)
g = Val(Text3.Text)
b = Val(Text4.Text)
Text2.Text = Int((r + g + b) / 3)
Text3.Text = Int((r + g + b) / 3)
Text4.Text = Int((r + g + b) / 3)
End Sub

Private Sub picCursor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.MousePointer = 99 'set the form to allow other mouse cursors
    Me.MouseIcon = picCursor.Picture 'change mouse cursor to hold whats in the picture box
    picCursor.Visible = False 'set the visible property of the picturebox to false
    frmZoom.Show ' load the zooming window
    Call SendMessage(Me.hwnd, WM_NCACTIVATE, 1, ByVal 0&)
    timSpy.Enabled = True 'enable the timer to spy for code

End Sub

Private Sub picCursor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.MousePointer = 0 'set cursor to default
    picCursor.Visible = True 'show picture box
    Unload frmZoom ' unload the zooming window
    timSpy.Enabled = False 'disable timer to spy

End Sub

Private Sub tmrPick_Timer()
'This routine adapted from a project by Matt Hart

'Matt's comments follow:
' Getpixel sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This sample shows how to get the pixel color of any point
' on the screen. The GetPixel API requires CLIENT coordinates,
' so you must first get the window handle and hDC where the
' cursor is. Once you get that, you can get the pixel.
'
' However, there's one "gotcha" I found while writing this.
' Window titlebars return a "-1" for the pixel color, which
' is invalid! So, what I did to get around that was use
' BitBlt to copy a pixel from that device to the PictureBox
' control I'm using to show the colors, then use the Point
' method to check the color.

'for detailed comments, see corresponding function in tmr5x5
Static lX As Long, lY As Long
On Local Error Resume Next
Dim P As PointAPI, H As Long, hD As Long, r As Long
    GetCursorPos P
    If P.X = lX And P.Y = lY Then Exit Sub
    lX = P.X: lY = P.Y
    H = WindowFromPoint(lX, lY)
    hD = GetDC(H)
    ScreenToClient H, P
    r = GetPixel(hD, P.X, P.Y)
    If r = -1 Then
        BitBlt picColor.hDC, 0, 0, 1, 1, hD, P.X, P.Y, vbSrcCopy
        r = picColor.Point(0, 0)
    Else
        picColor.PSet (0, 0), r
    End If
    ReleaseDC H, hD
    picColor.BackColor = r
    Text2.Text = RGBRed(r)
    Text3.Text = RGBGreen(r)
    Text4.Text = RGBBlue(r)
End Sub

Private Sub picColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lReturn As Long
    'check that we are actually sampling
    If tmrPick.Interval > 0 Then
        tmrPick.Interval = 0
        lReturn = ReleaseCapture
    End If
End Sub
' Ending Here

Private Sub Check1_Click()
Call Text2_Change
Call Text3_Change
Call Text4_Change
Text9.SetFocus
End Sub

Private Sub Command12_Click()
With CommonDialog1
    .Color = Text5.Text
    .ShowColor
    Text2.Text = RGBRed(.Color)
    Text3.Text = RGBGreen(.Color)
    Text4.Text = RGBBlue(.Color)
End With
End Sub

Private Sub Command6_Click()
Text2.Text = 255 - Text2.Text
Text3.Text = 255 - Text3.Text
Text4.Text = 255 - Text4.Text
End Sub

Private Sub cmdConvert_click()
Dim e As Integer
Dim msgE, ss, hexR, hexG, hexB, hR, hG, hB As String
Dim clr As OLE_COLOR
ss = ""
If txtB.Text <> "" And txtG.Text <> "" And txtR.Text <> "" Then
    e = 0
    msgE = ""
    If Not (IsNumeric(txtR.Text)) Then
        msgE = "R is invalid."
        e = e + 1
    Else
        If txtR.Text < 0 Or txtR.Text > 255 Or txtR.Text <> CInt(txtR.Text) Then
            msgE = "R is invalid."
            e = e + 1
        Else
            msgE = ""
        End If
    End If
    If Not (IsNumeric(txtG.Text)) Then
        If e > 0 Then
            msgE = msgE & vbCrLf
        End If
        msgE = msgE & "G is invalid."
        e = e + 1
    ElseIf txtG.Text < 0 Or txtG.Text > 255 Or txtG.Text <> CInt(txtG.Text) Then
        If e > 0 Then
            msgE = msgE & vbCrLf
        End If
        msgE = msgE & "G is invalid."
        e = e + 1
    End If
    If Not (IsNumeric(txtB.Text)) Then
        If e > 0 Then
            msgE = msgE & vbCrLf
        End If
        msgE = msgE & "B is invalid."
        e = e + 1
    ElseIf txtB.Text < 0 Or txtB.Text > 255 Or txtB.Text <> CInt(txtB.Text) Then
        If e > 0 Then
            msgE = msgE & vbCrLf
        End If
        msgE = msgE & "B is invalid."
        e = e + 1
    End If
    If e = 0 And IsNumeric(txtR.Text) And IsNumeric(txtG.Text) And IsNumeric(txtB.Text) Then
        hexR = Hex(txtR.Text)
        hexG = Hex(txtG.Text)
        hexB = Hex(txtB.Text)
        If Len(hexR) = 1 Then
            hexR = "0" & hexR
        End If
        If Len(hexG) = 1 Then
            hexG = "0" & hexG
        End If
        If Len(hexB) = 1 Then
            hexB = "0" & hexB
        End If
        txtVBColor.Text = "&H00" & hexB & hexG & hexR & "&"
        picColor.BackColor = Val(txtVBColor.Text)
        Text5.Text = Val(txtVBColor.Text)
        Text6.Text = txtVBColor.Text
        Text7.Text = "#" & hexR & hexG & hexB
        HScroll1.Value = txtR.Text
        HScroll2.Value = txtG.Text
        HScroll3.Value = txtB.Text
        Text2.Text = txtR.Text
        Text3.Text = txtG.Text
        Text4.Text = txtB.Text
        GoTo afterr
    End If
    If e > 1 Then
        ss = "s"
    End If
    MsgBox e & " error" & ss & ":" & vbCrLf & msgE & vbCrLf & vbCrLf & "Please correct the error" & ss & ".", vbCritical, "Error!"
End If
afterr:
txtR.Text = ""
txtG.Text = ""
txtB.Text = ""
Text2.SetFocus
End Sub

Private Sub Button1_Click()
If txtVBColor.Text = "" Or Len(txtVBColor.Text) <> 11 Then
    MsgBox "Value could not be copied."
Else
    Clipboard.Clear
    Clipboard.SetText txtVBColor.Text
    MsgBox txtVBColor.Text & " copied to Clipboard!", vbOKOnly, "Notification"
End If
txtVBColor.Text = ""
txtR.Text = ""
txtG.Text = ""
txtB.Text = ""
txtR.SetFocus
End Sub

Private Sub cmdConvert2_Click()
Dim p1, hR, hG, hB As String
If Text1.Text = "" Or Len(Text1.Text) <> 11 Then
    MsgBox "You either have not entered a value or you have and it is not valid.", vbCritical, "Error!"
    GoTo afterr2
End If
p1 = Right(Text1.Text, 10)
If Left(p1, 1) = "H" Then
    p1 = Right(p1, 9)
Else
    GoTo afterr2
End If
If Left(p1, 1) = "8" Then
    MsgBox "Sorry. This program doesn't convert system colors from this box.", vbCritical, "Not a Feature!"
    GoTo afterr2
Else
    opR.Text = ""
    opG.Text = ""
    opB.Text = ""
    p1 = Left(Right(p1, 7), 6)
End If
hR = Right(p1, 2)
hB = Mid(p1, 3, 2)
hG = Left(p1, 2)
opR.Text = Val("&H000000" & hR & "&")
opG.Text = Val("&H000000" & hB & "&")
opB.Text = Val("&H000000" & hG & "&")
picColor.BackColor = RGB(opR.Text, opG.Text, opB.Text)
Text5.Text = Val(Text1.Text)
Text6.Text = Text1.Text
Text7.Text = "#" & hR & hB & hG
HScroll1.Value = opR.Text
HScroll2.Value = opG.Text
HScroll3.Value = opB.Text
Text2.Text = opR.Text
Text3.Text = opG.Text
Text4.Text = opB.Text
opR.Text = ""
opG.Text = ""
opB.Text = ""
afterr2:
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Combo1_Click()
varRGB = GetSysColor(Combo1.ListIndex)
Text2.Text = varRGB And &HFF
Text3.Text = (varRGB And &H100FF00) / &H100
Text4.Text = (varRGB And &H1FF0000) / &H10000
End Sub

Private Sub Command1_Click()
Dim nosym As String
Dim length As Long
Dim i As Integer
On Error GoTo afterr5
i = 0
length = Len(Text8.Text)
If length = 0 Then
    MsgBox "You must enter a value first.", vbCritical, "Error!"
    Text8.SetFocus
    GoTo aftdone
ElseIf length = 7 Then
    nosym = Replace(Text8.Text, "#", "")
    If Len(nosym) = 7 Then
        MsgBox "Invalid entry.", vbCritical, "Error!"
        Text8.SetFocus
        GoTo aftdone
    End If
ElseIf length = 6 Then
    nosym = Replace(Text8.Text, "#", "")
    If Len(nosym) = 5 Then
        MsgBox "Invalid entry.", vbCritical, "Error!"
        Text8.SetFocus
        GoTo aftdone
    End If
Else
    MsgBox "Invalid length.", vbCritical, "Error!"
    Text8.SetFocus
    GoTo aftdone
End If
Text2.Text = Val("&H000000" & Left(nosym, 2) & "&")
Text3.Text = Val("&H000000" & Mid(nosym, 3, 2) & "&")
Text4.Text = Val("&H000000" & Right(nosym, 2) & "&")
Text8.Text = ""
GoTo aftdone
afterr5:
MsgBox Err.Description, vbCritical, "Error!"
aftdone:
End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText Text5.Text
End Sub

Private Sub Command3_Click()
Clipboard.Clear
Clipboard.SetText Text6.Text
End Sub

Private Sub Command5_Click()
Clipboard.Clear
Clipboard.SetText Text7.Text
End Sub

Private Sub Command7_Click()
Clipboard.Clear
Clipboard.SetText Text12.Text
End Sub

Private Sub Command8_Click()
Clipboard.Clear
Clipboard.SetText Text15.Text
End Sub

Private Sub Command9_Click()
Clipboard.Clear
Clipboard.SetText Text14.Text
End Sub

Private Sub Command10_Click()
Clipboard.Clear
Clipboard.SetText Text13.Text
End Sub

Private Sub Command13_Click()
Dim lowerbound As Long, upperbound As Long
lowerbound = 0
upperbound = 255
Text2.Text = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Text3.Text = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Text4.Text = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text8.Text = ""
txtVBColor.Text = ""
txtR.Text = ""
txtG.Text = ""
txtB.Text = ""
opR.Text = ""
opG.Text = ""
opB.Text = ""
picColor.BackColor = Val("&H00FFFFFF&")
End Sub

Private Sub Form_Load()
Dim min, rev As String
StatusBar1.SimpleText = "Created by: Matt DeKok    Last Updated: 02/25/06    Codes By: Me, dreamvb, Chetan Sarva, and a lot by Dan Redding - Blue Knot Software"
i = 0
min = ""
rev = ""
If App.Revision <> 0 Then rev = "." & App.Revision
If App.Minor <> 0 Then min = "." & App.Minor & rev
Me.Caption = Me.Caption & " v" & App.Major & min
End Sub

Private Sub timSpy_Timer()
Dim DeskTopWindow As Long, DeskTopDC As Long
Dim CurPos As PointAPI, ScreenPixel As Long
Dim strRed As String, strGreen As String
Dim strBlue As String, htmlformat As String

'use the getcursorpos api function to retrieve the current
'position of the cursor on the screen and set it to curpos
Call GetCursorPos(CurPos)

'this sets the desktop's dc in the DeskTopDC variable
DeskTopDC = GetDC(0)

'set the current pixel color in the ScreenPixel variable
'use GetPixel api function to retrieve colors from pixels
'and you use the DeskTopDC as the dc for it and we set
'the CurPos variable to hold the values of the position
'on the screen in pixels
ScreenPixel = GetPixel(DeskTopDC, CurPos.X, CurPos.Y)

'if the pictures backcolor doesn't allready = the currentolor
'pixel then dont add it to the picture box's backc
If picColor.BackColor <> ScreenPixel Then
 picColor.BackColor = ScreenPixel
End If

'set the txtColor's text to the background color of the pixel
Text2.Text = RGBRed(picColor.BackColor)
Text3.Text = RGBGreen(picColor.BackColor)
Text4.Text = RGBBlue(picColor.BackColor)
End Sub

Private Sub HScroll1_Change()
Call HScroll1_Scroll
End Sub

Private Sub HScroll2_Change()
Call HScroll2_Scroll
End Sub

Private Sub HScroll3_Change()
Call HScroll3_Scroll
End Sub

Private Sub HScroll1_Scroll()
Text2.Text = HScroll1.Value
End Sub

Private Sub HScroll2_Scroll()
Text3.Text = HScroll2.Value
End Sub

Private Sub HScroll3_Scroll()
Text4.Text = HScroll3.Value
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuMixer_Click()
frmMixer.Show
Me.Hide
End Sub

Private Sub QB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.Text = RGBRed(QBColor(Index))
Text3.Text = RGBGreen(QBColor(Index))
Text4.Text = RGBBlue(QBColor(Index))
End Sub

Private Sub Text1_Click()
If Text1.Text <> "" Then
    Text1.Text = ""
End If
If txtVBColor.Text <> "" Or opR.Text <> "" Or opG.Text <> "" Or opB.Text <> "" Then
    Call Command4_Click
End If
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Or IsNumeric(Text2.Text) = False Then
    Text2.Text = 0
End If
Text2.Text = Replace(Text2.Text, "&", "")
Text2.Text = Replace(Text2.Text, "$", "")
Text2.Text = Replace(Text2.Text, ".", "")
Text2.Text = Replace(Text2.Text, "+", "")
Text2.Text = Replace(Text2.Text, "-", "")
Text2.Text = Replace(Text2.Text, " ", "")
HScroll1.Value = Text2.Text
txtR.Text = Text2.Text
txtG.Text = Text3.Text
txtB.Text = Text4.Text
Call cmdConvert_click
txtVBColor.Text = ""
If Check1.Value = 0 Then
    Text9.Text = Text2.Text
Else
    Text9.Text = Hex(Text2.Text)
End If
Text9.SetFocus
If Check1.Value = 1 Then
    Label15.BackColor = Val("&H000000" & Text9.Text & "&")
Else
    Label15.BackColor = Val("&H000000" & Hex(Text9.Text) & "&")
End If
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Or IsNumeric(Text3.Text) = False Then
    Text3.Text = 0
End If
Text3.Text = Replace(Text3.Text, "&", "")
Text3.Text = Replace(Text3.Text, "$", "")
Text3.Text = Replace(Text3.Text, ".", "")
Text3.Text = Replace(Text3.Text, "+", "")
Text3.Text = Replace(Text3.Text, "-", "")
Text3.Text = Replace(Text3.Text, " ", "")
HScroll2.Value = Text3.Text
txtR.Text = Text2.Text
txtG.Text = Text3.Text
txtB.Text = Text4.Text
Call cmdConvert_click
txtVBColor.Text = ""
If Check1.Value = 0 Then
    Text10.Text = Text3.Text
Else
    Text10.Text = Hex(Text3.Text)
End If
Text10.SetFocus
If Check1.Value = 1 Then
    Label16.BackColor = Val("&H0000" & Text10.Text & "00&")
Else
    Label16.BackColor = Val("&H0000" & Hex(Text10.Text) & "00&")
End If
If Text3.Text > 172 Then
    Label16.ForeColor = 0
Else
    Label16.ForeColor = 16777215
End If
End Sub

Private Sub Text4_Change()
If Text4.Text = "" Or IsNumeric(Text4.Text) = False Then
    Text4.Text = 0
End If
Text4.Text = Replace(Text4.Text, "&", "")
Text4.Text = Replace(Text4.Text, "$", "")
Text4.Text = Replace(Text4.Text, ".", "")
Text4.Text = Replace(Text4.Text, "+", "")
Text4.Text = Replace(Text4.Text, "-", "")
Text4.Text = Replace(Text4.Text, " ", "")
HScroll3.Value = Text4.Text
txtR.Text = Text2.Text
txtG.Text = Text3.Text
txtB.Text = Text4.Text
Call cmdConvert_click
txtVBColor.Text = ""
If Check1.Value = 0 Then
    Text11.Text = Text4.Text
Else
    Text11.Text = Hex(Text4.Text)
End If
Text11.SetFocus
If Check1.Value = 1 Then
    Label17.BackColor = Val("&H00" & Text11.Text & "0000&")
Else
    Label17.BackColor = Val("&H00" & Hex(Text11.Text) & "0000&")
End If
End Sub

Private Sub Text5_Change()
Dim hexR As String, hexG As String, hexB As String, hexes As String
Label13.Caption = RGBtoHSL(Text5.Text).Hue
Label14.Caption = RGBtoHSL(Text5.Text).Sat
Label18.Caption = RGBtoHSL(Text5.Text).Lum
hexes = String$(8 - Len(Hex$(Text5.Text)), "0") & Hex$(Text5.Text)
hexR = StrReverse(Hex$(RGBRed(Text5.Text)))
If Len(hexR) = 1 Then hexR = hexR & "0"
hexG = StrReverse(Hex$(RGBGreen(Text5.Text)))
If Len(hexG) = 1 Then hexG = hexG & "0"
hexB = StrReverse(Hex$(RGBBlue(Text5.Text)))
If Len(hexB) = 1 Then hexB = hexB & "0"
Text12.Text = "0x" & hexes
Text13.Text = "$" & hexes
Text14.Text = "0x" & hexR & hexG & Replace(hexB, "0", "")
Text15.Text = hexR & hexG & hexB
End Sub

Private Sub Text5_Click()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Text6_Click()
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Text7_Click()
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)
End Sub

Private Sub Text8_Click()
If Text8.Text <> "" Then
    Text8.Text = ""
End If
If txtVBColor.Text <> "" Or opR.Text <> "" Or opG.Text <> "" Or opB.Text <> "" Then
    Call Command4_Click
End If
End Sub

Private Sub Text9_Change()
On Error GoTo afterr5
If Check1.Value = 0 Then
    If Text9.Text < 0 Then Text9.Text = 0
    If Text9.Text > 255 Then Text9.Text = 255
    Text2.Text = Text9.Text
Else
    If Len(Text9.Text) > 2 Then Text9.Text = Left(Text9.Text, 2)
    Text2.Text = CDec("&H" & Text9.Text)
End If
Exit Sub
afterr5:
Text9.Text = 0
End Sub

Private Sub Text10_Change()
On Error GoTo afterr5
If Check1.Value = 0 Then
    If Text10.Text < 0 Then Text10.Text = 0
    If Text10.Text > 255 Then Text10.Text = 255
    Text3.Text = Text10.Text
Else
    If Len(Text10.Text) > 2 Then Text10.Text = Left(Text10.Text, 2)
    Text3.Text = CDec("&H" & Text10.Text)
End If
Exit Sub
afterr5:
Text10.Text = 0
End Sub

Private Sub Text11_Change()
On Error GoTo afterr5
If Check1.Value = 0 Then
    If Text11.Text < 0 Then Text11.Text = 0
    If Text11.Text > 255 Then Text11.Text = 255
    Text4.Text = Text11.Text
Else
    If Len(Text11.Text) > 2 Then Text11.Text = Left(Text11.Text, 2)
    Text4.Text = CDec("&H" & Text11.Text)
End If
Exit Sub
afterr5:
Text11.Text = 0
End Sub

Private Sub Text9_Click()
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub

Private Sub Text10_Click()
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
End Sub

Private Sub Text11_Click()
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)
End Sub

Private Sub txtVBColor_GotFocus()
txtR.Text = ""
txtG.Text = ""
txtB.Text = ""
Call Button1_Click
txtR.SetFocus
End Sub

Private Sub SplitThem(ByVal Red, Gre, Blu, RedGreBlu As Long)
Red = RedGreBlu And &HFF
Gre = (RedGreBlu \ 2 ^ 8) And &HFF
Blu = (RedGreBlu \ 2 ^ 16) And &HFF
End Sub
