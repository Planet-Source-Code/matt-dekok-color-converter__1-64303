VERSION 5.00
Begin VB.Form frmMixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Mixer"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   9975
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "255"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "128"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "33023"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "&H000080FF&"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "#FF8000"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "65535"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "255"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label13 
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
      Left            =   8040
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label11 
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
      Left            =   8040
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label9 
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
      Left            =   8040
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label10 
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
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label12 
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
      Left            =   8040
      TabIndex        =   9
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label8 
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
      Left            =   8040
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   5280
      X2              =   5640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   5280
      X2              =   5640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   2580
      X2              =   2580
      Y1              =   960
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2400
      X2              =   2760
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r1, g1, b1, r2, g2, b2 As Integer, clr As Long
Private Sub Form_Load()
Me.Icon = frmRGB2VB2RGB.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmRGB2VB2RGB.Show
End Sub

Private Sub Label5_Change()
Label4.Caption = RGB(Label5.Caption, Label6.Caption, Label7.Caption)
End Sub

Private Sub Label6_Change()
Label4.Caption = RGB(Label5.Caption, Label6.Caption, Label7.Caption)
End Sub

Private Sub Label7_Change()
Label4.Caption = RGB(Label5.Caption, Label6.Caption, Label7.Caption)
End Sub

Private Sub Label4_Change()
Label3.BackColor = Label4.Caption
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Or IsNumeric(Text1.Text) = False Then
    Text1.Text = 0
End If
Text1.Text = Replace(Text1.Text, "&", "")
Text1.Text = Replace(Text1.Text, "$", "")
Text1.Text = Replace(Text1.Text, ".", "")
Text1.Text = Replace(Text1.Text, "+", "")
Text1.Text = Replace(Text1.Text, "-", "")
Text1.Text = Replace(Text1.Text, " ", "")
If Text1.Text < 0 Then Text1.Text = 0
If Text1.Text > 16777215 Then Text1.Text = 16777215
Label1.BackColor = Text1.Text
Call Mix
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
r2 = Val(Text2.Text And &HFF&)
g2 = Val((Text2.Text And &HFF00&) / &H100&)
b2 = Val((Text2.Text And &HFF0000) / &H10000)
If Text2.Text < 0 Then Text2.Text = 0
If Text2.Text > 16777215 Then Text2.Text = 16777215
Label2.BackColor = Text2.Text
Call Mix
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Text7_GotFocus()
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)
End Sub

Private Sub Text8_GotFocus()
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)
End Sub

Private Sub Mix()
Dim r, g, b As Integer, zr, zg, zb As String
r1 = Val(Text1.Text And &HFF&)
g1 = Val((Text1.Text And &HFF00&) / &H100&)
b1 = Val((Text1.Text And &HFF0000) / &H10000)
r2 = Val(Text2.Text And &HFF&)
g2 = Val((Text2.Text And &HFF00&) / &H100&)
b2 = Val((Text2.Text And &HFF0000) / &H10000)
r = Round(Val(r1 + r2) / 2, 0)
g = Round(Val(g1 + g2) / 2, 0)
b = Round(Val(b1 + b2) / 2, 0)
Text3.Text = RGB(r, g, b)
zr = ""
zg = ""
zb = ""
If r < 16 Then zr = "0"
If g < 16 Then zg = "0"
If b < 16 Then zb = "0"
Text4.Text = "&H00" & zb & Hex(b) & zg & Hex(g) & zr & Hex(r) & "&"
Text5.Text = "#" & zr & Hex(r) & zg & Hex(g) & zb & Hex(b)
Text6.Text = r
Text7.Text = g
Text8.Text = b
Label3.BackColor = Text3.Text
End Sub
