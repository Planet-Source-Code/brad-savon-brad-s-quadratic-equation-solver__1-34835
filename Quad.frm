VERSION 5.00
Begin VB.Form frmQuad 
   Caption         =   "Quadradic Equation"
   ClientHeight    =   5280
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtImag2 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   67
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtImag1 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   66
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Fram13 
      Caption         =   "2 * A"
      Height          =   615
      Left            =   5280
      TabIndex        =   63
      Top             =   5640
      Width           =   1095
      Begin VB.Label lblWork8 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "- ANS/2*A"
      Height          =   615
      Left            =   4440
      TabIndex        =   61
      Top             =   6480
      Width           =   1215
      Begin VB.Label lblWork7 
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "+ ANS/2*A"
      Height          =   615
      Left            =   3120
      TabIndex        =   59
      Top             =   6480
      Width           =   1215
      Begin VB.Label lblWork6 
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "-B - SQR"
      Height          =   615
      Left            =   1920
      TabIndex        =   57
      Top             =   6480
      Width           =   1095
      Begin VB.Label lblWork5 
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "-B + SQR"
      Height          =   615
      Left            =   720
      TabIndex        =   55
      Top             =   6480
      Width           =   1095
      Begin VB.Label lblWork4 
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Square Root"
      Height          =   615
      Left            =   3960
      TabIndex        =   53
      Top             =   5640
      Width           =   1215
      Begin VB.Label lblWork3 
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "B^2 - 4*A*C"
      Height          =   615
      Left            =   2640
      TabIndex        =   51
      Top             =   5640
      Width           =   1215
      Begin VB.Label lblWork2 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "4*A*C"
      Height          =   615
      Left            =   1440
      TabIndex        =   49
      Top             =   5640
      Width           =   1095
      Begin VB.Label lblWork1 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "B^2"
      Height          =   615
      Left            =   240
      TabIndex        =   47
      Top             =   5640
      Width           =   1095
      Begin VB.Label lblWork 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Work"
      Height          =   255
      Left            =   5280
      TabIndex        =   46
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox lstTable 
      Height          =   1425
      Left            =   1680
      TabIndex        =   38
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   3840
      ScaleHeight     =   2235
      ScaleWidth      =   2475
      TabIndex        =   37
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Caption         =   "Discriminant"
      Height          =   615
      Left            =   3360
      TabIndex        =   28
      Top             =   1320
      Width           =   1575
      Begin VB.Label lblDis5 
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDis4 
         Caption         =   "="
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblDis3 
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDis2 
         Caption         =   "-"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblDis1 
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtAnswer1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "C"
      Height          =   735
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   1095
      Begin VB.TextBox txtThree 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtTwo 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "B"
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1455
      Begin VB.Label Label2 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox txtOne 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "A"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      Begin VB.Label Label1 
         Caption         =   "x^2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label lblImag 
      Caption         =   "Imaginary Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   65
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Y"
      Height          =   375
      Left            =   2640
      TabIndex        =   45
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "X"
      Height          =   375
      Left            =   1920
      TabIndex        =   44
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblGra3 
      Height          =   495
      Left            =   5880
      TabIndex        =   43
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label lblGra2 
      Height          =   375
      Left            =   5520
      TabIndex        =   42
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label lblGra 
      Height          =   375
      Left            =   5040
      TabIndex        =   40
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "Graph for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   39
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblEight 
      Height          =   375
      Left            =   5160
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblNeg 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   35
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lbl2a 
      Height          =   255
      Left            =   1440
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line7 
      X1              =   2760
      X2              =   2640
      Y1              =   1800
      Y2              =   1680
   End
   Begin VB.Line Line6 
      X1              =   2760
      X2              =   2640
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line5 
      X1              =   2640
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lbl4ac 
      Height          =   255
      Left            =   2040
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblb2 
      Height          =   255
      Left            =   1560
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblEleven 
      Height          =   615
      Left            =   5880
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblTen 
      Height          =   615
      Left            =   2400
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblNine 
      Height          =   615
      Left            =   5880
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblSeven 
      Height          =   615
      Left            =   3960
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblSix 
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblFive 
      Height          =   615
      Left            =   2400
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFour 
      Height          =   495
      Left            =   1800
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblThree 
      Height          =   495
      Left            =   1320
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblTwo 
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOne 
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "X="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "2a"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "x ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "- b +_"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   840
      X2              =   2640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Caption         =   "b^2  -  4ac"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   1440
      X2              =   2400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   1440
      Y1              =   1800
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   1440
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Label lblGra1 
      Caption         =   "^2 +    x+      = 0"
      Height          =   495
      Left            =   5160
      TabIndex        =   41
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Clear 
         Caption         =   "Clear"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmQuad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
Dim Message As String
Message = "Quadradic Equation Solver" + vbCr
Message = Message + "Created by: Brad Savon" + vbCr
Message = Message + "May 12, 2002" + vbCr
Message = Message + "" + vbCr
Message = Message + "Questions:" + vbCr
Message = Message + "Email me at: bsavon@hotmail.com" + vbCr
MsgBox Message, vbOKOnly + vbInformation, "Quadradic Equation Solver"
End Sub

Private Sub Check1_Click()
If Check1.Value = Checked Then
frmQuad.Height = 8250
Else
If Check1.Value = Unchecked Then
frmQuad.Height = 6075
End If
End If
End Sub

Private Sub Clear_Click()
'Clear the text boxes
txtOne.Text = ""
txtTwo.Text = ""
txtThree.Text = ""
txtAnswer1.Text = ""
txtAnswer2.Text = ""
lbl4ac.Visible = False
lblSix.Visible = False
lblb2.Visible = False
lbl2a.Visible = False
lblDis1.Visible = False
lblDis3.Visible = False
lblDis5.Visible = False
lblNeg.Caption = ""
lblGra1.Visible = False
lblGra.Caption = ""
lblGra2.Caption = ""
lblGra3.Caption = ""
lblWork.Caption = ""
lblWork1.Caption = ""
lblWork2.Caption = ""
lblWork3.Caption = ""
lblWork4.Caption = ""
lblWork5.Caption = ""
lblWork6.Caption = ""
lblWork7.Caption = ""
lblWork8.Caption = ""
lstTable.Clear
Picture1.Cls
End Sub

Private Sub cmdCalculate_Click()
lstTable.Clear
Picture1.Cls
If txtOne.Text = "" Then
MsgBox "Enter your Number", vbOKOnly + vbCritical, "Number Error"
Else
If txtTwo.Text = "" Then
MsgBox "Enter your Number", vbOKOnly + vbCritical, "Number Error"
Else
If txtThree.Text = "" Then
MsgBox "Enter your Number", vbOKOnly + vbCritical, "Number Error"
Else
lblOne.Caption = txtTwo.Text * txtTwo.Text 'B Squared
lblTwo.Caption = 4 * txtOne.Text * txtThree.Text '4*A*C
lblThree.Caption = Val(lblOne.Caption) - Val(lblTwo.Caption) 'b^2-4AC
lblFour.Caption = Abs(lblThree.Caption) 'ABSOLUTE Value of b^2-4AC
lblFive.Caption = Sqr(lblFour.Caption) 'Square Root of Absolute Value
lblSix.Caption = -(txtTwo.Text) 'Opposite of B
lblSeven.Caption = Val(lblSix.Caption) + Val(lblFive.Caption) 'Opposite of B + Square Root
lblEight.Caption = 2 * (txtOne.Text) '2*A
lblNine.Caption = Val(lblSeven.Caption) / Val(lblEight.Caption) 'Addition / 2A
lblTen.Caption = Val(lblSix.Caption) - Val(lblFive.Caption) 'Opposite of B - Square Root
lblEleven.Caption = Val(lblTen.Caption) / Val(lblEight.Caption) 'Subtraction / 2A
txtAnswer1.Text = Val(lblNine.Caption) '1st Equation
txtAnswer2.Text = Val(lblEleven.Caption) '2nd Equation
lblb2.Caption = Val(txtTwo.Text) * Val(txtTwo.Text) 'b2 Value in the Example equation
lbl4ac.Caption = 4 * Val(txtOne.Text) * Val(txtThree.Text) '4ac Value in the Example Equation
lbl4ac.Visible = True 'Show new Equation
lblSix.Visible = True 'Show new Equation
lblb2.Visible = True 'Show new Equation
lbl2a.Visible = True 'Show new Equation
lblDis1.Visible = True 'Show New Discriminant
lblDis3.Visible = True 'Show New Discriminant
lblDis5.Visible = True 'Show New Discriminant
lblGra1.Visible = True 'Show Graph Equation
lblDis1.Caption = Val(lblb2.Caption) 'Discriminant
lblDis3.Caption = Val(lbl4ac.Caption) 'Discriminant
lblDis5.Caption = Val(lblb2.Caption) - Val(lbl4ac.Caption) 'Discriminant
lbl2a.Caption = 2 * Val(txtOne.Text) 'Discriminant
If Val(lblDis1.Caption) < Val(lblDis3.Caption) Then
lblNeg.Caption = "Negative"
lblNeg.ForeColor = vbRed
lblImag.Visible = True
txtImag1.Visible = True
txtImag2.Visible = True
txtImag1.Text = Val(lblNine.Caption)
txtImag2.Text = Val(lblEleven.Caption)
txtAnswer1.Text = "No Real Solutions"
txtAnswer2.Text = "No Real Solutions"
lblWork3.Caption = "Impossible"
lblWork4.Caption = "Impossible"
lblWork5.Caption = "Impossible"
lblWork6.Caption = "Impossible"
lblWork7.Caption = "Impossible"
Picture1.Cls
lstTable.Clear
lblGra.Caption = ""
lblGra2.Caption = ""
lblGra3.Caption = ""
Else
If Val(lblDis1.Caption) > Val(lblDis3.Caption) Then
lblNeg.Caption = "Positive"
lblNeg.ForeColor = vbGreen
lblImag.Visible = False
txtImag1.Visible = False
txtImag2.Visible = False
lblWork.Caption = Val(txtTwo.Text) ^ 2 'Show Work
lblWork1.Caption = 4 * Val(txtOne.Text) * Val(txtTwo.Text) 'Show Work
lblWork2.Caption = Val(lblOne.Caption) - Val(lblTwo.Caption) 'Show Work
lblWork8.Caption = 2 * Val(txtOne.Text)
lblWork3.Caption = Sqr(lblWork2.Caption)
lblWork4.Caption = -Val(txtTwo.Text) + Val(lblWork3.Caption)
lblWork5.Caption = -Val(txtTwo.Text) - Val(lblWork3.Caption)
lblWork6.Caption = Val(lblWork4.Caption) / Val(lblWork8.Caption)
lblWork7.Caption = Val(lblWork5.Caption) / Val(lblWork8.Caption)
lblGra.Caption = Val(txtOne.Text)
lblGra2.Caption = Val(txtTwo.Text)
lblGra3.Caption = Val(txtThree.Text)
Dim x As Single, y As Single
Dim PointNumber As Integer
Const INCREMENT = 0.01
ReDim YVals(1 To 21 / INCREMENT) As Single
PointNumber = 1
For x = -5 To 5 Step INCREMENT
y = Val(txtOne.Text) * x * x + Val(txtTwo.Text) * x + Val(txtThree.Text)
YVals(PointNumber) = y
PointNumber = PointNumber + 1
lstTable.AddItem Format(x, "Fixed") & vbTab & Format(y, "Fixed")
Next x
Picture1.Scale (-10, 20)-(10, -20)
Picture1.Line (0, -20)-(0, 20), RGB(0, 200, 0)
Picture1.Line (-10, 0)-(10, 0), RGB(0, 200, 0)
PointNumber = 1
For x = -5 To 5 Step INCREMENT
Picture1.PSet (x, YVals(PointNumber))
PointNumber = PointNumber + 1
Next x
End If
End If
End If
End If
End If
End Sub

Private Sub Quit_Click()
End
End Sub

Private Sub txtOne_KeyPress(KeyAscii As Integer) 'ONLY ALOW NUMBERS
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyDecPt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtThree_KeyPress(KeyAscii As Integer) 'ONLY ALOW NUMBERS
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyDecPt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtTwo_KeyPress(KeyAscii As Integer) 'ONLY ALOW NUMBERS
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyDecPt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
