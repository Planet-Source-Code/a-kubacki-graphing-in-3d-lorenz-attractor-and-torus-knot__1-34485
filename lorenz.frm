VERSION 5.00
Begin VB.Form mainfrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "3D Graphing Demonstration"
   ClientHeight    =   6930
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "5"
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "3"
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   120
      Max             =   2000
      Min             =   1
      TabIndex        =   0
      Top             =   360
      Value           =   500
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   2040
      ScaleHeight     =   2775
      ScaleWidth      =   3135
      TabIndex        =   8
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Please select a graph from the ""Graph"" menu"
      ForeColor       =   &H000000FF&
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Torus knot controls:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Q:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "P:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Number of points:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu XX 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu SH 
         Caption         =   "Show in-program help"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu graph 
      Caption         =   "Graph"
      Begin VB.Menu LA 
         Caption         =   "Lorenz Attractor"
      End
      Begin VB.Menu TK 
         Caption         =   "Torus Knot"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu inst 
         Caption         =   "Instructions"
      End
      Begin VB.Menu AB 
         Caption         =   "About...."
      End
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer
Dim m As Integer
Dim asdf As Integer
Dim yang As Double
Dim xang As Double
Dim zang As Double
Dim mdn As Boolean
Dim counter As Integer
Dim color As Long
Dim needupd As Boolean
Dim mcax As Integer, mcay As Integer, mcbx As Integer, mcby As Integer
Dim rat(10000) As Double
Dim X(10000) As Double
Dim Y(10000) As Double
Dim z(10000) As Double
Dim xa As Double
Dim ya As Double
Dim za As Double
Dim xb As Double
Dim yb As Double
Dim zb As Double
Dim xr(10000) As Double
Dim yr(10000) As Double
Dim h As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim gm As Integer

Const pi = 3.1415926



Private Sub torus()
Dim j As Integer
Dim t As Double
Dim p As Integer
Dim q As Integer


p = Text1.Text
q = Text2.Text

counter = 0

For t = 0 To 2 * pi * p Step (2 * pi * p) / N
counter = counter + 1
X(counter) = 40 * Cos(t) * (1 + 0.5 * Cos((q / p) * t))
Y(counter) = 40 * Sin(t) * (1 + 0.5 * Cos((q / p) * t))
z(counter) = 40 * 0.5 * Sin((q / p) * t)
Next t
asdf = counter

End Sub
Private Sub lorenz()
Dim j As Integer
counter = 0

For j = 1 To N
xa = xb + h * a * (yb - xb)
ya = yb + h * (xb * (b - zb) - yb)
za = zb + h * (xb * yb - c * zb)

xb = xa
yb = ya
zb = za


counter = counter + 1
X(counter) = xb
Y(counter) = yb
z(counter) = zb
asdf = counter

Next j




End Sub





Private Sub AB_Click()
Form1.Show
mainfrm.Hide

End Sub

Private Sub Form_Load()
Picture1.Width = ScaleWidth - 2500
Picture1.Height = ScaleHeight - 1600
Picture1.Left = 2500
Picture1.Top = 1600


h = 0.02
a = 10
b = 28
c = 8 / 3
xb = 0.1
yb = 0
zb = 0
xang = 0
yang = 0
zang = 0
N = 500
scrollupd

needupd = True

End Sub
Private Sub rotate()
Dim i As Integer
Dim cx, sx, sy, cy, sz, cz As Double
Dim xt, yt, zt As Double
cx = Cos(xang)
sx = Sin(xang)
sy = Sin(yang)
cy = Cos(yang)
sz = Sin(zang)
cz = Cos(zang)

For i = 1 To asdf

yt = Y(i) * cx - z(i) * sx
zt = Y(i) * sx + z(i) * cx
Y(i) = yt
z(i) = zt
xt = X(i) * cy - z(i) * sy
zt = X(i) * sy + z(i) * cy
X(i) = xt
z(i) = zt
xt = X(i) * cz - Y(i) * sz
yt = X(i) * sz + Y(i) * cz
X(i) = xt
Y(i) = yt
Next i

End Sub
Private Sub transform()
Dim k As Integer

For k = 1 To asdf
xr(k) = X(k) * (12 / (-z(k) / 10 + 36))
yr(k) = Y(k) * (26 / (-z(k) / 10 + 36))
Next k

End Sub

Private Sub draw()
Dim l As Integer
Picture1.Cls

For l = 1 To asdf - 1
color = RGB(128 + 1.5 * z(l), 0, 0)
Picture1.Line (xr(l) * Picture1.Width / 70 + Picture1.Width / 2, yr(l) * Picture1.Height / 100 + Picture1.Height / 2)-(xr(l + 1) * Picture1.Width / 70 + Picture1.Width / 2, yr(l + 1) * Picture1.Height / 100 + Picture1.Height / 2), color

Next l
If gm = 2 Then
Picture1.Line (xr(1) * Picture1.Width / 70 + Picture1.Width / 2, yr(1) * Picture1.Height / 100 + Picture1.Height / 2)-(xr(asdf) * Picture1.Width / 70 + Picture1.Width / 2, yr(asdf) * Picture1.Height / 100 + Picture1.Height / 2), color
End If

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdn = True
mcax = X
mcay = Y
If needupd = True Then
 If gm = 1 Then
 lorenz
 ElseIf gm = 2 Then
 torus
 End If
 
 
End If
needupd = False

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mdn = True Then

mcbx = X
mcby = Y

xang = (mcay - mcby) / 1000
yang = (mcax - mcbx) / 1000
zang = 0
mcax = X
mcay = Y

rotate
transform
draw
End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdn = False

End Sub

Private Sub scrollupd()
N = HScroll1.Value
Label1.Caption = "Number of points: " & N

End Sub

Private Sub Form_Resize()
Picture1.Width = ScaleWidth - 2500
Picture1.Height = ScaleHeight - 1600
Picture1.Left = 2500
Picture1.Top = 1600
Label4.Top = ScaleHeight - Label4.Height


End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()
needupd = True

scrollupd
End Sub



Private Sub inst_Click()
Form2.Show
mainfrm.Hide

End Sub

Private Sub LA_Click()
lorenz
gm = 1
Label3.Visible = False
Label2(1).Visible = False
Label2(2).Visible = False
Text1.Visible = False
Text2.Visible = False
rotate
transform
draw
LA.Checked = True
TK.Checked = False

Label4.Caption = "The Lorenz Attractor is a graph that represents three differential equations:                                                                       dx/dt=p(y-x)                              dy/dt=-xz+rx-y                              dz/dt=xy-bz                                                                             Where P, R, dt, and B are constants.   This graph is used in meteorology to demonstrate how changing one small variable results in chaotic behavior and nearly unpredictable results.  For this reason, this graph has also been used to demonstrate the chaos theory.  In fact, the ""butterfly effect"" was started by this very graph because it looks like a butterfly from certain angles."
End Sub


Private Sub SH_Click()

If Label4.Visible = True Then
Label4.Visible = False
SH.Checked = False
Picture1.Width = ScaleWidth
Picture1.Height = ScaleHeight - 1600
Picture1.Left = 0
Picture1.Top = 1600
ElseIf Label4.Visible = False Then
Label4.Visible = True
SH.Checked = True
Picture1.Width = ScaleWidth - 2500
Picture1.Height = ScaleHeight - 1600
Picture1.Left = 2500
Picture1.Top = 1600
End If

rotate
transform
draw


End Sub

Private Sub Text1_Change()
needupd = True

End Sub
Private Sub Text2_Change()
needupd = True

End Sub


Private Sub TK_Click()
torus
gm = 2
Label3.Visible = True
Label2(1).Visible = True
Label2(2).Visible = True
Text1.Visible = True
Text2.Visible = True
rotate
transform
draw
TK.Checked = True
LA.Checked = False

Label4.Caption = "The Torus Knot is a representation of what happens when a line is wrapped on the outside of a torus, or a donut.  The ""P"" value represents the number of horizontal passes along the torus and the ""Q"" value represents the number of vertical passes around the torus."
End Sub



Private Sub XX_Click()
End

End Sub
