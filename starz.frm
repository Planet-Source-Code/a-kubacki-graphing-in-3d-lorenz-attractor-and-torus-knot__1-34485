VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000008&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"starz.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(1000) As Integer
Dim Y(1000) As Integer
Dim z(1000) As Integer
Dim xv(1000) As Long
Dim yv(1000) As Long
Dim sss As Long
Dim numstarz As Integer

Dim colora As Long
Dim i, j, k, m As Integer





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then


Form1.Hide

mainfrm.Show

End If

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 1
mainfrm.Hide
Form1.Show

numstarz = 250
Form1.Enabled = True


init

 
End Sub
Private Sub init()

For i = 1 To numstarz
X(i) = Rnd() * 1000 - 500
Y(i) = Rnd() * 1000 - 500
z(i) = 1 + Rnd() * 256
Next i

End Sub


Private Sub drawptz()

Dim asdf As Integer

sss = 40 * ScaleWidth / 1000
For k = 1 To numstarz
xv(k) = X(k) * 3 * sss / ((z(k)))
yv(k) = Y(k) * 3 * sss / ((z(k)))
Next k

For j = 1 To numstarz
z(j) = z(j) - 1

    If z(j) < 40 Then
        
        z(j) = 200 + Rnd() * 40
    End If

Next j

For i = 1 To numstarz
asdf = 256 - z(i)
colora = RGB(asdf, asdf, asdf)
PSet (xv(i) + ScaleWidth / 2, yv(i) + ScaleHeight / 2), colora

Next i

End Sub



Private Sub Form_Unload(Cancel As Integer)
mainfrm.Show

End Sub

Private Sub Timer1_Timer()
 Form1.Enabled = True
 Timer1.Enabled = True

For m = 1 To 10

Cls


drawptz
Next m

End Sub


