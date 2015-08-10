VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "两地距离计算器 - By Zaodie"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "重置"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "输入具体经纬度"
      Height          =   3015
      Left            =   4680
      TabIndex        =   41
      Top             =   600
      Width           =   5175
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   4935
         TabIndex        =   42
         Top             =   360
         Width           =   4935
         Begin VB.Frame Frame4 
            Caption         =   "地点1"
            Height          =   1215
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   4935
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   2520
               ScaleHeight     =   855
               ScaleWidth      =   375
               TabIndex        =   50
               Top             =   240
               Width           =   375
               Begin VB.CommandButton p1yc 
                  Caption         =   "←"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   9
                  Top             =   480
                  Width           =   375
               End
               Begin VB.CommandButton p1xc 
                  Caption         =   "←"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   4
                  Top             =   120
                  Width           =   375
               End
            End
            Begin VB.TextBox p1y 
               Height          =   270
               Left            =   840
               TabIndex        =   49
               Tag             =   "1"
               Text            =   "25"
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox p1x 
               Height          =   270
               Left            =   840
               TabIndex        =   48
               Tag             =   "1"
               Text            =   "25"
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox p1x2 
               Height          =   270
               Left            =   3480
               TabIndex        =   2
               Tag             =   "1"
               Text            =   "54"
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox p1x3 
               Height          =   270
               Left            =   3960
               TabIndex        =   3
               Tag             =   "1"
               Text            =   "57"
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox p1y2 
               Height          =   270
               Left            =   3480
               TabIndex        =   7
               Tag             =   "1"
               Text            =   "23"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox p1y3 
               Height          =   270
               Left            =   3960
               TabIndex        =   8
               Tag             =   "1"
               Text            =   "26"
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox p1y1 
               Height          =   270
               Left            =   3000
               TabIndex        =   6
               Tag             =   "1"
               Text            =   "116"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox p1x1 
               Height          =   270
               Left            =   3000
               TabIndex        =   1
               Tag             =   "1"
               Text            =   "39"
               Top             =   360
               Width           =   495
            End
            Begin VB.ComboBox p1y0 
               Height          =   300
               ItemData        =   "Form1.frx":0000
               Left            =   120
               List            =   "Form1.frx":0002
               TabIndex        =   5
               Tag             =   "1"
               Text            =   "E"
               Top             =   720
               Width           =   615
            End
            Begin VB.ComboBox p1x0 
               Height          =   300
               ItemData        =   "Form1.frx":0004
               Left            =   120
               List            =   "Form1.frx":0006
               TabIndex        =   0
               Tag             =   "1"
               Text            =   "N"
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "地点2"
            Height          =   1215
            Left            =   0
            TabIndex        =   43
            Top             =   1320
            Width           =   4935
            Begin VB.ComboBox p2x0 
               Height          =   300
               ItemData        =   "Form1.frx":0008
               Left            =   120
               List            =   "Form1.frx":000A
               TabIndex        =   10
               Tag             =   "1"
               Text            =   "S"
               Top             =   360
               Width           =   615
            End
            Begin VB.ComboBox p2y0 
               Height          =   300
               ItemData        =   "Form1.frx":000C
               Left            =   120
               List            =   "Form1.frx":000E
               TabIndex        =   15
               Tag             =   "1"
               Text            =   "E"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox p2x1 
               Height          =   270
               Left            =   3000
               TabIndex        =   11
               Tag             =   "1"
               Text            =   "33"
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox p2y1 
               Height          =   270
               Left            =   3000
               TabIndex        =   16
               Tag             =   "1"
               Text            =   "151"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox p2y3 
               Height          =   270
               Left            =   3960
               TabIndex        =   18
               Tag             =   "1"
               Text            =   "40"
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox p2y2 
               Height          =   270
               Left            =   3480
               TabIndex        =   17
               Tag             =   "1"
               Text            =   "12"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox p2x3 
               Height          =   270
               Left            =   3960
               TabIndex        =   13
               Tag             =   "1"
               Text            =   "35.9"
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox p2x2 
               Height          =   270
               Left            =   3480
               TabIndex        =   12
               Tag             =   "1"
               Text            =   "51"
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox p2x 
               Height          =   270
               Left            =   840
               TabIndex        =   46
               Tag             =   "1"
               Text            =   "25"
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox p2y 
               Height          =   270
               Left            =   840
               TabIndex        =   45
               Tag             =   "1"
               Text            =   "25"
               Top             =   720
               Width           =   1575
            End
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   2520
               ScaleHeight     =   855
               ScaleWidth      =   375
               TabIndex        =   44
               Top             =   240
               Width           =   375
               Begin VB.CommandButton p2xc 
                  Caption         =   "←"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   14
                  Top             =   120
                  Width           =   375
               End
               Begin VB.CommandButton p2yc 
                  Caption         =   "←"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   19
                  Top             =   480
                  Width           =   375
               End
            End
         End
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "←填入"
      Height          =   1335
      Left            =   3360
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   6480
      TabIndex        =   23
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "输出"
      Height          =   1095
      Left            =   360
      TabIndex        =   36
      Top             =   3720
      Width           =   3615
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1560
         TabIndex        =   39
         Tag             =   "1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1560
         TabIndex        =   37
         Tag             =   "1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "球面距离(弧长)"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "直线距离(弦长)"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "所需变量"
      Height          =   1455
      Left            =   360
      TabIndex        =   32
      Top             =   1560
      Width           =   2895
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   840
         TabIndex        =   28
         Tag             =   "1"
         Text            =   "13"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   840
         TabIndex        =   26
         Tag             =   "1"
         Text            =   "25"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   840
         TabIndex        =   27
         Tag             =   "1"
         Text            =   "39"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "纬度1"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "纬度2"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "经度差"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "常数取值"
      Height          =   1095
      Left            =   360
      TabIndex        =   29
      Top             =   240
      Width           =   2895
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1080
         TabIndex        =   24
         Text            =   "3.1415926535898"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1080
         TabIndex        =   25
         Text            =   "6371.004"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "圆周率π"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "地球半径r"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   810
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算↓"
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo hander
Dim a As Double
Dim b As Double
Dim c As Double
Dim r As Double
Dim pi As Double
Dim k As Double

pi = CDbl(Text5)
a = CDbl(Text1) * pi / 180
b = CDbl(Text3) * pi / 180
c = CDbl(Text2) * pi / 180
r = CDbl(Text4)

k = Sqr((Sin(c) - Sin(a)) ^ 2 + (Cos(c) - Cos(a) * Cos(b)) ^ 2 + (Cos(a) * Sin(b)) ^ 2)
Text6 = r * k

Text9 = r * ArcSin((k / 2)) * 2
Debug.Print Sin(a)
Debug.Print Cos(a)
Debug.Print Sin(c)
Debug.Print Cos(c)
Debug.Print Sin(b)
Debug.Print Cos(b)

hander:

Exit Sub
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Text1 = ns(p1x0.Text) & p1x
Text2 = ns(p2x0.Text) & p2x

If p2y0.Text = p1y0.Text Then
Text3.Text = Abs(Val(p2y) - Val(p1y))
Else
Text3.Text = 360 - Val(p2y) - Val(p1y)
End If
End Sub

Private Sub Command4_Click()
For Each i In Me.Controls
If Val(i.Tag) Then i.Text = ""
Next
End Sub

Private Sub Form_Load()
p1x0.AddItem "N"
p1x0.AddItem "S"
p1y0.AddItem "E"
p1y0.AddItem "W"
p2x0.AddItem "N"
p2x0.AddItem "S"
p2y0.AddItem "E"
p2y0.AddItem "W"
p1xc_Click
p2xc_Click
p1yc_Click
p2yc_Click
Command3_Click
End Sub

Private Sub p2xa_Change()

End Sub

Private Sub p1xc_Click()
p1x = p1x1 + p1x2 / 60 + p1x3 / 3600
End Sub
Private Sub p1yc_Click()
p1y = p1y1 + p1y2 / 60 + p1y3 / 3600
End Sub
Private Sub p2xc_Click()
p2x = p2x1 + p2x2 / 60 + p2x3 / 3600
End Sub
Private Sub p2yc_Click()
p2y = p2y1 + p2y2 / 60 + p2y3 / 3600
End Sub

Private Function ns(inpt As String) As String
If inpt = "S" Then ns = "-"
End Function

Function ArcSin(X As Double) As Double

Dim pi As Double
pi = CDbl(Text5)

    Dim Temp As Double
    If X = 0 Then
        Temp = 0
      ElseIf Abs(X) = 1 Then Temp = Sgn(X) * pi / 2
      Else
        Temp = Atn(X / Sqr(1 - X * X))
      End If
    ArcSin = Temp

End Function

