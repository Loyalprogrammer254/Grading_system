VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21330
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtgrade 
      Height          =   855
      Left            =   4440
      TabIndex        =   8
      Top             =   9240
      Width           =   3255
   End
   Begin VB.TextBox txtmean 
      Height          =   975
      Left            =   4440
      TabIndex        =   7
      Top             =   7920
      Width           =   3735
   End
   Begin VB.TextBox txttotal 
      Height          =   1095
      Left            =   4440
      TabIndex        =   6
      Top             =   6360
      Width           =   3135
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Top             =   9840
      Width           =   2775
   End
   Begin VB.TextBox txtmarks5 
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   5280
      Width           =   3495
   End
   Begin VB.TextBox txtmarks4 
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txtmarks3 
      Height          =   645
      Left            =   4200
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtmarks2 
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtmarks1 
      Height          =   735
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCalculate_Click()
 Dim marks1 As Integer
 Dim marks2 As Integer
 Dim marks3 As Integer
 Dim marks4 As Integer
 Dim marks5 As Integer
 Dim total As Integer
 Dim mean As Double
 Dim grade As Double
 

marks1 = Val(txtmarks1.Text)
marks2 = Val(txtmarks2.Text)
marks3 = Val(txtmarks3.Text)
marks4 = Val(txtmarks4.Text)
marks5 = Val(txtmarks5.Text)

txttotal.Text = Val(txtmarks1.Text) + Val(txtmarks2.Text) + Val(txtmarks3.Text) + Val(txtmarks4.Text) + Val(txtmarks5.Text)
txtmean.Text = Val(txttotal.Text / 5)

If mean >= 76 Then
txtgrade.Text = "Distiction"
ElseIf mean >= 51 <= 75 Then
txtgrade.Text = "Credit"
ElseIf mean >= 26 <= 50 Then
txtgrade.Text = "Pass"
ElseIf mean >= 1 <= 25 Then
txtgrade.Text = "Fail"
Else
txtgrade.Text = "Reffer"
End If



End Sub
