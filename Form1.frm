VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Big Integer Addition"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   1950
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3510
      Width           =   8115
   End
   Begin VB.TextBox Text2 
      Height          =   3435
      Left            =   4095
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   0
      Width           =   4020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   330
      Left            =   7065
      TabIndex        =   1
      Top             =   5490
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Height          =   3435
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1035
      TabIndex        =   7
      Top             =   5535
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "IntegerSize:"
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   5535
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3780
      TabIndex        =   5
      Top             =   1575
      Width           =   285
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   7065
      TabIndex        =   2
      Top             =   7875
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prime As Double
Dim p As Double
Dim Akt As Double, Merk As Double


Private Sub Command1_Click()
Form1.MousePointer = 11
Text3.Text = Addition(Text1.Text, Text2.Text)
Label4.Caption = Len(Text3.Text)
Form1.MousePointer = 0
End Sub

Public Function Addition(Summand1 As String, Summand2 As String) As String
'z.B.
'321+482
'-------
'      3
'     0  1 gm
'    8
'-------
'    803
Dim gm As String, neu As String
Dim additionT As String



Do


additionT = Val(gm) + Val(Right$(Summand1, 1)) + Val(Right$(Summand2, 1))
gm = ""
'Merken
If Len(additionT) > 1 Then gm = Left$(CStr(additionT), 1)
'Neu
neu = Right$(CStr(additionT), 1)

Addition = neu + Addition

'Summanden kürzen


If Len(Summand1) = 1 And Len(Summand2) = 1 Then Exit Do

If Len(Summand1) > 1 Then
Summand1 = Left$(Summand1, Len(Summand1) - 1)
Else
Summand1 = 0
End If


If Len(Summand2) > 1 Then
Summand2 = Left$(Summand2, Len(Summand2) - 1)
Else
Summand2 = 0
End If



Loop

l:
Addition = gm + Addition
If Left$(Addition, 1) = "0" Then Addition = Right$(Addition, Len(Addition) - 1)



End Function

