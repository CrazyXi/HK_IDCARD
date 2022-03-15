VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H0071CC2E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   4740
   ClientTop       =   4575
   ClientWidth     =   1335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   1335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0060AE27&
      ForeColor       =   &H8000000E&
      Height          =   300
      IMEMode         =   1  'ON
      ItemData        =   "Form1.frx":169D2
      Left            =   420
      List            =   "Form1.frx":169D4
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H0060AE27&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0060AE27&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   300
      TabIndex        =   2
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00DB9834&
      Caption         =   "生成"
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   300
      TabIndex        =   1
      Top             =   1100
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H003C4CE7&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mX As Long, mY As Long
Dim AZ As String


Private Sub Command1_Click()
End
End Sub



Private Sub Combo1_Click()
AZ = Combo1.Text

'Print (Asc(AZ) - 64)
End Sub



Private Sub Form_Load()
    
    
  
    Combo1.AddItem "A"
    Combo1.AddItem "B"
    Combo1.AddItem "C"
    Combo1.AddItem "D"
    Combo1.AddItem "E"
    Combo1.AddItem "F"
    Combo1.AddItem "G"
    Combo1.AddItem "H"
    Combo1.AddItem "I"
    Combo1.AddItem "J"
    Combo1.AddItem "K"
    Combo1.AddItem "L"
    Combo1.AddItem "M"
    Combo1.AddItem "N"
    Combo1.AddItem "O"
    Combo1.AddItem "P"
    Combo1.AddItem "Q"
    Combo1.AddItem "R"
    Combo1.AddItem "S"
    Combo1.AddItem "T"
    Combo1.AddItem "U"
    Combo1.AddItem "V"
    Combo1.AddItem "W"
    Combo1.AddItem "X"
    Combo1.AddItem "Y"
    Combo1.AddItem "Z"
    Combo1.ListIndex = 0
    AZ = "A"
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        mX = X
        mY = Y
    End If
    
   
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        Me.Move Me.Left - mX + X, Me.Top - mY + Y
    End If
End Sub '

Private Sub Label1_Click()
End
End Sub
Private Sub Label2_Click()
If (Text1.Text <> "" And Text1.Text <> " ") Then
Dim fnum As Integer, i As Integer, r As Integer, AZ As String, sum As Integer, d As Integer, f As Integer, rom As String, texts As String

AZ = Combo1.Text
fnum = CInt(Text1.Text)
    For i = 1 To fnum
    f = Asc(AZ) - 64
    d = 6
    Dim stin As Integer
    If f > 9 Then
        stin = (CInt(Left(CStr(f), 1)) * 8)
        sum = sum + stin
        stin = (CInt(Right(CStr(f), 1)) * 7)
        sum = sum + stin
    End If
    If f <= 9 Then
        stin = (CInt(Left(CStr(f), 1)) * 7)
        sum = sum + stin
    End If
        For r = 1 To 6
        Dim sjint As Integer
        Randomize
        sjint = Int(Rnd * 9)
        rom = rom + CStr(sjint)
        sum = sum + (sjint * d)
        d = d - 1
        Next r
    Dim value As String
    texts = AZ + CStr(rom)
    rom = ""
    sum = sum Mod 11
    If sum = 1 Then
        value = "A"
    End If
    If sum = 0 Then
        value = "0"
    End If
    If sum >= 2 And sum <= 10 Then
        value = CStr((11 - sum))
    End If
    sum = 0
    Text2.Text = Text2.Text + texts + "(" + value + ")" & vbCrLf
    texts = ""
    Next i
End If

Dim CurrentPath As String
    
If Right$(App.Path, 1) = "\" Then
    CurrentPath = App.Path
Else
    CurrentPath = App.Path + "\"
End If
Open CurrentPath & "idcard.txt" For Output As #1
Print #1, Text2.Text
Close #1
    
End Sub


