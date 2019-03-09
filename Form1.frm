VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cifrario di Atbash"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cosa è il cifrario di Atbash"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cripta/Decripta"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Antaryl"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim pass, lunghezza, i, j, x, b, sar
pass = Text1.Text
lunghezza = Len(pass)
Text2.Text = ""
passx = LCase(pass)
For i = 1 To lunghezza
j = i
x = Mid(passx, j, 1)
j = j + 1
If x = "a" Then
b = "z"
ElseIf x = "b" Then
b = "y"
ElseIf x = "c" Then
b = "x"
ElseIf x = "d" Then
b = "w"
ElseIf x = "e" Then
b = "v"
ElseIf x = "f" Then
b = "u"
ElseIf x = "g" Then
b = "t"
ElseIf x = "h" Then
b = "s"
ElseIf x = "i" Then
b = "r"
ElseIf x = "j" Then
b = "q"
ElseIf x = "k" Then
b = "p"
ElseIf x = "l" Then
b = "o"
ElseIf x = "m" Then
b = "n"
ElseIf x = "n" Then
b = "m"
ElseIf x = "o" Then
b = "l"
ElseIf x = "p" Then
b = "k"
ElseIf x = "q" Then
b = "j"
ElseIf x = "r" Then
b = "i"
ElseIf x = "s" Then
b = "h"
ElseIf x = "t" Then
b = "g"
ElseIf x = "u" Then
b = "f"
ElseIf x = "v" Then
b = "e"
ElseIf x = "w" Then
b = "d"
ElseIf x = "x" Then
b = "c"
ElseIf x = "y" Then
b = "b"
ElseIf x = "z" Then
b = "a"
ElseIf x = "0" Then
b = "9"
ElseIf x = "1" Then
b = "8"
ElseIf x = "2" Then
b = "7"
ElseIf x = "3" Then
b = "6"
ElseIf x = "4" Then
b = "5"
ElseIf x = "5" Then
b = "4"
ElseIf x = "6" Then
b = "3"
ElseIf x = "7" Then
b = "2"
ElseIf x = "8" Then
b = "1"
ElseIf x = "9" Then
b = "0"
Else
MsgBox "Inserisci soltanto caratteri e/o numeri" & Chr(13) & "non caratteri speciali: |\!£$%&/()=?^é*ç°§_-ùàò+èì'[]@#;,:.", vbCritical, "Immissione dati errata"
Text1.Text = ""
Text2.Text = ""
End If
sar = b
Text2.Text = Text2.Text & sar
Next i
'togliendo gli apici dalle quattro righe sotto il programma creerà una cartella chiamata pass e salverà in pass.txt la password criptata
'MkDir (App.Path & "\pass")
'Open App.Path & "\pass\pass.txt" For Output As #1
'Print #1, Text2.Text
'Close #1
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

