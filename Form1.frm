VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   705
   End
   Begin VB.TextBox txtNumero 
      Height          =   315
      Left            =   750
      TabIndex        =   0
      Top             =   390
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim dblNumero As Double
   iniciaTyNumeroExtenso
   
   dblNumero = InputBox("Digite o n�mero:")
   ConverteNumeroParaExtenso dblNumero
   
End Sub
