Attribute VB_Name = "Module1"
Option Explicit
 
Public Type tyNumeroExtenso
   Numero As Integer
   Extenso As String
 End Type
 
 Public gIntNumero() As tyNumeroExtenso

Public Sub ConverteNumeroParaExtenso(pDblNumero As Double)

   Dim strExtenso As tyNumeroExtenso
   Dim vliAuxiliar As Integer

   For vliAuxiliar = 0 To UBound(gIntNumero)
      If pDblNumero = gIntNumero(vliAuxiliar).Numero Then
         MsgBox gIntNumero(vliAuxiliar).Extenso

      End If
   Next vliAuxiliar

End Sub


Public Sub iniciaTyNumeroExtenso()
   
   ReDim gIntNumero(10)
   
   gIntNumero(0).Numero = 1
   gIntNumero(0).Extenso = "Um"
   
   gIntNumero(1).Numero = 2
   gIntNumero(1).Extenso = "Dois"
   
   gIntNumero(2).Numero = 3
   gIntNumero(2).Extenso = "Tres"
   
   gIntNumero(3).Numero = 4
   gIntNumero(3).Extenso = "Quatro"
   
   gIntNumero(4).Numero = 5
   gIntNumero(4).Extenso = "Cinco"
   
   gIntNumero(5).Numero = 6
   gIntNumero(5).Extenso = "Seis"
   
   gIntNumero(6).Numero = 7
   gIntNumero(6).Extenso = "Sete"
   
   gIntNumero(7).Numero = 8
   gIntNumero(7).Extenso = "Oito"
   
   gIntNumero(8).Numero = 9
   gIntNumero(8).Extenso = "Nove"
   
   gIntNumero(9).Numero = 10
   gIntNumero(9).Extenso = "Dez"
   

End Sub
