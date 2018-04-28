Attribute VB_Name = "Module1"
Option Explicit
 
Public Type tyNumeroExtenso
   Numero As Integer
   Extenso As String
End Type

Private Const CO_MAPA_NUMEROS = 20
 
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

Private Function RetornarUnidade(pDblNumero As Integer) As String

   For vliAuxiliar = 0 To UBound(gIntNumero)
      If pDblNumero = gIntNumero(vliAuxiliar).Numero Then
         MsgBox gIntNumero(vliAuxiliar).Extenso
      End If
   Next vliAuxiliar

End Function

Public Sub iniciaTyNumeroExtenso()

   ReDim gIntNumero(20)
   
   gIntNumero(0).Numero = 0
   gIntNumero(0).Extenso = "Zero"
   
   gIntNumero(1).Numero = 1
   gIntNumero(1).Extenso = "Um"

   gIntNumero(1).Numero = 1
   gIntNumero(1).Extenso = "Um"

   gIntNumero(2).Numero = 2
   gIntNumero(2).Extenso = "Dois"

   gIntNumero(3).Numero = 3
   gIntNumero(3).Extenso = "Tres"

   gIntNumero(4).Numero = 4
   gIntNumero(4).Extenso = "Quatro"

   gIntNumero(5).Numero = 5
   gIntNumero(5).Extenso = "Cinco"

   gIntNumero(6).Numero = 6
   gIntNumero(6).Extenso = "Seis"

   gIntNumero(7).Numero = 7
   gIntNumero(7).Extenso = "Sete"

   gIntNumero(8).Numero = 8
   gIntNumero(8).Extenso = "Oito"

   gIntNumero(9).Numero = 9
   gIntNumero(9).Extenso = "Nove"

   gIntNumero(10).Numero = 10
   gIntNumero(10).Extenso = "Dez"

   gIntNumero(11).Numero = 11
   gIntNumero(11).Extenso = "Onze"

   gIntNumero(12).Numero = 12
   gIntNumero(12).Extenso = "Doze"
   
   gIntNumero(13).Numero = 13
   gIntNumero(13).Extenso = "Treze"
   
   gIntNumero(14).Numero = 14
   gIntNumero(14).Extenso = "Quatorze"
   
   gIntNumero(15).Numero = 15
   gIntNumero(15).Extenso = "Quinze"
   
   gIntNumero(16).Numero = 16
   gIntNumero(16).Extenso = "Dezesseis"
   
   gIntNumero(17).Numero = 17
   gIntNumero(17).Extenso = "Dezesete"
   
   gIntNumero(18).Numero = 18
   gIntNumero(18).Extenso = "Dezoito"
   
   gIntNumero(19).Numero = 19
   gIntNumero(19).Extenso = "Dezenove"
   
   gIntNumero(20).Numero = 20
   gIntNumero(20).Extenso = "20"
   
   gIntNumero(21).Numero = 30
   gIntNumero(21).Extenso = "Trinta"
   
   gIntNumero(22).Numero = 40
   gIntNumero(22).Extenso = "Quarenta"
   
   gIntNumero(23).Numero = 50
   gIntNumero(23).Extenso = "Cinquenta"
   
   gIntNumero(24).Numero = 60
   gIntNumero(24).Extenso = "Sessenta"
   
   gIntNumero(25).Numero = 70
   gIntNumero(25).Extenso = "Setenta"
   
   gIntNumero(26).Numero = 80
   gIntNumero(26).Extenso = "Oitenta"
   
   gIntNumero(27).Numero = 90
   gIntNumero(27).Extenso = "Noventa"
   
   gIntNumero(28).Numero = 100
   gIntNumero(28).Extenso = "Cem"
   
   gIntNumero(29).Numero = 200
   gIntNumero(29).Extenso = "Duzentos"
   
   gIntNumero(30).Numero = 300
   gIntNumero(30).Extenso = "Trezentos"
   
   gIntNumero(31).Numero = 400
   gIntNumero(31).Extenso = "Quatrocentos"
   
   gIntNumero(32).Numero = 500
   gIntNumero(32).Extenso = "Quinhentos"
   
   gIntNumero(33).Numero = 600
   gIntNumero(33).Extenso = "Seiscentos"
   
   gIntNumero(34).Numero = 700
   gIntNumero(34).Extenso = "Setecentos"
   
End Sub
