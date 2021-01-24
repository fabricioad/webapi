VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  'AUTOR: FABRÍCIO A. DINIZ
  'DATA DA ÚLTIMA ALTERAÇÃO: 24/01/2021
  'HORA DAQ ÚLTIMA ALTERAÇÃO: 13:40
  
  
  'Objetivo: Estrutura com os atributos de um PAÍS de acordo com a API
  'que será utilizada
  
   Public aISOCode As String
   Public aName As String
   Public aCapitalCity As String
   Public aPhoneCode As String
   Public aContinentCode As String
   Public aCurrencyISOCode As String
   Public aCountryFlag As String
   Public aTotalDeLinguasFaladas As Long
   
   Private oLinguasFaladas() As ClsLinguaFalada
  

Public Property Get aTotalDeLinguas() As Integer
  On Error GoTo ArrayVazio
    Err.Clear
    
    aTotalDeLinguasFaladas = UBound(aLinguasFaladas) + 1
    
    aTotalDeLinguas = aTotalDeLinguasFaladas
    
ArrayVazio:

    If Err.Number <> 0 Then
        aTotalDeLinguasFaladas = 0
        aTotalDeLinguas = aTotalDeLinguasFaladas
    End If
    
End Property

Public Property Get aLinguasFaladas() As ClsLinguaFalada()
    
    If aTotalDeLinguasFaladas >= 0 Then
        aLinguasFaladas = oLinguasFaladas
    End If
    
    
End Property

Public Sub setarLinguaFalada(ByVal iIndice As Integer, ByVal iColuna As Integer, ByVal sValor As String)
    
    Select Case iColuna
     Case 0
        oLinguasFaladas(iIndice).aISOCode = sValor
     Case 1
        oLinguasFaladas(iIndice).aName = sValor
    End Select
    
End Sub
Public Sub RedimensionarArrayDeLinguasFaladas(ByVal iMax As Integer)
    Dim i As Integer
        
    aTotalDeLinguasFaladas = iMax
    If iMax >= 0 Then
        ReDim oLinguasFaladas(iMax)
        For i = 0 To iMax
            Set oLinguasFaladas(i) = New ClsLinguaFalada
        Next i
    End If
    
End Sub
Private Sub Class_Initialize()
    
   aISOCode = ""
   aName = ""
   aCapitalCity = ""
   aPhoneCode = ""
   aContinentCode = ""
   aCurrencyISOCode = ""
   aCountryFlag = ""
   aTotalDeLinguasFaladas = -1
   
   Erase oLinguasFaladas
   
End Sub
