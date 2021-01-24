VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEstruturadorDeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'AUTOR: FABRÍCIO A. DINIZ
'DATA DA ÚLTIMA ALTERAÇÃO: 24/01/2021
'HORA DAQ ÚLTIMA ALTERAÇÃO: 13:40

'Objetivo: Estruturar os dados coletados do XML em ARRAYs
'de objetos de classes específicas:

'CLASSE ClsPais: Informações dos PAÍSES e suas LÍNGUAS FALADAS

Private sNomeDoNodoPai As String
Private sNomeDoNodoFilho As String
Private oLog As ClsLog


Public Property Get aNomeDoNodoPai() As String
 aNomeDoNodoPai = sNomeDoNodoPai
End Property

Public Property Let aNomeDoNodoPai(ByVal sValor As String)
 sNomeDoNodoPai = sValor
End Property

Public Property Get aNomeDoNodoFilho() As String
 aNomeDoNodoFilho = sNomeDoNodoFilho
End Property

Public Property Let aNomeDoNodoFilho(ByVal sValor As String)
 sNomeDoNodoFilho = sValor
End Property

'------------------------------------------------------
'Converte os dados recebidos de uma COLLECTOIN para um ARRAY de OBJETOS da classe PAÍS
'------------------------------------------------------
Public Sub EstruturarDadosDosPaises(oDados() As Collection, oPaises() As ClsPais)
 
On Error GoTo TRATAR_ERRO_ESTRUTURAR_DADOS_DOS_PAISES

    Dim i As Integer
    Dim j As Integer
    Dim ContIdioma As Integer
    Dim iTotalDeLinguasFaladas As Integer
    Dim iColunaIdioma As Integer
    
    Erase oPaises
    ReDim oPaises(UBound(oDados))
    
    'Referência: sNomeDoNodoPai = "tCountryInfo"/ sNomeDoNodoFilho = "tLanguage"
    
    
    '-------------------------------------------------------
    For i = 0 To UBound(oPaises)
        
       Set oPaises(i) = New ClsPais
        
       oPaises(i).aISOCode = oDados(i).Item(CStr(i) & "_0")
       oPaises(i).aName = oDados(i).Item(CStr(i) & "_1")
       oPaises(i).aCapitalCity = oDados(i).Item(CStr(i) & "_2")
       oPaises(i).aPhoneCode = oDados(i).Item(CStr(i) & "_3")
       oPaises(i).aContinentCode = oDados(i).Item(CStr(i) & "_4")
       oPaises(i).aCurrencyISOCode = oDados(i).Item(CStr(i) & "_5")
       oPaises(i).aCountryFlag = oDados(i).Item(CStr(i) & "_6")
       
       iTotalDeLinguasFaladas = oDados(i).Item("NodoFilho").Item("TotalDeAtributos")
       
       oPaises(i).aTotalDeLinguasFaladas = (iTotalDeLinguasFaladas / 3)
       Call oPaises(i).RedimensionarArrayDeLinguasFaladas(oPaises(i).aTotalDeLinguasFaladas - 1)
       
       
       '-------------------------------------------------------
       'Nodos filhos (Línguas faladas)
       '-------------------------------------------------------
       ContIdioma = -1
       iColunaIdioma = 0
       
       '-------------------------------------------------------
       For j = 0 To (iTotalDeLinguasFaladas - 1) 'Começa do índice 1: -> Ignorar as linhas que contém ISOCODE e NAME juntos
        
        If (j Mod 3) <> 0 Then 'Ignorar as linhas que contém ISOCODE e NAME juntos
            
            iColunaIdioma = (iColunaIdioma Mod 2) 'Coluna 0 (ISOCODE) ou 1(NAME)
            Call oPaises(i).setarLinguaFalada(ContIdioma, iColunaIdioma, oDados(i).Item("NodoFilho").Item(CStr(i) & "_" & j))
            iColunaIdioma = iColunaIdioma + 1
            
        Else
            'Linha com ISOCODE e NAME juntos... marca o início de um novo NODO de IDIOMA
            ContIdioma = ContIdioma + 1
        End If
        
       Next j 'Para cada idioma do País
       '-------------------------------------------------------
       oPaises(i).aTotalDeLinguasFaladas = (iTotalDeLinguasFaladas / 3)
       
    Next i 'Para cada País...
    '-------------------------------------------------------

TRATAR_ERRO_ESTRUTURAR_DADOS_DOS_PAISES:

  If Err.Number <> 0 Then
    oLog.gravarLog ("EstruturadorDeDados.EstruturarDadosDosPaises -> " & Err.Number & " " & Err.Description)
  Else
    oLog.gravarLog ("EstruturadorDeDados.EstruturarDadosDosPaises -> OK")
  End If
  Err.Clear

End Sub
'------------------------------------------------------

Private Sub pSetarValoresIniciais()
    Set oLog = New ClsLog
End Sub

Private Sub pSetarValoresFinais()
    Set oLog = Nothing
End Sub

Private Sub Class_Initialize()
    Call pSetarValoresIniciais
End Sub

Private Sub Class_Terminate()
    Call pSetarValoresFinais
End Sub
