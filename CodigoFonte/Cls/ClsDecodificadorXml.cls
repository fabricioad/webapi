VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDecodificadorXml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'AUTOR: FABRÍCIO A. DINIZ
'DATA DA ÚLTIMA ALTERAÇÃO: 24/01/2021
'HORA DAQ ÚLTIMA ALTERAÇÃO: 13:40

'Objetivo: Decodificar o conteúdo XML em estruturas manipuláveis de dados
'--------------------------------------------------------------------------------
'**REFERÊNCIA NECESSÁRIA** -> ..\Windows\SysWOW64\msxml3.dll#Microsoft XML, v3.0
'--------------------------------------------------------------------------------

'CONSTANTES
'------------------------

'CREATEOBJECT:
Const OBJ_MSXML2 = "MSXML2.DOMDocument"
'------------------------

'Atributos da Classe
'---------------------
Public Enum eTipoDeFiltro
 Contem = 1
 IniciaCom = 2
End Enum

Private iTotalDeElementos As Long

Private sNodo As String

Private bFiltroAtivo As Boolean
Private sNodoFiltro As String
Private sTextoDoFiltro As String
Private iTipoDeFiltro As eTipoDeFiltro

Private cDadosDoNodoPai() As Collection
Private cDadosNodoFilho As Collection

Private sNomeDoNodoPai As String
Private sNomeDoPrimeiroAtributoDoNodoPai As String

Private oLog As ClsLog
Private Nodo As Object
Private Nodos As Object
Private oXML As Object
'---------------------


'------------------------------------------------
'INÍCIO: ATRIBUTOS DA CLASSE
'------------------------------------------------
Public Property Get aTotalDeElementos() As Long
 aTotalDeElementos = iTotalDeElementos
End Property

Public Property Get aFiltroAtivo() As Boolean
 aFiltroAtivo = bFiltroAtivo
End Property

Public Property Let aFiltroAtivo(ByVal bValor As Boolean)
 bFiltroAtivo = bValor
End Property

Public Property Get aTipoDeFiltro() As eTipoDeFiltro
 aTipoDeFiltro = iTipoDeFiltro
End Property

Public Property Let aTipoDeFiltro(ByVal eValor As eTipoDeFiltro)
 iTipoDeFiltro = eValor
End Property

Public Property Get aNomeDoNodoPai() As String
 aNomeDoNodoPai = sNomeDoNodoPai
End Property

Public Property Let aNomeDoNodoPai(ByVal sValor As String)
 sNomeDoNodoPai = sValor
End Property

Public Property Get aNomeDoPrimeiroAtributoDoNodoPai() As String
 aNomeDoPrimeiroAtributoDoNodoPai = sNomeDoPrimeiroAtributoDoNodoPai
End Property

Public Property Let aNomeDoPrimeiroAtributoDoNodoPai(ByVal sValor As String)
 sNomeDoPrimeiroAtributoDoNodoPai = sValor
End Property

Public Property Get aNodoFiltro() As String
 aNodoFiltro = sNodoFiltro
End Property

Public Property Let aNodoFiltro(ByVal sValor As String)
 sNodoFiltro = sValor
End Property

Public Property Get aTextoDoFiltro() As String
 aTextoDoFiltro = sTextoDoFiltro
End Property

Public Property Let aTextoDoFiltro(ByVal sValor As String)
 sTextoDoFiltro = sValor
End Property


'Seta os nomes dos atributos vinculados àquele Nodo
Public Property Get aDadosDoNodoPai() As Collection()
 aDadosDoNodoPai = cDadosDoNodoPai
End Property
'------------------------------------------------
'FIM: ATRIBUTOS DA CLASSE
'------------------------------------------------


Private Sub estruturarDadosDoNodo(ByVal j As Integer, ByVal i As Integer, ByVal iContNodo As Integer)
            
            'Grava dado sobre os nodos já lidos até aqui
            '------------------------------
            If j = 0 Then 'Não havia Language?
              Set cDadosNodoFilho = New Collection 'Seta novo objeto inicial...
            End If
            
            Call cDadosNodoFilho.Add(CStr(j), "TotalDeAtributos")
            Call cDadosDoNodoPai(iContNodo).Add(cDadosNodoFilho, "NodoFilho")
            Call cDadosDoNodoPai(iContNodo).Add(CStr(i), "TotalDeAtributos")
            
            'Prepara para o próximo
            Set cDadosNodoFilho = Nothing
            '------------------------------
            
End Sub
'------------------------------------------------
'INÍCIO: MÉTODOS DA CLASSE
'------------------------------------------------
Public Sub mDecodificar()
     
On Error GoTo TRATAR_ERRO_DECODIFICAR

    Dim i As Integer
    Dim j As Integer
    
    Dim iContNodo As Integer
    
    Dim bEstaDentroDoFiltro As Boolean
    Dim sSelecaoDeNodosXML As String
    Dim sFiltroSelecaoDeNodosXML As String
        
    Dim Nodo As MSXML2.IXMLDOMNode
        
    Call LimparCollection
    
    
    'Seleciona com XPATH (XML)
    '-----------
    oXML.setProperty "SelectionLanguage", "XPath"
    '-----------
            
    'Possui NodoFilho?
    sSelecaoDeNodosXML = "//" & sNomeDoNodoPai
            
    'Seleciona os Nodos do XML de acordo os parâmetros configurados
    '------------------------------------
    If bFiltroAtivo And Len(sNodoFiltro) <> 0 Then
        
        If iTipoDeFiltro = IniciaCom Then 'Filtro: atributo iniciado por xxxx...
            sFiltroSelecaoDeNodosXML = "[" & sNodoFiltro & "[starts-with(.,'" & sTextoDoFiltro & "')]]"
        Else 'Filtro padrão: contém xxxx...
            sFiltroSelecaoDeNodosXML = "[" & sNodoFiltro & "[contains(.,'" & sTextoDoFiltro & "')]]"
        End If
        
    Else 'Sem filtro
        sFiltroSelecaoDeNodosXML = "" 'Retorna todos nodos
    End If
    '------------------------------------
    Set Nodos = oXML.selectNodes(sSelecaoDeNodosXML & sFiltroSelecaoDeNodosXML & "//*")
    
    Erase cDadosDoNodoPai
    iContNodo = -1
    Set cDadosNodoFilho = Nothing
    '-------------------------------------------------------------

    'Para cada nodo no na estrutura do XML
    For Each Nodo In Nodos
        
        'É um NODO PAI?
        If Nodo.parentNode.nodeName = sNomeDoNodoPai And Nodo.nodeName = sNomeDoPrimeiroAtributoDoNodoPai Then
         
         If (iContNodo >= 0) Then 'Não é o primeiro registro?
            
            'Grava dado sobre os nodos já lidos até aqui
            '------------------------------
            Call estruturarDadosDoNodo(j, i, iContNodo) 'Total de línguas faladas e total de atributos por país consultado
            '------------------------------
            
         End If
         
         'Contabiliza um NOVO NODO PAI
         iContNodo = iContNodo + 1
         ReDim Preserve cDadosDoNodoPai(iContNodo)
         Set cDadosDoNodoPai(iContNodo) = New Collection
         
         i = 0
         j = 0
         
        End If
        
        'Nodo filho correto?
        If (Nodo.parentNode.nodeName <> sNomeDoNodoPai) Then
        
            'Fim do nodo filho
            If j = 0 Then
                Set cDadosNodoFilho = New Collection
            End If
            
            Call cDadosNodoFilho.Add(Nodo.Text, CStr(iContNodo) & "_" & CStr(j)) 'j & "_" & Nodo.parentNode.nodeName & "." & Nodo.nodeName
            j = j + 1
            
        Else
        
            'Nodo filho
            Call cDadosDoNodoPai(iContNodo).Add(Nodo.Text, CStr(iContNodo) & "_" & CStr(i))
            i = i + 1
            
        End If
        
        
        
    Next Nodo
    
    '-------------------------------------------------------------
    'Grava dados sobre o último nodo lido
    '------------------------------
    Call estruturarDadosDoNodo(j, i, iContNodo) 'Total de línguas faladas e total de atributos por país consultado
    '------------------------------
    
    iTotalDeElementos = iContNodo
    
TRATAR_ERRO_DECODIFICAR:
    
  If Err.Number <> 0 Then
    oLog.gravarLog ("DecodificarXml.mDecodificar -> " & Err.Number & " " & Err.Description)
  Else
    oLog.gravarLog ("DecodificarXml.mDecodificar -> OK")
  End If
  Err.Clear
    
End Sub

Private Function LimparCollection()

    Erase cDadosDoNodoPai
        
End Function

Public Function mCarregarXML(ByVal sXml As String) As String
        
    Dim sRetorno As String
    '-----------------------------
    sRetorno = "OK"
    Call pInstanciarObjMSXML2
    '-----------------------------
        
    '-----------------------------
    'Carregar XML do parâmetro string recebido
    '-----------------------------
    Call oXML.loadXML(sXml)
    '-----------------------------
    
    'Carregou com sucesso?
    '-----------------------------
    If oXML.parseError.reason <> "" Then
       'Não carregou com sucesso....
       sRetorno = oXML.parseError.reason 'Retorna o erro encontrado
    Else
        Call mDecodificar 'Extrai os dados do XML para uma estrutura de dados
    End If
    '-----------------------------
   
    mCarregarXML = sRetorno
End Function

'------------------------------------------------
'FIM: MÉTODOS DA CLASSE
'------------------------------------------------

Private Sub pInstanciarObjMSXML2()
  
  If oXML Is Nothing Then
    Set oXML = CreateObject(OBJ_MSXML2)
  End If
    
End Sub
Private Sub pLiberarInstanciaMSXML2()
    Set oXML = Nothing
End Sub

Private Sub pSetarValoresIniciais()
    
    sNomeDoNodoPai = ""
    sNomeDoPrimeiroAtributoDoNodoPai = ""
    Set oLog = New ClsLog
    
    Call LimparCollection
        
    iTotalDeElementos = -1
    Call pLiberarInstanciaMSXML2
    
End Sub

Private Sub pSetarValoresFinais()
    Set oLog = Nothing
    Call pLiberarInstanciaMSXML2
End Sub

Private Sub Class_Initialize()
    Call pSetarValoresIniciais
End Sub

Private Sub Class_Terminate()
    Call pSetarValoresFinais
End Sub
