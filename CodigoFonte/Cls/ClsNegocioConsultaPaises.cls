VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNegocioConsultaPaises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'AUTOR: FABRÍCIO A. DINIZ
'DATA DA ÚLTIMA ALTERAÇÃO: 24/01/2021
'HORA DAQ ÚLTIMA ALTERAÇÃO: 13:40

'------------------------------------
'PRINCIPAL classe da CAMADA DE NEGÓCIO
'------------------------------------
'Objetivo: Invocar a OPERAÇÃO de uma API de uma determinada URL,
'exibir e manipular os dados retornados, gravar os dados em um Banco De Dados.
'------------------------------------

'FUNÇÃO DA API: CONSULTAR INFORMAÇÕES DOS PAÍSES E SEUS IDIOMAS FALADOS
'------------------------------------

'-------------------------------------------
Private oWebAPI As ClsWebAPI
Private oDecodificadorXML As ClsDecodificadorXml 'Decodifica e Analisa os dados de uma String no padrão XML
Private oEstruturadorDeDados As ClsEstruturadorDeDados
Private oPaises() As ClsPais
Private oLog As ClsLog
Private oBancoDeDados As ClsBancoDeDados

'Atributos do PAIS
'---------------------
Public Enum eAtributosDoPais
  aCapitalCity = 0
  aContinentCode = 1
  aCountryFlag = 2
  aCurrencyISOCode = 3
  aISOCode = 4
  aName = 5
  aPhoneCode = 6
  aTotalDeLinguasFaladas = 7
End Enum
'-------------------------------------------


Public Property Get aWebAPI() As ClsWebAPI
 Set aWebAPI = oWebAPI
End Property

Public Property Get aDecodificadorXML() As ClsDecodificadorXml
 Set aDecodificadorXML = oDecodificadorXML
End Property

Public Property Get aBancoDeDados() As ClsBancoDeDados
 Set aBancoDeDados = oBancoDeDados
End Property

Public Property Get aLog() As ClsLog
 Set aLog = oLog
End Property

Public Property Let aNomeDoNodoPai(ByVal objValor As ClsDecodificadorXml)
 Set oDecodificadorXML = objValor
End Property

Public Property Get aPaises() As ClsPais()
 aPaises = oPaises
End Property

Public Property Get aTotalDePaises() As Integer
  On Error GoTo ArrayVazio
    Err.Clear
    aTotalDePaises = UBound(oPaises) + 1
    
ArrayVazio:
    If Err.Number <> 0 Then
        aTotalDePaises = 0
    End If
    
End Property
'Faz a leitura dos dados para ACESSO ao BD no arquivo INI e tenta conectar ao BD
Public Sub bdConectar()

    Call oBancoDeDados.mFazerLeituraDoArquivoIniAcessoBD
    Call oBancoDeDados.mConectar
    
End Sub

'Percorre todos os países atuais e insere no Banco de dados
Public Function bdInserirDadosDosPaises() As Integer
    
    Dim sNomeDosCampos() As String
    Dim sTipoDosCampos() As String
    Dim sValoresDosCampos() As String
    Dim sValores As String
    
    Dim i As Integer
    Dim iRetorno As Integer
    Dim iRegistroExistente As Integer
        
        
    iRetorno = 0
    
    'DADOS FORMATADOS PARA SEREM INSERIDOS NO BANCO DE DADOS
    '---------------------------------
    sNomeDosCampos = Split("CapitalCity|ContinentCode|CountryFlag|CurrencyISOCode|IsoCode|Name|PhoneCode", "|")
    sTipoDosCampos = Split("STR|STR|STR|STR|STR|STR|INT", "|")
    For i = 0 To UBound(oPaises)
        
        sValores = _
        oPaises(i).aCapitalCity & "|" & _
        oPaises(i).aContinentCode & "|" & _
        oPaises(i).aCountryFlag & "|" & _
        oPaises(i).aCurrencyISOCode & "|" & _
        oPaises(i).aISOCode & "|" & _
        oPaises(i).aName & "|" & _
        oPaises(i).aPhoneCode
        
        sValoresDosCampos = Split(sValores, "|")
        
        'Possui ao menos IDOCODE e NAME?
        If Len(Trim(oPaises(i).aISOCode)) <> 0 And Len(Trim(oPaises(i).aName)) <> 0 Then
        
            'Verifica se o País já existe no BD
            '---------------------------------------------------
            iRegistroExistente = oBancoDeDados.mExecutarSQL("" & _
            "SELECT idPais from PAIS Where ISOCode='" & oPaises(i).aISOCode & "' AND Name='" & oPaises(i).aName & "'")
            '---------------------------------------------------
            
            'País ainda não cadastrado?
            '---------------------------------------------------
            If iRegistroExistente = 0 Then
                'Tenta inserir e contabiliza inserção com sucesso
                iRetorno = iRetorno + oBancoDeDados.mInserirDados("PAIS", sNomeDosCampos, sTipoDosCampos, sValoresDosCampos)
            End If
            '---------------------------------------------------
        
        Else
            Call oLog.gravarLog("INSERT DO PAIS IGNORADO: ISOCode -> " & oPaises(i).aISOCode & " | " & oPaises(i).aName)
        End If
        
    Next i
    
    '---------------------------------
    
    bdInserirDadosDosPaises = iRetorno 'Retorna total de registros inseridos com sucesso
    
End Function

'Percorre todos os países atuais e insere no Banco de dados
Public Function bdInserirDadosDosIdiomas() As Integer
    
    Dim sNomeDosCampos() As String
    Dim sTipoDosCampos() As String
    Dim sValoresDosCampos() As String
    Dim sValores As String
    
    Dim i As Integer
    Dim j As Integer
    Dim iRetorno As Integer
    Dim iRegistroExistente As Integer
        
        
    iRetorno = 0
    
    'DADOS FORMATADOS PARA SEREM INSERIDOS NO BANCO DE DADOS
    '---------------------------------
    sNomeDosCampos = Split("IsoCode|Name", "|")
    sTipoDosCampos = Split("STR|STR", "|")
    
    For i = 0 To UBound(oPaises) 'Para cada país
                
        'O PAÍS POSSUI UM IDIOMA VINCULADO?
        '--------------------------------------------------------------------------
        If oPaises(i).aTotalDeLinguasFaladas > 0 Then
        
                For j = 0 To (oPaises(i).aTotalDeLinguasFaladas - 1) 'Para cada idioma deste país
                    
                    sValores = "" & _
                    oPaises(i).aLinguasFaladas()(j).aISOCode & "|" & _
                    oPaises(i).aLinguasFaladas()(j).aName
                    
                    'Possui ISOCODE e NAME?
                    If Len(Trim(oPaises(i).aLinguasFaladas()(j).aISOCode)) <> 0 And Len(Trim(oPaises(i).aLinguasFaladas()(j).aName)) <> 0 Then
                        
                        sValoresDosCampos = Split(sValores, "|")
                    
                        'Verifica se o IDIOMA já existe no BD
                        '---------------------------------------------------
                        iRegistroExistente = oBancoDeDados.mExecutarSQL("" & _
                        "SELECT idIdioma from IDIOMA Where ISOCode='" & oPaises(i).aLinguasFaladas()(j).aISOCode & "' AND Name='" & oPaises(i).aLinguasFaladas()(j).aName & "'")
                        '---------------------------------------------------
                    
                        'IDIOMA ainda não cadastrado?
                        '---------------------------------------------------
                        If iRegistroExistente = 0 Then
                            'Tenta inserir e contabiliza inserção com sucesso
                            iRetorno = iRetorno + oBancoDeDados.mInserirDados("IDIOMA", sNomeDosCampos, sTipoDosCampos, sValoresDosCampos)
                        End If
                        '---------------------------------------------------
                    
                    Else
                        Call oLog.gravarLog("INSERT DO IDIOMA IGNORADO: ISOCode -> " & oPaises(i).aLinguasFaladas()(j).aISOCode & " | " & oPaises(i).aLinguasFaladas()(j).aName)
                    End If
                    
                Next j 'próximo idioma do país atual
        
        Else 'PAÍS SEM IDIOMA VINCULADO NO XML (API)
            Call oLog.gravarLog("** IDIOMA DO PAIS NÃO ENCONTRADO **: ISOCode -> " & oPaises(i).aISOCode & " | " & oPaises(i).aName)
        End If
        '--------------------------------------------------------------------------
        
    Next i 'próximo país
    
    '---------------------------------
    
    bdInserirDadosDosIdiomas = iRetorno 'Retorna total de registros inseridos com sucesso
    
End Function


'Percorre todos os países atuais e insere no Banco de dados
Public Function bdInserirDadosDosIdiomasDoPaises() As Integer
    
    Dim sNomeDosCampos() As String
    Dim sTipoDosCampos() As String
    Dim sValoresDosCampos() As String
    Dim sValores As String
    
    Dim i As Integer
    Dim j As Integer
    Dim iRetorno As Integer
    Dim iRegistroExistente As Integer
    
    Dim iIdPais As Integer
    Dim iIdIdioma As Integer
    Dim bIdiomaVinculadoAoPais As Boolean
        
    iRetorno = 0
    
    'DADOS FORMATADOS PARA SEREM INSERIDOS NO BANCO DE DADOS
    '---------------------------------
    sNomeDosCampos = Split("IdPais|IdIdioma", "|")
    sTipoDosCampos = Split("INT|INT", "|")
    
    For i = 0 To UBound(oPaises) 'Para cada país
        
        iIdPais = 0 'País ainda não caastrado
        
        'ID do PAÍS cadastrado
        '---------------------------------------------------
        iRegistroExistente = oBancoDeDados.mExecutarSQL("" & _
        "SELECT idPais from PAIS Where ISOCode='" & oPaises(i).aISOCode & "' AND Name='" & oPaises(i).aName & "'")
        
        'Id do PAÍS retornado?
        '---------------------------------------------------
        If iRegistroExistente > 0 Then
                
                iIdPais = oBancoDeDados.aRecordset.fields(0).Value 'ID no banco de dados do PAÍS ATUAL no loop
        
                For j = 0 To (oPaises(i).aTotalDeLinguasFaladas - 1) 'Para cada idioma deste país
                        
                    iIdIdioma = 0 'Idioma ainda não cadastrado
                    
                    'ID do IDIOMA cadastrado
                    '---------------------------------------------------
                    iRegistroExistente = oBancoDeDados.mExecutarSQL("" & _
                    "SELECT idIdioma from IDIOMA Where ISOCode='" & oPaises(i).aLinguasFaladas()(j).aISOCode & "' AND Name='" & oPaises(i).aLinguasFaladas()(j).aName & "'")
                    
                    'Id do IDIOMA retornado?
                    '---------------------------------------------------
                    If iRegistroExistente > 0 Then
                        
                        iIdIdioma = oBancoDeDados.aRecordset.fields(0).Value 'ID no banco de dados do IDIOMA ATUAL no loop
                        
                        'Registro da TABELA idioma_do_pais
                        '-------------------
                        sValores = iIdPais & "|" & iIdIdioma
                        sValoresDosCampos = Split(sValores, "|")
                        '-------------------
                        
                        iRegistroExistente = oBancoDeDados.mExecutarSQL("" & _
                        "SELECT idIdiomas_do_Pais from IDIOMAS_DO_PAIS Where idPais=" & iIdPais & " AND idIdioma=" & iIdIdioma)
                        
                        'IDIOMA DO PAÍS não vinculado?
                        '---------------------------------------------------
                        If iRegistroExistente = 0 Then
                            'Tenta inserir e contabiliza inserção com sucesso
                            iRetorno = iRetorno + oBancoDeDados.mInserirDados("IDIOMAS_DO_PAIS", sNomeDosCampos, sTipoDosCampos, sValoresDosCampos)
                        End If
                        '---------------------------------------------------
                    
                    End If 'Retornou ID do Idioma?
                    '---------------------------------------------------
                    
                Next j 'próximo idioma do país atual
        
        End If 'Retornou um ID do PAÍS?
        '---------------------------------------------------
        
        
    Next i 'próximo país
    
    '---------------------------------
    
    bdInserirDadosDosIdiomasDoPaises = iRetorno 'Retorna total de registros inseridos com sucesso
    
End Function



'Retorna uma lista de ATRIBUTOS separada por quebra de linha como padrão
Public Function mObterListaDosPaises(ByVal eAtributo As eAtributosDoPais, Optional sSeparador As String = vbCrLf) As String
    
    Dim i As Integer
    Dim sRetorno As String
    
    sRetorno = ""
    
    '-------------------------------------------------------
    For i = 0 To UBound(oPaises)
    
     Select Case eAtributo
      
      Case 0
       sRetorno = sRetorno & oPaises(i).aCapitalCity & sSeparador
      Case 1
       sRetorno = sRetorno & oPaises(i).aContinentCode & sSeparador
      Case 2
       sRetorno = sRetorno & oPaises(i).aCountryFlag & sSeparador
      Case 3
       sRetorno = sRetorno & oPaises(i).aCurrencyISOCode & sSeparador
      Case 4
       sRetorno = sRetorno & oPaises(i).aISOCode & sSeparador
      Case 5
       sRetorno = sRetorno & oPaises(i).aName & sSeparador
      Case 6
       sRetorno = sRetorno & oPaises(i).aPhoneCode & sSeparador
      Case 7
       sRetorno = sRetorno & oPaises(i).aTotalDeLinguasFaladas & sSeparador
       
     End Select
     
    Next i
    '-------------------------------------------------------
    
    'Retorna a lista do atributo de cada País da estrutura ATUAL
    'separados por um caractere SEPARADOR -> sSeparador
    '-------------------------------------------------------
    mObterListaDosPaises = sRetorno

End Function

Private Sub gravarLOG_InvocarAPI()

  Call oLog.gravarLog("Invocar API: " & oWebAPI.aUrl)
  Call oLog.gravarLog("Operação: " & oWebAPI.aOperacao)
  Call oLog.gravarLog("NodoFiltro: " & oDecodificadorXML.aNodoFiltro)
  Call oLog.gravarLog("TipoDeFiltro: " & oDecodificadorXML.aTipoDeFiltro)
  Call oLog.gravarLog("TextoDoFiltro: " & oDecodificadorXML.aTextoDoFiltro)
  Call oLog.gravarLog("FiltroAtivo: " & oDecodificadorXML.aFiltroAtivo)
  
End Sub

Private Sub gravarLOG_StatusWebAPI()

  Call oLog.gravarLog("Status WebAPI: " & oWebAPI.aStatusCode & " - " & oWebAPI.aStatusText)
  
End Sub
Public Function mInvocarAPI() As Integer
    
  Dim sConteudoXML  As String
  Dim iRespostaInvocarAPI As Integer
  
  '--------------------------------------------------
  'Invoca a API da OPERAÇÃO na URL especificada
  '--------------------------------------------------
  'Exemplo:
  'oWebAPI.aUrl = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso"
  'oWebAPI.aOperacao = "FullCountryInfoAllCountries"
  '--------------------------------------------------
  
  Call gravarLOG_InvocarAPI 'Grava o LOG
  
  'OBTEVE SUCESSO AO INVOCAR A API?
  '-----------------------------------------
  iRespostaInvocarAPI = oWebAPI.mInvocar
  
  If iRespostaInvocarAPI = 1 Then
  
    'Resposta obtida
    sConteudoXML = oWebAPI.aResposta
    Call gravarLOG_StatusWebAPI 'Grava o LOG
    '----------------------
    
    'Extrai os dados do XML para uma estrutura de dados
    '----------------------
    Call oDecodificadorXML.mCarregarXML(sConteudoXML)
    '----------------------
    
    'Estrutura os Dados retornados pelo XML de API de CONSULTA PAÍSES
    '----------------------
    oEstruturadorDeDados.aNomeDoNodoPai = "tCountryInfo"
    oEstruturadorDeDados.aNomeDoNodoFilho = "tLanguage"
    Call oEstruturadorDeDados.EstruturarDadosDosPaises(oDecodificadorXML.aDadosDoNodoPai, oPaises)
    '----------------------
   
   End If
   '-----------------------------------------
   
   'Retorna se obteve sucesso (1) ou não(0) em invocar a API
   mInvocarAPI = iRespostaInvocarAPI

End Function

Private Sub pSetarValoresIniciais()
    
    Set oBancoDeDados = New ClsBancoDeDados
    Set oWebAPI = New ClsWebAPI
    Set oDecodificadorXML = New ClsDecodificadorXml
    Set oEstruturadorDeDados = New ClsEstruturadorDeDados
    Set oLog = New ClsLog
    
End Sub

Private Sub pSetarValoresFinais()
    
    Set oBancoDeDados = Nothing
    Set oWebAPI = Nothing
    Set oDecodificadorXML = Nothing
    Set oEstruturadorDeDados = Nothing
    Set oLog = Nothing
    
End Sub

Private Sub Class_Initialize()
    Call pSetarValoresIniciais
End Sub

Private Sub Class_Terminate()
    Call pSetarValoresFinais
End Sub
