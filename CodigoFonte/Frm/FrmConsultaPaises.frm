VERSION 5.00
Begin VB.Form FrmConsultaPaises 
   BackColor       =   &H00404040&
   Caption         =   "WEBAPI - Consulta a pa�ses"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9225
   Icon            =   "FrmConsultaPaises.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameDadosAcao 
      BackColor       =   &H00404040&
      Caption         =   "Dados"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   8775
      Begin VB.CommandButton BtnSalvarDados 
         Appearance      =   0  'Flat
         Caption         =   "&Salvar"
         Height          =   615
         Left            =   6000
         MouseIcon       =   "FrmConsultaPaises.frx":0ECA
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton BtnBaixarDados 
         Appearance      =   0  'Flat
         Caption         =   "&Baixar"
         Height          =   615
         Left            =   3240
         MouseIcon       =   "FrmConsultaPaises.frx":101C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame frameDadosRetornados 
      BackColor       =   &H00404040&
      Caption         =   "Dados retornados"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox TxtDados 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   8535
      End
   End
End
Attribute VB_Name = "FrmConsultaPaises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
'AUTOR: FABR�CIO A. DINIZ
'DATA DA �LTIMA ALTERA��O: 24/01/2021
'HORA DAQ �LTIMA ALTERA��O: 13:40

'CAMADA DE APRESENTA��O
'OBJETIVO: Consumir os dados de uma WebAPI,exib�-los e armazen�-los no banco de dados MYSQL
'---------------------------------------------------------------

Dim objConsultaPaises As New ClsNegocioConsultaPaises
Dim objLog As New ClsLog
Dim sDadosRetornados As String

'Seta as configura��es da API que ser� utilizada
'----------------------------------------------
Private Sub setarAPIConsultaPaises()

    'Seta a URL da API que ser� chamada e qual a OPERA��O
    '------------------------------------
    objConsultaPaises.aWebAPI.aUrl = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso"
    objConsultaPaises.aWebAPI.aOperacao = "FullCountryInfoAllCountries"
    '------------------------------------
    
    'Prepara o DecodificadorXML com Nomes e Filtros
    '----------------------
    'NODO PAI
    objConsultaPaises.aDecodificadorXML.aNomeDoNodoPai = "tCountryInfo"
    objConsultaPaises.aDecodificadorXML.aNomeDoPrimeiroAtributoDoNodoPai = "sISOCode"
    '-------------------------------------------
        
    'Filtro
    '-------------------------------------------
    objConsultaPaises.aDecodificadorXML.aNodoFiltro = "sISOCode"  'Nome do Atributo -> sISOCode, sName, sCapitalCity, sPhoneCode, sContinentCode, sCurrencyISOCode, sCountryFlag
    objConsultaPaises.aDecodificadorXML.aTipoDeFiltro = IniciaCom
    objConsultaPaises.aDecodificadorXML.aTextoDoFiltro = "A"
    objConsultaPaises.aDecodificadorXML.aFiltroAtivo = True 'Ativa o filtro ao extrair os dados do XML
    '-------------------------------------------
    
End Sub

Private Sub pPermitirSalvarDados()

    'Retornou algum dado?
    If Len(sDadosRetornados) <> 0 Then
     'Permite tentar gravar no banco de dados
     BtnSalvarDados.Enabled = True
    Else
     'N�o permite tentar gravar no banco de dados
     BtnSalvarDados.Enabled = False
    End If
    
End Sub

'----------------------------------------------
'INVOCA A API configurada via URL e OPERA��O
'-----------------------------------------------------------
Private Sub pInvocarAPI()
    
    Dim iRespostaInvarApi As Integer
    'Seta as configura��es da API que ser� utilizada
    '----------------------------------------------
    Call setarAPIConsultaPaises

    'INVOCA a API (retorna uma STRING com o conte�do em XML)
    '-------------------------------------------
    iRespostaInvarApi = objConsultaPaises.mInvocarAPI
    
    If iRespostaInvarApi = 1 Then 'OBTEVE SUCESSO AO INVOCAR A API?
    
        'Retorna a lista do ATRIBUTO (Name) desejado dos Pa�ses filtrados
        '-------------------------------------------
        sDadosRetornados = objConsultaPaises.mObterListaDosPaises(aName)
        TxtDados.Text = sDadosRetornados
        
        'Retorna mensagem para o usu�rio
        '--------------------------------------------
        MsgBox "A opera��o da WebAPI invocada retornou:" & vbCrLf & vbCrLf & _
        "Total de pa�ses: " & objConsultaPaises.aTotalDePaises, vbInformation, "WebApi"
        '--------------------------------------------
        
        Call pPermitirSalvarDados 'Se retornou dados, poder� tentar salvar no Banco de Dados
    
    Else 'N�O OBTEVE SUCESSO EM INVOCAR A WEB API!
        MsgBox objConsultaPaises.aWebAPI.aResposta & vbCrLf & vbCrLf & _
        "1) Verifique se este computador est� com o acesso � Internet." & vbCrLf & _
        "2) Confirme as configura��es de ANTI-V�RUS e FIREWALL que podem bloquear o acesso." & vbCrLf & _
        "3) Confirme se a URL est� acess�vel: " & vbCrLf & _
        objConsultaPaises.aWebAPI.aUrl, vbCritical, "WebApi"
    End If
    
End Sub
'-----------------------------------------------------------


'-----------------------------------------------------------
'Manipula o Banco de Dados e grava as informa��es obtidas
'-----------------------------------------------------------
Private Sub pGravar_Dados()
    
    Dim iPaisesInseridos As Integer
    Dim iIdiomasInseridos As Integer
    Dim iIdiomasDosPaisesVinculados As Integer
    
    Call objConsultaPaises.bdConectar
    
    iPaisesInseridos = objConsultaPaises.bdInserirDadosDosPaises
    iIdiomasInseridos = objConsultaPaises.bdInserirDadosDosIdiomas
    iIdiomasDosPaisesVinculados = objConsultaPaises.bdInserirDadosDosIdiomasDoPaises
    
    TxtDados.Text = "Foram inseridos no Banco de Dados:" & vbCrLf & vbCrLf & _
    "Novos pa�ses: " & iPaisesInseridos & vbCrLf & _
    "Novos idiomas: " & iIdiomasInseridos & vbCrLf & _
    "Novos v�nculos de idiomas com respectivos pa�ses: " & iIdiomasDosPaisesVinculados
    
    'Retorna mensagem para o usu�rio se inseriu algo ou n�o...
    '--------------------------------------------
    If (iPaisesInseridos > 0 Or iIdiomasInseridos > 0 Or iIdiomasDosPaisesVinculados > 0) Then
        MsgBox "Novos dados foram inseridos no Banco de Dados!", vbInformation, "WebApi"
    Else
        MsgBox "N�o foi inserido nenhum novo dado no Banco de Dados!", vbInformation, "WebApi"
    End If
    '--------------------------------------------
    
    'Desabilita op��o de SALVAR dados (aguarda nova consulta para permitir)
    BtnSalvarDados.Enabled = False
    
End Sub
'-----------------------------------------------------------

'Retorno na interface indicando que est� processando a requisi��o...
Private Sub pIndicarInicioDoProcessamento()
    Screen.MousePointer = 11
    Me.MousePointer = 11
    TxtDados.Text = ""
    DoEvents
End Sub

'Retorno na interface indicando que finalizou a requisi��o...
Private Sub pIndicarTerminoDoProcessamento()
    
    Me.MousePointer = 0
    Screen.MousePointer = 0
    
End Sub


Private Sub pGravarLOG_APRESENTACAO(ByVal sNomeEvento As String, ByVal iNumeroErro As Integer, ByVal sErro As String)

    If iNumeroErro Then
        Call objLog.gravarLog(sNomeEvento & " -> " & iNumeroErro & " " & sErro)  'Grava LOG
    Else
        Call objLog.gravarLog(sNomeEvento & " -> OK")   'Grava LOG
    End If
    
    If Err.Number <> 0 Then
        Me.MousePointer = 0
        Screen.MousePointer = 0
        MsgBox sNomeEvento & " -> " & iNumeroErro & " " & sErro, vbCritical
    End If
    
    Err.Clear

End Sub

Private Sub BtnBaixarDados_Click()
        
On Error GoTo TRATAR_ERRO_BAIXAR_DADOS
    
    BtnBaixarDados.Enabled = False
    
    'Retorno na tela indicando que est� processando a requisi��o...
    '----------------------------
    Call pIndicarInicioDoProcessamento
    '----------------------------
        
    Call pInvocarAPI 'invoca a API configurada
        
    'Requisi��o finalizada.
    '----------------------------
    Call pIndicarTerminoDoProcessamento
    '----------------------------
    
TRATAR_ERRO_BAIXAR_DADOS:

    BtnBaixarDados.Enabled = True
    Call pGravarLOG_APRESENTACAO("BtnBaixarDados_Click", Err.Number, Err.Description)
    

End Sub

Private Sub BtnSalvarDados_Click()

On Error GoTo TRATAR_ERRO_SALVAR_DADOS
    
    BtnSalvarDados.Enabled = False
    
    'Retorno na tela indicando que est� processando a requisi��o...
    '----------------------------
    Call pIndicarInicioDoProcessamento
    '----------------------------
        
    Call pGravar_Dados 'grava dados obtidos no Banco de Dados
        
    'Requisi��o finalizada.
    '----------------------------
    Call pIndicarTerminoDoProcessamento
    '----------------------------

TRATAR_ERRO_SALVAR_DADOS:
    
    Call pGravarLOG_APRESENTACAO("BtnSalvarDados_Click", Err.Number, Err.Description)
    
End Sub

Private Sub pSetarValoresIniciais()
    
    Call objLog.gravarLog("Carregou FrmConsultaPaises...") 'Grava LOG
    sDadosRetornados = ""
    Call pPermitirSalvarDados 'N�o permite TENTAR salvar os dados se n�o tiver retornado algum dado ainda
    
End Sub

Private Sub pSetarValoresFinais()

    Call objLog.gravarLog("Encerrou FrmConsultaPaises...") 'Grava LOG
    Set objConsultaPaises = Nothing
    Set objLog = Nothing
    
    sDadosRetornados = ""

End Sub

Private Sub Form_Load()

    Call pSetarValoresIniciais
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call pSetarValoresFinais
    
End Sub
