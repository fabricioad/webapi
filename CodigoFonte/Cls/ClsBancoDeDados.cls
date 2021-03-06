VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBancoDeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'AUTOR: FABR�CIO A. DINIZ
'DATA DA �LTIMA ALTERA��O: 24/01/2021
'HORA DAQ �LTIMA ALTERA��O: 13:40

'Objetivo: Manipular o Banco de Dados MySQL para salvar os dados obtidos

'--------------------------------------------------------------------------------
'**REFER�NCIA NECESS�RIA** -> ..\Program Files (x86)\Common Files\System\ado\msado15.dll#Microsoft ActiveX Data Objects 6.1 Library
'**REFER�NCIA NECESS�RIA** -> Driver={MySQL ODBC 5.3 ANSI Driver}
'--------------------------------------------------------------------------------

Const ARQUIVO_INI_BANCO_DE_DADOS = "\bd.ini" 'Arquivo com dados para conex�o com o Banco de Dados

'-----------------------------------------------------------------------
'APIs do WINDOWS
'-----------------------------------------------------------------------
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                        ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                        ByVal lpDefault As String, ByVal lpReturnedString As String, _
                        ByVal nSize As Long, ByVal lpFileName As String) As Long
'-----------------------------------------------------------------------

'--------------------------------------------
Private oConnection As ADODB.Connection
Private oRecordSet As ADODB.Recordset
Private oCmdSQL As ADODB.Command
Private oParametroSQL() As ADODB.Parameter

Private oLog As ClsLog

Private bConectado As Boolean
Private sDriver As String
Private sServer As String
Private sDatabase As String
Private sUser As String
Private sPassword As String
Private sNomeArquivoConfigInicial As String
'----------------------------------------------

Public Property Get aDriver() As String
 aDriver = sDriver
End Property

Public Property Let aDriver(ByVal sValor As String)
 sDriver = sValor
End Property

Public Property Get aServer() As String
 aServer = sServer
End Property

Public Property Let aServer(ByVal sValor As String)
 sServer = sValor
End Property

Public Property Get aDatabase() As String
 aDatabase = sDatabase
End Property

Public Property Let aDatabase(ByVal sValor As String)
 sDatabase = sValor
End Property

Public Property Get aUser() As String
 aUser = sUser
End Property

Public Property Let aUser(ByVal sValor As String)
 sUser = sValor
End Property

Public Property Get aConectado() As Boolean
 aConectado = bConectado
End Property


Public Property Get aPassword() As String
 aPassword = sPassword
End Property

Public Property Let aPassword(ByVal sValor As String)
 sPassword = sValor
End Property

Public Property Get aCmdSQL() As ADODB.Command
 Set aCmdSQL = oCmdSQL
End Property

Public Property Get aConnection() As ADODB.Connection
 Set aConnection = oConnection
End Property

Public Property Get aRecordset() As ADODB.Recordset
 Set aRecordset = oRecordSet
End Property

Public Property Get aLog() As ClsLog
 Set aLog = oLog
End Property



'-------------------------------------------------------------
'LEITURA e GRAVA��O em arquivos INI
'-------------------------------------------------------------
Private Function pLerINI(sSection As String, sKeyName As String, sINIFileName As String) As String
    Dim sRet As String

  sRet = String(255, Chr(0))
  pLerINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))
End Function


Private Function pGravarINI(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
  
  Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
  pGravarINI = (Err.Number = 0)
  
End Function
'-------------------------------------------------------------


'------------------------------------------------------------------------------
'Leitura do arquivo INI contendo DRIVER, SERVER, DATABASE, USER E PASSWORD
'para estebelecer a conex�o com o Banco De Dados
'------------------------------------------------------------------------------
Public Sub mFazerLeituraDoArquivoIniAcessoBD()
    
    sNomeArquivoConfigInicial = App.Path & ARQUIVO_INI_BANCO_DE_DADOS
    
    'Ainda n�o existe arquivo INI?
    If Dir(sNomeArquivoConfigInicial, vbArchive) = "" Then
        
        'N�o existe! Cria um novo...
        sDriver = pGravarINI("BancoDeDados", "Driver", "MySQL ODBC 5.3 ANSI Driver", sNomeArquivoConfigInicial)
        sServer = pGravarINI("BancoDeDados", "Server", "127.0.0.1", sNomeArquivoConfigInicial)
        sDatabase = pGravarINI("BancoDeDados", "Database", "mutant", sNomeArquivoConfigInicial)
        sUser = pGravarINI("BancoDeDados", "User", "seu_usuario_aqui", sNomeArquivoConfigInicial)
        sPassword = pGravarINI("BancoDeDados", "Password", "sua_senha_aqui", sNomeArquivoConfigInicial)
        
    Else 'existe, faz a leitura...
    
        sDriver = pLerINI("BancoDeDados", "Driver", sNomeArquivoConfigInicial)
        sServer = pLerINI("BancoDeDados", "Server", sNomeArquivoConfigInicial)
        sDatabase = pLerINI("BancoDeDados", "Database", sNomeArquivoConfigInicial)
        sUser = pLerINI("BancoDeDados", "User", sNomeArquivoConfigInicial)
        sPassword = pLerINI("BancoDeDados", "Password", sNomeArquivoConfigInicial)
        
    End If
End Sub
'------------------------------------------------------------------------------

Private Sub pSetarValoresIniciais()
    
    bConectado = False
    
    Call pFecharRecordSet
    Call pFecharConnection
    
    
    Set oConnection = New ADODB.Connection
    Set oRecordSet = New ADODB.Recordset
    
    Set oLog = New ClsLog
    
    Call mFazerLeituraDoArquivoIniAcessoBD
    
End Sub

Private Sub pSetarValoresFinais()
    
    bConectado = False
    
    Call pFecharRecordSet
    Call pFecharConnection
    Set oLog = Nothing
    
End Sub

Private Function pRemoverUltimoCaractere(ByVal sTexto As String) As String
    Dim sRetorno As String
    
    sRetorno = ""
    
    If Len(sTexto) > 0 Then
        sRetorno = Mid(sTexto, 1, Len(sTexto) - 1) 'remove o �ltimo caractere
    End If
    
    pRemoverUltimoCaractere = sRetorno
    
End Function

Private Sub pFecharRecordSet()
        
    If Not (oRecordSet Is Nothing) Then
        If oRecordSet.State = 1 Then 'estava aberto?
            oRecordSet.Close 'fecha!
        End If
        Set oRecordSet = Nothing
    End If
    
End Sub

Private Sub pFecharConnection()
        
    If Not (oConnection Is Nothing) Then
        If oConnection.State = adStateOpen Then 'estava aberto?
           oConnection.Close  'fecha!
        End If
        Set oConnection = Nothing
    End If
    
End Sub

Private Sub pResetCmdSQL()
        
        Set oCmdSQL = Nothing
        Set oCmdSQL = New ADODB.Command
    
End Sub

'Executa SQL e retorna o total de registros retornados ou -1 se ocorrer erro
'---------------------------------------------------------------------------
Public Function mExecutarSQL(ByVal strSQL As String) As Integer
    
  On Error GoTo TRATAR_ERRO_EXECUTAR_SQL
    
    '---------------------------
    'Banco de dados n�o conectado?
    If oConnection.State <> adStateOpen Then
        Call mConectar 'Tenta conectar
        'N�o conectou ainda?
        If oConnection.State <> adStateOpen Then
            'Aborta procedimento
            Call oLog.gravarLog("EXECUTAR SQL: Conex�o com BD n�o estabelecida...")
            mExecutarSQL = 0
            Exit Function
        End If
    End If
    '---------------------------
    
    Call oLog.gravarLog("EXECUTAR SQL: " & strSQL)
    
    '---------------------------------------------
    
    Call pFecharRecordSet
    Set oRecordSet = New Recordset
    oRecordSet.CursorLocation = adUseClient
    oRecordSet.open strSQL, oConnection
    '---------------------------------------------
    
    mExecutarSQL = CStr(oRecordSet.RecordCount)  'Retorna total de registros retornados/afetados
    Exit Function
    
    
TRATAR_ERRO_EXECUTAR_SQL:

    '---------------------------------------------------
    Call oLog.gravarLog("EXECUTAR SQL: " & Err.Number & " - " & Err.Description)
    mExecutarSQL = -1 'Erro ao executar SQL
    '---------------------------------------------------


End Function

Public Function mInserirDados(ByVal sNomeDaTabela As String, sNomeDosCampos() As String, sTipoDosCampos() As String, sValoresDosCampos() As String) As Integer
    
    On Error GoTo TRATAR_ERRO_INSERIR_BD
    
    '---------------------------
    Dim strSQL As String
    Dim strSQLCampos As String
    Dim strSQLValores As String
    Dim strValores As String
    '---------------------------
    Dim i As Integer
    
    '---------------------------
    'Banco de dados n�o conectado?
    If oConnection.State <> adStateOpen Then
        Call mConectar 'Tenta conectar
        'N�o conectou ainda?
        If oConnection.State <> adStateOpen Then
            'Aborta procedimento
            Call oLog.gravarLog("INSERT SQL***: Conex�o com BD n�o estabelecida...")
            mInserirDados = 0
            Exit Function
        End If
    End If
    '---------------------------
    
    '---------------------------
    'SQL:
    'INSERT INTO table_name (column1, column2, column3, ...)
    'VALUES (?, ?, ?, ...);
    '------------------------------------------
    
    'COLUNAS
    '------------------------------------------
    strSQLCampos = "("
    strSQLValores = "("
    strValores = "("
    For i = 0 To UBound(sNomeDosCampos)
    
        strSQLCampos = strSQLCampos & sNomeDosCampos(i) & ","
        strSQLValores = strSQLValores & "?," 'strSQLValores & "@" & sNomeDosCampos(i) & ","
        strValores = strValores & sValoresDosCampos(i) & ","
        
    Next i
    
    strSQLCampos = pRemoverUltimoCaractere(strSQLCampos) 'remove a �ltima ","
    strSQLValores = pRemoverUltimoCaractere(strSQLValores) 'remove a �ltima ","
    strValores = pRemoverUltimoCaractere(strValores) 'remove a �ltima ","
    
    strSQLValores = strSQLValores & ")"
    strValores = strValores & ")"
    strSQLCampos = strSQLCampos & ")"
    
    '------------------------------------------
    
    'VALORES DAS COLUNAS - DE ACORDO COM O TIPO: INT ou STR?
    '------------------------------------------
    
    'Executar comando com preven��o de SQL INJECTION
    '-----------------------------------------------
    
    Erase oParametroSQL
    ReDim oParametroSQL(UBound(sNomeDosCampos))
    Call pResetCmdSQL
    '-----------------------------------------------
    '-----------------------------------------------
    
    Call oLog.gravarLog("INSERT SQL VALORES: " & strValores) 'LOG dos dados que ser�o inseridos no BD
    
    'MONTA QUERY DE INSER��O
    '-----------------------------------------------------------
    strSQL = "INSERT INTO " & sNomeDaTabela & " " & strSQLCampos & " VALUES " & strSQLValores
    '-----------------------------------------------------------
    
    '-----------------------------------------------------------
    'PREVEN��O DE SQL INJECTION
    '-----------------------------------------------------------
    With oCmdSQL
        .CommandText = strSQL
        .CommandType = adCmdText
        .Prepared = True
    End With
    '-----------------------------------------------------------
    
    'strSQLValores = "("
    For i = 0 To UBound(sNomeDosCampos)
        
            Set oParametroSQL(i) = New ADODB.Parameter
            
            Select Case UCase(sTipoDosCampos(i))
            
             Case "STR" 'VALOR STRING
                Set oParametroSQL(i) = oCmdSQL.CreateParameter("@" & sNomeDosCampos(i), adVarChar, adParamInput, 120, sValoresDosCampos(i))
                
             Case "INT" 'VALOR INTEIRO
                Set oParametroSQL(i) = oCmdSQL.CreateParameter("@" & sNomeDosCampos(i), adInteger, adParamInput, , CInt(Val(sValoresDosCampos(i))))
                                              
            End Select
            '-----------------------------------------------------------
        
        '------------------------------------------------
        'PREVEN��O A SQL INJECTION
        '------------------------------------------------
        oCmdSQL.Parameters.Append oParametroSQL(i)
        '------------------------------------------------
        
    Next i
    '------------------------------------------
        
    'INSER��O COM SUCESSO...
    '------------------------------------------------
    oCmdSQL.ActiveConnection = oConnection
    Call oLog.gravarLog("INSERT SQL: " & strSQL)
    Set oRecordSet = oCmdSQL.Execute  'oConnection.Execute(strSQL)
    mInserirDados = 1
    Exit Function
    '------------------------------------------------
    
    
    
TRATAR_ERRO_INSERIR_BD:
    
    'ERRO NA TENTATIVA DE INSER��O...
    '---------------------------------------------------
    Call oLog.gravarLog("ERRO INSERT SQL**: " & Err.Number & " - " & Err.Description)
    mInserirDados = 0
    '---------------------------------------------------
    
End Function

Private Sub pGravar_LOG_CONFIG_CONEXAO_BANCO_DE_DADOS()
    
    With oLog
        Call .gravarLog("BD Driver: " & sDriver)
        Call .gravarLog("BD Server: " & sServer)
        Call .gravarLog("BD Database: " & sDatabase)
        Call .gravarLog("BD Usuario: " & sUser)
        Call .gravarLog("BD Password: " & String$(Len(sPassword), "*"))
    End With
    
End Sub

Private Sub pGravar_LOG_STATUS_CONEXAO_BD(ByVal sResultadoConexao As String)

    Call oLog.gravarLog("Conex�o com Banco De Dados: " & sResultadoConexao)
    
End Sub

Public Function mConectar() As String
    
   Dim sRetorno As String
    
   On Error GoTo TratarErro
   
    'Dados para estabelecer CONEX�O com o Banco de Dados
    '------------------------------------
    oConnection.ConnectionString = _
    "Driver={" & sDriver & "};" & _
    "Server=" & sServer & ";" & _
    "Database=" & sDatabase & ";" & _
    "User=" & sUser & ";" & _
    "Password=" & sPassword & ";"
    Call pGravar_LOG_CONFIG_CONEXAO_BANCO_DE_DADOS 'LOG dos dados De conex�o
    '----------------------------
    
    
    oConnection.open 'Estabelece a conex�o
    
    sRetorno = "OK"
    Call pGravar_LOG_STATUS_CONEXAO_BD(sRetorno) 'LOG de sucesso
    mConectar = sRetorno
    bConectado = True
    Exit Function
    
TratarErro:
    
    bConectado = False
    sRetorno = Err.Number & " " & vbCrLf & Err.Description 'Erro de conex�o
    Call pGravar_LOG_STATUS_CONEXAO_BD(sRetorno) 'LOG do erro de conex�o
    mConectar = sRetorno
   
End Function

Private Sub Class_Initialize()
    Call pSetarValoresIniciais
End Sub
