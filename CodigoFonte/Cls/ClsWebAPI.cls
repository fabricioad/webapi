VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsWebAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'AUTOR: FABR�CIO A. DINIZ
'DATA DA �LTIMA ALTERA��O: 24/01/2021
'HORA DAQ �LTIMA ALTERA��O: 13:40

'Objetivo: Realizar a requisi��o da API em uma URL especificada

'--------------------------------------------------------------------------------
'**DLL DE REFER�NCIA** -> ..\WINDOWS\system32\winhttpcom.dll#Microsoft WinHTTP
'--------------------------------------------------------------------------------

'CONSTANTES
'------------------------

'CREATEOBJECT:
Const OBJ_WINHTTP = "WinHttp.WinHttprequest.5.1"
'------------------------

'Atributos da Classe
'---------------------
Private sURLBase As String 'URL base para fazer a requisi��o
Private sOperacao As String 'Nome da Opera��o a ser executada na requisi��o
Private sResposta As String 'Resposta da requisi��o
Private sStatusCode As String 'Retorna o C�DIGO do status HTTP da �ltima resposta
Private sStatusText As String 'Retorna o TEXTO do status HTTP da �ltima resposta
Private sRetornoXML As String 'Valores extra�dos de um Atributo de uma TAG existente no XML ap�s a requisi��o

'Objetos
Private oLog As ClsLog
Private oHttp As Object 'Faz a requisi��o da Opera��o na URL Base
'---------------------

'------------------------------------------------
'IN�CIO: ATRIBUTOS DA CLASSE
'------------------------------------------------
Public Property Get aUrl() As String
 aUrl = sURLBase
End Property

Public Property Let aUrl(ByVal sValor As String)
 sURLBase = sValor
End Property

Public Property Get aOperacao() As String
 aOperacao = sOperacao
End Property

Public Property Let aOperacao(ByVal sValor As String)
 sOperacao = sValor
End Property

Public Property Get aResposta() As String
 aResposta = sResposta
End Property

Public Property Get aStatusCode() As String
 aStatusCode = sStatusCode
End Property

Public Property Get aStatusText() As String
 aStatusText = sStatusText
End Property

Public Property Get aRetornoXML() As String
 aRetornoXML = sRetornoXML
End Property


'------------------------------------------------
'FIM: PROPRIEDADES DA CLASSE
'------------------------------------------------

'------------------------------------------------
'IN�CIO: M�TODOS DA CLASSE
'------------------------------------------------
Public Function mInvocar() As Integer
  
  On Error GoTo TRATAR_ERRO_INVOCAR
  
  sStatusCode = ""
  sStatusText = ""
  
  'Inst�ncia do objeto necess�rio para efetuar a requisi��o da API
  Call pInstanciarObjWinHttp
  
  'Efetua a requisi��o da API na URL + Operacao
  '---------------------------------
  oHttp.open "Get", sURLBase & "/" & sOperacao, False
  oHttp.send
  '---------------------------------
  
  sStatusCode = CStr(oHttp.Status)
  sStatusText = oHttp.statusText
  
  sResposta = oHttp.responseText 'Resposta retornada pela requisi��o
  
TRATAR_ERRO_INVOCAR:
  
  If Err.Number <> 0 Then
    oLog.gravarLog ("WebAPI.mInvocar -> " & Err.Number & " " & Err.Description)
    sResposta = "WebAPI.mInvocar -> " & Err.Number & " " & Err.Description
    mInvocar = 0
  Else
    oLog.gravarLog ("WebAPI.mInvocar -> OK")
    mInvocar = 1
  End If
  
  Err.Clear
  
End Function

'------------------------------------------------
'FIM: M�TODOS DA CLASSE
'------------------------------------------------

'------------------------------------------------
'IN�CIO: Fun��es e Procedimento INTERNOS (Privados)
'------------------------------------------------

Private Sub pSetarValoresIniciais()
   
   sResposta = ""
   sStatusCode = ""
   sStatusText = ""
   sOperacao = ""
   sURLBase = ""
   Set oLog = New ClsLog
   
   Call pLiberarInstanciaWinHttp
    
End Sub

Private Sub pSetarValoresFinais()
    Set oLog = Nothing
    Call pLiberarInstanciaWinHttp
End Sub


Private Sub pLiberarInstanciaWinHttp()
     Set oHttp = Nothing
End Sub

Private Sub pInstanciarObjWinHttp()
   
   If oHttp Is Nothing Then
    Set oHttp = CreateObject(OBJ_WINHTTP)
  End If
  
End Sub
'------------------------------------------------
'FIM: Fun��es e Procedimento INTERNOS (Privados)
'------------------------------------------------

'CONSTRUTOR
Private Sub Class_Initialize()
    Call pSetarValoresIniciais
End Sub

'DESTRUTOR
Private Sub Class_Terminate()
    Call pSetarValoresFinais
End Sub
