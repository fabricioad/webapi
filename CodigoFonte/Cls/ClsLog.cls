VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'AUTOR: FABRÍCIO A. DINIZ
'DATA DA ÚLTIMA ALTERAÇÃO: 24/01/2021
'HORA DAQ ÚLTIMA ALTERAÇÃO: 13:40

'Objetivo: Registrar um LOG simplificado em arquivo de texto

Dim sNomeArquivoDeLog As String
Dim sCaminhoDoArquivo As String

Public Property Get aNomeArquivoDeLog() As String
 aNomeArquivoDeLog = sNomeArquivoDeLog
End Property

Public Property Let aNomeArquivoDeLog(ByVal sValor As String)
 sNomeArquivoDeLog = sValor
 sCaminhoDoArquivo = App.Path & "\" & sNomeArquivoDeLog
End Property

Private Sub pSetarValoresIniciais()
    
     sNomeArquivoDeLog = "WebAPI.log"
     sCaminhoDoArquivo = App.Path & "\" & sNomeArquivoDeLog
    
End Sub

Public Sub gravarLog(ByVal sConteudo As String)
 Dim numArq As Long
 
 numArq = FreeFile
 
 If Dir(sCaminhoDoArquivo, vbArchive) <> "" Then
 
    Open sCaminhoDoArquivo For Append As #numArq 'ABRIR arquivo para adicionar em LOG existente
 Else
    Open sCaminhoDoArquivo For Output As #numArq 'ABRIR arquivo para criar um LOG novo
 End If
        
    'Grava o conteúdo no arquivo de LOG
    Print #numArq, Format(Now, "dd-mm-yyyy hh:mm:ss") & " -> " & sConteudo
        
    Close #numArq 'FECHAR arquivo de LOG aberto
End Sub

Private Sub Class_Initialize()
 Call pSetarValoresIniciais
End Sub
