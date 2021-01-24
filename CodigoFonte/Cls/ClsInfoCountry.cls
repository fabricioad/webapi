VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInfoCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'REFERÊNCIA:
'http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries



'ATRIBUTOS DA TAG tCountryInfo no XML específico
'-----------------------------------------------
Public sISOCode As String
Public sName As String
Public sCapitalCity As String
Public sPhoneCode As String
Public sContinentCode As String
Public sCurrencyISOCode As String
Public sCountryFlag

Private oLanguages() As ClsLanguage

Public iTotalDeLanguages As Long
'-----------------------------------------------

Public Property Get aLanguages() As ClsLanguage()
    Set aLanguages = oLanguages
End Property

Public Sub adicionarLanguage(ByVal sISOCode As String, ByVal sName As String)
    
    ReDim Preserve Languages(iTotalDeLanguages)
    Set oLanguages(iTotalDeLanguages) = New ClsLanguage
    
    oLanguages(iTotalDeLanguages).sISOCode = sISOCode
    oLanguages(iTotalDeLanguages).sName = sName
    
    iTotalDeLanguages = iTotalDeLanguages + 1
    
End Sub

Private Sub Class_Initialize()

    iTotalDeLanguages = 0
    Erase Languages
        
End Sub
