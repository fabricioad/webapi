VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'REFERÊNCIA:
'http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries

'ATRIBUTOS DA TAG tLanguage no XML específico
Public sISOCode As String
Public sName As String
