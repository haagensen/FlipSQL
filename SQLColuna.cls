VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLColuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msNome As String
Private meTipo As ADODB.DataTypeEnum
Private msngTamanho As Single
Private mbNull As Boolean
Private mbIdentity As Boolean

Public Property Get Nome() As String
    Nome = msNome
End Property
Public Property Let Nome(ByVal sNome As String)
    msNome = sNome
End Property

Public Property Get Tipo() As ADODB.DataTypeEnum
    Tipo = meTipo
End Property
Public Property Let Tipo(ByVal eTipo As ADODB.DataTypeEnum)
    meTipo = eTipo
End Property

Public Property Get Tamanho() As Single
    Tamanho = msngTamanho
End Property
Public Property Let Tamanho(ByVal sngTamanho As Single)
    msngTamanho = sngTamanho
End Property

Public Property Get IsNull() As Boolean
    IsNull = mbNull
End Property
Public Property Let IsNull(ByVal bNull As Boolean)
    mbNull = bNull
End Property

Public Property Get IsIdentity() As Boolean
    IsIdentity = mbIdentity
End Property
Public Property Let IsIdentity(ByVal bIdentity As Boolean)
    mbIdentity = bIdentity
End Property

Public Sub Novo(ByVal sNome As String, ByVal eTipo As ADODB.DataTypeEnum, ByVal sngTamanho As Single, Optional ByVal bIsNull As Boolean = True, Optional ByVal bIsIdentity As Boolean = False)
    msNome = sNome
    meTipo = eTipo
    msngTamanho = sngTamanho
    mbNull = IIf(Not IsMissing(bIsNull), bIsNull, True)
    mbIdentity = IIf(Not IsMissing(bIsIdentity), bIsIdentity, False)
    If mbIdentity Then mbNull = False 'Identity nunca pode ser Null
End Sub

Private Sub Class_Initialize()
    mbNull = True 'Default...
End Sub
