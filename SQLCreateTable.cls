VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private meBanco As AcessoBancoGerenciadores

'
'##############################################################################
' PROPRIEDADES
'##############################################################################
'

Public Property Let Banco(ByVal v As AcessoBancoGerenciadores)
    meBanco = v
End Property
Public Property Get Banco() As AcessoBancoGerenciadores
    Banco = meBanco
End Property

'
'##############################################################################
' MÉTODOS
'##############################################################################
'

Public Function Create(ByVal sTabela As String, ByVal col As SQLColunas) As String

    Dim oTipo As SQLTipos
    Dim elemento As SQLColuna
    Dim sb As StringBuilder

    If col Is Nothing Then Exit Function

    Set sb = New StringBuilder
    Set oTipo = New SQLTipos
    oTipo.Banco = meBanco

    For Each elemento In col
        sb.Add oTipo.ToString(elemento)
    Next

    Create = "CREATE TABLE " & sTabela & " (" & _
             sb.ToDelimitedString(", ") & _
             ")"

End Function
