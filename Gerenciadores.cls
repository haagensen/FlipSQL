VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gerenciadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Tipos de banco de dados
'
Public Enum AcessoBancoGerenciadores
    Gerenciador_Nenhum = 0
    Gerenciador_Oracle = 1
    Gerenciador_SQLServer = 2
    Gerenciador_MySql = 3
    Gerenciador_Access = 4
    Gerenciador_TCPServer = 5
End Enum
