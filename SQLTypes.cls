VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLTipos"
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

Public Function ToString(ByVal oCol As SQLColuna) As String

    Dim sTipo As String
    Dim sIdentity As String
    
    sTipo = "?"

    Select Case oCol.Tipo
        Case adBigInt
            If meBanco = Gerenciador_Oracle Then
                sTipo = "NUMBER(19,0)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "BIGINT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "BIGINT"
            End If
        Case adBinary
            If meBanco = Gerenciador_Oracle Then
                sTipo = IIf(oCol.Tamanho < 2001, "RAW(" & CInt(oCol.Tamanho) & ")", "LONG RAW")
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "IMAGE"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "LONGBLOB"
            End If
        Case adBoolean
            If meBanco = Gerenciador_Oracle Then
                sTipo = "NUMBER(1,0)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "BIT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "BOOLEAN"
            End If
        Case adChar
            If meBanco = Gerenciador_Oracle Then
                sTipo = "CHAR(" & CInt(oCol.Tamanho) & ")"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "CHAR(" & CInt(oCol.Tamanho) & ")"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "CHAR(" & CInt(oCol.Tamanho) & ")"
            End If
        Case adCurrency
            If meBanco = Gerenciador_Oracle Then
                sTipo = "DECIMAL(13,3)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "MONEY"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "DECIMAL(13,3)"
            End If
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            If meBanco = Gerenciador_Oracle Then
                sTipo = "DATE"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "DATETIME"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "DATETIME"
            End If
        Case adDecimal, adNumeric
            If meBanco = Gerenciador_Oracle Then
                sTipo = "FLOAT(24,0)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "DECIMAL"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "DECIMAL"
            End If
        Case adDouble
            If meBanco = Gerenciador_Oracle Then
                sTipo = "FLOAT"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "FLOAT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "DOUBLE"
            End If
        Case adInteger
            If meBanco = Gerenciador_Oracle Then
                sTipo = "NUMBER(10,0)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "INT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "INT"
            End If
        Case adLongVarBinary
            If meBanco = Gerenciador_Oracle Then
                sTipo = "BLOB" ' sem especificar tamanho, o default é de 2 GB
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "VARBINARY(MAX)" ' acima de 8.000 bytes
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "VARBINARY(" & CInt(oCol.Tamanho) & ")"
            End If
        Case adLongVarChar
            If meBanco = Gerenciador_Oracle Then
                sTipo = "CLOB"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "TEXT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "LONGTEXT"
            End If
        Case adLongVarWChar
            If meBanco = Gerenciador_Oracle Then
                sTipo = "NCLOB"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "NTEXT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "NLONGTEXT"
            End If
        Case adSingle
            If meBanco = Gerenciador_Oracle Then
                sTipo = "FLOAT(24,0)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "REAL"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "REAL"
            End If
        Case adSmallInt
            If meBanco = Gerenciador_Oracle Then
                sTipo = "NUMBER(5,0)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "SMALLINT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "SMALLINT"
            End If
        Case adTinyInt
            If meBanco = Gerenciador_Oracle Then
                sTipo = "NUMBER(3,0)"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "TINYINT"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "TINYINT"
            End If
        Case adVarBinary
            If meBanco = Gerenciador_Oracle Then
                sTipo = "BLOB" ' sem especificar tamanho, o default é de 2 GB
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = IIf(oCol.Tamanho > 8000, "VARBINARY(MAX)", "VARBINARY(" & CInt(oCol.Tamanho) & ")")
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "VARBINARY(" & CInt(oCol.Tamanho) & ")"
            End If
        Case adVarChar, adBSTR
            If meBanco = Gerenciador_Oracle Then
                sTipo = IIf(oCol.Tamanho > 4000, "CLOB", "VARCHAR2(" & CInt(oCol.Tamanho) & ")")
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = IIf(oCol.Tamanho > 4000, "TEXT", "VARCHAR(" & CInt(oCol.Tamanho) & ")")
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = IIf(oCol.Tamanho > 4000, "TEXT", "VARCHAR(" & CInt(oCol.Tamanho) & ")")
            End If
        Case adVarWChar
            If meBanco = Gerenciador_Oracle Then
                sTipo = IIf(oCol.Tamanho > 4000, "NCLOB", "NVARCHAR2(" & CInt(oCol.Tamanho) & ")")
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = IIf(oCol.Tamanho > 4000, "NTEXT", "NVARCHAR(" & CInt(oCol.Tamanho) & ")")
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = IIf(oCol.Tamanho > 4000, "NTEXT", "NVARCHAR(" & CInt(oCol.Tamanho) & ")")
            End If
        Case adWChar
            If meBanco = Gerenciador_Oracle Then
                sTipo = "NCHAR (" & CInt(oCol.Tamanho) & ")"
            ElseIf meBanco = Gerenciador_SQLServer Then
                sTipo = "NCHAR (" & CInt(oCol.Tamanho) & ")"
            ElseIf meBanco = Gerenciador_MySql Then
                sTipo = "NCHAR (" & CInt(oCol.Tamanho) & ")"
            End If
    End Select

    If oCol.IsIdentity Then
        Select Case meBanco
            Case Gerenciador_SQLServer: sIdentity = "IDENTITY"
            Case Gerenciador_MySql:     sIdentity = "AUTO_INCREMENT"
        End Select
    End If

    If sIdentity <> "" Then
        ToString = oCol.Nome & " " & sTipo & " " & sIdentity
    Else
        ToString = oCol.Nome & " " & sTipo & IIf(oCol.IsNull, "", " NOT") & " NULL"
    End If
    
End Function

