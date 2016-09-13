VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLAlterTable"
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

Public Function AddColumn(ByVal sTabela As String, ByVal oCol As SQLColuna) As String

    Dim oTipo As SQLTipos
    
    Set oTipo = New SQLTipos
    oTipo.Banco = meBanco
    
    If meBanco = Gerenciador_MySql Then
        AddColumn = "ALTER TABLE " & sTabela & " ADD COLUMN " & oTipo.ToString(oCol)
    ElseIf meBanco = Gerenciador_Oracle Then
        AddColumn = "ALTER TABLE " & sTabela & " ADD (" & oTipo.ToString(oCol) & ")"
    ElseIf meBanco = Gerenciador_SQLServer Then
        AddColumn = "ALTER TABLE " & sTabela & " ADD " & oTipo.ToString(oCol)
    End If

End Function

Public Function ModifyColumn(ByVal sTabela As String, ByVal oCol As SQLColuna) As String

    Dim oTipo As SQLTipos

    Set oTipo = New SQLTipos
    oTipo.Banco = meBanco

    If meBanco = Gerenciador_MySql Then
        ModifyColumn = "ALTER TABLE " & sTabela & " CHANGE COLUMN " & oCol.Nome & " " & oTipo.ToString(oCol)
    ElseIf meBanco = Gerenciador_Oracle Then
        ModifyColumn = "ALTER TABLE " & sTabela & " MODIFY " & oTipo.ToString(oCol)
        ModifyColumn = Replace(ModifyColumn, "NOT NULL", "") ' Informar NOT NULL em uma coluna já NOT NULL gera "ORA-01451: coluna não pode ser modificada para NULL"
        ModifyColumn = Replace(ModifyColumn, "NULL", "")
    ElseIf meBanco = Gerenciador_SQLServer Then
        ModifyColumn = "ALTER TABLE " & sTabela & " ALTER COLUMN " & oTipo.ToString(oCol)
    End If

End Function

Public Function AddPK(ByVal sTabela As String, ByVal col As SQLColunas) As String

    Dim oTipo As SQLTipos
    Dim oColuna As SQLColuna
    Dim sChaves As String

    Set oTipo = New SQLTipos
    oTipo.Banco = meBanco

    For Each oColuna In col
        sChaves = sChaves & "," & oColuna.Nome
    Next
    sChaves = Mid$(sChaves, 2)

    If meBanco = Gerenciador_MySql Then
        AddPK = "ALTER TABLE " & sTabela & " ADD PRIMARY KEY (" & sChaves & ")"
    ElseIf meBanco = Gerenciador_Oracle Then
        AddPK = "ALTER TABLE " & sTabela & " ADD (PRIMARY KEY (" & sChaves & "))"
    ElseIf meBanco = Gerenciador_SQLServer Then
        AddPK = "ALTER TABLE " & sTabela & " ADD PRIMARY KEY (" & sChaves & ")"
    End If

End Function

Public Function AddFK(ByVal sTabela As String, ByVal sCol As String, ByVal sTabelaReferenciada As String, ByVal sColReferenciada As String) As String

    If meBanco = Gerenciador_MySql Then
        ' "ALTER TABLE CAIXA ADD FOREIGN KEY (COD_DEPTO) REFERENCES DEPTO(COD_DEPTO)"
        AddFK = "ALTER TABLE " & sTabela & " ADD FOREIGN KEY (" & sCol & ") REFERENCES " & sTabelaReferenciada & "(" & sColReferenciada & ")"
    ElseIf meBanco = Gerenciador_Oracle Then
        AddFK = "ALTER TABLE " & sTabela & " ADD (FOREIGN KEY(" & sCol & ") REFERENCES " & sTabelaReferenciada & "(" & sColReferenciada & "))"
    ElseIf meBanco = Gerenciador_SQLServer Then
        ' alter table MyTable add constraint MyTable_MyColumn_FK FOREIGN KEY ( MyColumn ) references MyOtherTable(PKColumn)
        AddFK = "ALTER TABLE " & sTabela & " ADD CONSTRAINT " & sTabela & "_" & sCol & "_FK FOREIGN KEY(" & sCol & ") REFERENCES " & sTabelaReferenciada & "(" & sColReferenciada & ")"
    End If

End Function

