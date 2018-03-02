Attribute VB_Name = "code"
Option Explicit

'example of create table
Public Sub exampleCreateTable()
  Dim sqlLite: Set sqlLite = New sqlLite
  '----------------------------------------------'
  Dim qry As Variant, dbPath As String
  '----------------------------------------------'
  sqlLite.dbPath = ActiveWorkbook.Path & "\db\test.db"
  'sqlLite_qry.execute "create table abc (a string, b string)" 'faz o select na base de dados e printa as colunas do print'
  'sqlLite_qry.execute "delete from testeNum" 'comando delete
  '----------------------------------------------'
End Sub

'example of insert of data / single line and multiple lines
Public Sub exampleInsert()
  Dim sqlLite As sqlLite: Set sqlLite = New sqlLite
  '----------------------------------------------'
  Dim qry As Variant, dbPath As String
  '----------------------------------------------'
  sqlLite.dbPath = ActiveWorkbook.Path & "\db\test.db"
  'sqlLite_qry.execute "insert into testeNum(numeros) values(44000),(55000) "  '2 values insert
  sqlLite.execute "insert into testeNum(numeros) values" & montaQueryInsertToTeste   'multiple values
  '----------------------------------------------'
End Sub

'example of querying data
Public Sub exampleSelect()
  Dim sqlLite As sqlLite: Set sqlLite = New sqlLite
  '----------------------------------------------'
  Dim qry As Variant, dbPath As String
  '----------------------------------------------'
  sqlLite.dbPath = ActiveWorkbook.Path & "\db\test.db"
  sqlLite.selectQry "select * from testeNum limit 100"  'faz o select na base de dados e printa as colunas do print'
  '----------------------------------------------'
  Range(Cells(1, 1), Cells(1, sqlLite.qtdColunas)).Value = sqlLite.cabecalho 'cola cabecalho
  Range(Cells(2, 1), Cells(sqlLite.qtdLinhas + 1, sqlLite.qtdColunas)).Value = sqlLite.dados 'cola os dados
  '----------------------------------------------'
End Sub

'example multiple line insert
Public Function montaQueryInsertToTeste() As String
  Dim c As Range
  Dim arr() As Variant
  Dim k As Long
  Range("a1:a50000").Select
  '------------------------------
  For Each c In Selection
    k = k + 1
    c.Value = k
    ReDim Preserve arr(1 To k)
    arr(k) = "(" & c.Value & ")"
  Next c
  '------------------------------
  montaQueryInsertToTeste = Join(arr, ",")
End Function
