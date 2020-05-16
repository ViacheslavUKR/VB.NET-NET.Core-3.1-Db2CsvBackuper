Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class DBDumper
    Public Shared Sub DumpTableToFile(ByVal Connection As SqlConnection, ByVal TableName As String, ByVal DestinationFile As String)
        Using CMD = New SqlCommand("select * from " & TableName, Connection)
            Using RDR As SqlDataReader = CMD.ExecuteReader()
                Using OutFile = File.CreateText(DestinationFile)
                    Dim ColumnNames As String() = GetColumnNames(RDR).ToArray()
                    Dim NumFields As Integer = ColumnNames.Length
                    OutFile.WriteLine(String.Join(",", ColumnNames))
                    If RDR.HasRows Then
                        While RDR.Read()
                            Dim ColumnValues As String() = Enumerable.Range(0, NumFields).
                                Select(Function(i) DumpOneColumn(i, RDR)).
                                Select(Function(field) String.Concat("""", field.Replace("""", """"""), """")).ToArray()
                            OutFile.WriteLine(String.Join(",", columnValues))
                        End While
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Shared Function DumpOneColumn(I As Integer, RDR As SqlDataReader) As String
        If RDR.GetValue(I).GetType.FullName = "System.Byte[]" Then
            Dim Str1 = New StringBuilder()
            For Each One As Byte In RDR.GetValue(I)
                Str1.Append(One.ToString("x2"))
            Next
            Return Str1.ToString()
        Else
            Return RDR.GetValue(I).ToString()
        End If
    End Function

    Private Shared Iterator Function GetColumnNames(ByVal RDR As IDataReader) As IEnumerable(Of String)
        For Each Row As DataRow In RDR.GetSchemaTable().Rows
            Yield CStr(Row("ColumnName"))
        Next
    End Function
End Class
