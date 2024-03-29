Imports System.Data
Imports System.Data.SqlClient

Public Class SqlHandler

    Private _connectionString As String
    Public Property ConnectionString As String
        Get
            Return _connectionString
        End Get
        Set(value As String)
            _connectionString = value
        End Set
    End Property

    Sub New()

    End Sub
    ''' <summary>
    ''' SqlHandler Start
    ''' </summary>
    ''' <param name="IConnectionString">Example : Data Source=.\MSSQLSERVER2017;Initial Catalog=ExampleDatabase;Persist Security Info=True;User ID=SqlUserName;Password=123456789;MultipleActiveResultSets=True</param>
    Sub New(IConnectionString As String)
        _connectionString = IConnectionString
    End Sub

    Public Sub ExeCuteNonQuery(CommandText As String)
        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = CommandText
                Com.ExecuteNonQuery()
            End Using
            Con.Close()
        End Using
    End Sub

    Public Sub ExeCuteNonQuery(CommandText As String, ParameterName As String, ParameterValue As Object)
        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = CommandText
                Com.Parameters.AddWithValue(ParameterName, ParameterValue)
                Com.ExecuteNonQuery()
                Com.Parameters.Clear()
            End Using
            Con.Close()
        End Using

    End Sub

    Public Sub ExeCuteNonQuery(CommandText As String, Parameters As List(Of ISqlParameters))
        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = CommandText
                For Each P As ISqlParameters In Parameters
                    Com.Parameters.AddWithValue(P.Name, P.Value)
                Next
                Com.ExecuteNonQuery()
                Com.Parameters.Clear()
            End Using
            Con.Close()
        End Using
    End Sub


    Public Sub ExeCuteNonQueryAsStoredProcedure(StoredProcedureName As String)
        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = StoredProcedureName
                Com.CommandType = CommandType.StoredProcedure
                Com.ExecuteNonQuery()
                Com.Parameters.Clear()
            End Using
            Con.Close()
        End Using
    End Sub
    Public Sub ExeCuteNonQueryAsStoredProcedure(StoredProcedureName As String, ParameterName As String, ParameterValue As Object)
        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = StoredProcedureName
                Com.CommandType = CommandType.StoredProcedure
                Com.Parameters.AddWithValue(ParameterName, ParameterValue)
                Com.ExecuteNonQuery()
                Com.Parameters.Clear()
            End Using
            Con.Close()
        End Using
    End Sub

    Public Sub ExeCuteNonQueryAsStoredProcedure(StoredProcedureName As String, Parameters As List(Of ISqlParameters))
        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = StoredProcedureName
                Com.CommandType = CommandType.StoredProcedure
                For Each P As ISqlParameters In Parameters
                    Com.Parameters.AddWithValue(P.Name, P.Value)
                Next
                Com.ExecuteNonQuery()
                Com.Parameters.Clear()
            End Using
            Con.Close()
        End Using

    End Sub



    Public Function ExeCuteScalar(CommandText As String) As Object
        Dim gd As Object

        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = CommandText
                gd = Com.ExecuteScalar()
            End Using
            Con.Close()
        End Using

        Return gd
    End Function


    Public Function ExeCuteScalar(CommandText As String, ParameterName As String, ParameterValue As Object) As Object
        Dim gd As Object

        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = CommandText
                Com.Parameters.AddWithValue(ParameterName, ParameterValue)
                gd = Com.ExecuteScalar()
            End Using
            Con.Close()
        End Using

        Return gd
    End Function

    Public Function ExeCuteScalar(CommandText As String, Parameters As List(Of ISqlParameters)) As Object
        Dim gd As Object

        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = CommandText
                For Each P As ISqlParameters In Parameters
                    Com.Parameters.AddWithValue(P.Name, P.Value)
                Next
                gd = Com.ExecuteScalar()
            End Using
            Con.Close()
        End Using

        Return gd
    End Function





    Public Function ExeCuteScalarAsStoredProcedure(StoredProcedureName As String) As Object
        Dim gd As Object

        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = StoredProcedureName
                Com.CommandType = CommandType.StoredProcedure
                gd = Com.ExecuteScalar()
            End Using
            Con.Close()
        End Using

        Return gd
    End Function

    Public Function ExeCuteScalarAsStoredProcedure(StoredProcedureName As String, ParameterName As String, ParameterValue As Object) As Object
        Dim gd As Object

        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = StoredProcedureName
                Com.CommandType = CommandType.StoredProcedure
                Com.Parameters.AddWithValue(ParameterName, ParameterValue)
                gd = Com.ExecuteScalar()
            End Using
            Con.Close()
        End Using

        Return gd
    End Function

    Public Function ExeCuteScalarAsStoredProcedure(StoredProcedureName As String, Parameters As List(Of ISqlParameters)) As Object
        Dim gd As Object

        Using Con As New SqlConnection(_connectionString)
            Con.Open()
            Using Com As New SqlCommand
                Com.Connection = Con
                Com.CommandText = StoredProcedureName
                Com.CommandType = CommandType.StoredProcedure
                For Each P As ISqlParameters In Parameters
                    Com.Parameters.AddWithValue(P.Name, P.Value)
                Next
                gd = Com.ExecuteScalar()
            End Using
            Con.Close()
        End Using

        Return gd
    End Function



End Class

Public Structure ISqlParameters
    ''' <summary>
    ''' Parametre Adı
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Name As String
    ''' <summary>
    ''' Parametre Değeri
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Value As Object
End Structure
