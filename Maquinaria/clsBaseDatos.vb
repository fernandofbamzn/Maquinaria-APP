Imports System.Data
Imports MySql.Data.MySqlClient
Public Class clsBaseDatos

    Private cnn As MySqlConnection
    Private cmd As MySqlCommand
    Private adaptador As MySqlDataAdapter

    Public ReadOnly Property isNothing() As Boolean
        Get
            Return cnn Is Nothing
        End Get

    End Property

    Private Function ValorBD(ByVal valor As Object) As String
        'Funcion que comprueba si un campo tiene valor NULL 
        'En ese caso devuelve una cadena vacia "", sino devuelve su valor
        If valor Is System.DBNull.Value Then
            Return ""
        ElseIf valor Is Nothing Then
            Return ""
        Else
            Return CStr(valor)
        End If
    End Function

    Public Sub New()

        Try
            Select Case MsgBox("¿La conexión a la Base de Datos es desde la RED LOCAL (Casa)?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                    abreBaseDatos(True)
                Case MsgBoxResult.No
                    abreBaseDatos(False)
                Case MsgBoxResult.Cancel
                    Me.Close()
                    Exit Sub
            End Select


        Catch ex As Exception
            MsgBox("Error al conectar con la Base de datos." & vbCrLf & ex.Message & vbCrLf & "Consulta con Fernando.")
            Me.Close()
        End Try

    End Sub
    Private Sub abreBaseDatos(ByVal isLocal As Boolean)
        If isLocal Then
            cnn = New MySqlConnection(datosConexion.CadenaConexion(datosConexion.conexion.Local))
        Else
            cnn = New MySqlConnection(datosConexion.CadenaConexion(datosConexion.conexion.Internet))
        End If
        cnn.Open()
        cmd = New MySqlCommand("", cnn)
        adaptador = New MySqlDataAdapter(cmd)
    End Sub
    Public Sub Close()
        If Not cnn Is Nothing Then
            cnn.Close()
        End If
        cnn = Nothing
        cmd = Nothing
        adaptador = Nothing
        GC.SuppressFinalize(Me)
    End Sub

    Public Function AbreDataTable(ByVal Sql As String) As DataTable

        Dim DTable As New DataTable

        If adaptador.SelectCommand.Connection.State = ConnectionState.Closed OrElse adaptador.SelectCommand.Connection.State = ConnectionState.Broken Then
            adaptador.SelectCommand.Connection.Open()
        End If
        'Se pone la consulta y se carga el datatable
        adaptador.SelectCommand.CommandText = Sql
        adaptador.SelectCommand.Parameters.Clear()
        Try
            adaptador.Fill(DTable)
        Catch ex As Exception
            DTable = New DataTable
            adaptador.Fill(DTable)
        End Try

        Return DTable

    End Function
    Public Function DevuelveValor(ByVal Sql As String) As String
        '*** El parametro "bConError" indica si al ejecutar la funcion queremos que se muestre
        '    el mensaje de error si ocurre alguna excepcion.
        If adaptador.SelectCommand.Connection.State = ConnectionState.Closed OrElse adaptador.SelectCommand.Connection.State = ConnectionState.Broken Then
            adaptador.SelectCommand.Connection.Open()
        End If
        adaptador.SelectCommand.CommandText = Sql
        adaptador.SelectCommand.Parameters.Clear()
        Return ValorBD(adaptador.SelectCommand.ExecuteScalar())
    End Function
    Public Sub EjecutaConsulta(ByVal sql As String, ByVal params As List(Of KeyValuePair(Of String, Object)))
        If adaptador.SelectCommand.Connection.State = ConnectionState.Closed OrElse adaptador.SelectCommand.Connection.State = ConnectionState.Broken Then adaptador.SelectCommand.Connection.Open()
        adaptador.SelectCommand.CommandText = sql
        adaptador.SelectCommand.Parameters.Clear()

        If Not params Is Nothing Then
            For Each p As KeyValuePair(Of String, Object) In params
                adaptador.SelectCommand.Parameters.Add(New MySqlParameter(p.Key, p.Value))
            Next
        End If

        adaptador.SelectCommand.ExecuteNonQuery()
    End Sub
End Class
