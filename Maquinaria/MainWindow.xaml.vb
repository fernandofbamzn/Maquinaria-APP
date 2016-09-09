Imports System.Data

Class MainWindow
    Dim bd As clsBaseDatos
    Dim dsMaquinaria As DataSet
    Public Sub New()

        ' Llamada necesaria para el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        InicializaBDinicial()



        Beep()
    End Sub

    Private Sub MainWindow_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'Cargo la sección de personal
        cargaPersonas()
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles Me.Closing
        If Not bd Is Nothing Then bd.close()
        bd = Nothing
    End Sub
    Private Sub InicializaBDinicial()
        bd = New clsBaseDatos
        If Not bd.isNothing Then
            dsMaquinaria = New DataSet
            dsMaquinaria.DataSetName = "Maquinaria MB"
            'recupero las tablas
            Dim sql As String = "SELECT * FROM INFORMATION_SCHEMA.tables WHERE TABLE_SCHEMA='Maquinaria MB'"
            Dim tablas As DataTable = bd.AbreDataTable(sql)
            For Each tabla As DataRow In tablas.Rows
                Dim dtTemp As DataTable = bd.AbreDataTable("SELECT * FROM " & tabla.Item("TABLE_NAME").ToString)
                dtTemp.TableName = tabla.Item("TABLE_NAME").ToString
                dsMaquinaria.Tables.Add(dtTemp)
            Next
        Else
            bd = Nothing
            Me.Close()
        End If
    End Sub

    Private Sub cargaPersonas()
        pnlPersonas.Children.Clear()
        For Each tipoPersona As DataRow In dsMaquinaria.Tables("tipoPersona").Rows

            dsMaquinaria.Tables("PERSONA").DefaultView.RowFilter = "EXT_ID_TIPOPERSONA = " & tipoPersona.Item("id_tipoPersona").ToString
            Dim name As String = tipoPersona.Item("tipoPersona").ToString & "_" & tipoPersona.Item("id_tipoPersona").ToString
            Dim titulo As String = modFormatos.primeraLetraMayus(tipoPersona.Item("tipoPersona").ToString)
            Dim cantidad As Integer = dsMaquinaria.Tables("PERSONA").DefaultView.ToTable.Rows.Count

            Dim registro As New bldRegistroBaldosaPrincipal(name, titulo, cantidad)

            pnlPersonas.Children.Add(registro)
        Next

    End Sub
End Class
