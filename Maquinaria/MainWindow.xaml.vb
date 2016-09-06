Class MainWindow 
    Dim bd As clsBaseDatos
    Public Sub New()

        ' Llamada necesaria para el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        bd = New clsBaseDatos

        datos.DataContext = bd.AbreDataTable("SELECT * FROM TipoContacto ")
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles Me.Closing
        If Not bd Is Nothing Then bd.close()
        bd = Nothing
    End Sub
End Class
