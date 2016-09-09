Public Class bldRegistroBaldosaPrincipal
    Inherits UserControl
    Public Sub New(ByVal name As String, ByVal textoSeccion As String, ByVal numRegistros As Integer)

        ' Llamada necesaria para el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        Me.Name = name
        textoIdentificador.Text = textoSeccion
        textoCantidad.Text = numRegistros

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click
        Dim addRegistro As New winAddRegistro()
        addRegistro.ShowDialog()
    End Sub
End Class
