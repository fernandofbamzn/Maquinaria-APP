Public Module modFormatos
    Public Function primeraLetraMayus(ByVal texto As String, Optional ByVal soloPrimeraPalabra As Boolean = True)
        If soloPrimeraPalabra Then
            Return texto.Substring(0, 1).ToUpper & texto.Substring(1).ToLower.Trim
        Else
            Dim cadena As String = ""
            For Each palabra As String In Split(texto, " ")
                cadena &= " " & palabra.Substring(0, 1).ToUpper & palabra.Substring(1).ToLower
            Next
            Return cadena.Trim
        End If
    End Function
End Module
