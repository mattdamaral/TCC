Public Class Material

    Private nome As String
    Private quantidadeTotal As Integer

    Public Sub New(_nome As String, _quantidade As Integer)

        nome = _nome
        quantidadeTotal = _quantidade

    End Sub

    Public Sub AdicionaQuantidade(_quantidade As Integer)
        quantidadeTotal += _quantidade
    End Sub

    Public Function GetNome()
        Return nome
    End Function

    Public Function GetQuantidadeTotal()
        Return quantidadeTotal
    End Function

End Class
