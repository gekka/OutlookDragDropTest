Friend Class DescriptorItem
    Public Sub New(i As Integer, name As String)
        Me.Index = i
        Me.Name = name
    End Sub

    Public ReadOnly Name As String
    Public ReadOnly Index As Integer
End Class