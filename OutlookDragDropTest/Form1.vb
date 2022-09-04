Friend Class Form1

    Private WithEvents tool As New OutlookTool
    Private listBox1 As ListBox

    Sub New()
        Me.AllowDrop = True


        Me.Width = 400
        Me.Height = 400
        listBox1 = New ListBox()
        listBox1.Location = New Point(5, 5)
        listBox1.Size = New Size(Me.ClientSize.Width - 10, Me.ClientSize.Height - 10)
        listBox1.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom

        Me.Controls.Add(listBox1)
    End Sub

    Private Sub Form1_DragOver(sender As Object, e As DragEventArgs) Handles MyBase.DragOver
        tool.OnDragOver(e)
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Try
            tool.OnDragDrop(e)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub tool_Drop(sender As Object, e As DropEventArgs) Handles tool.Drop

        For Each item As DescriptorItem In e.Items
            Me.listBox1.Items.Add(item.Name)
        Next

        '保存先を指定
        e.OutputFolder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

        '保存を行う
        e.Cancel = False
    End Sub
End Class

