Friend Class OutlookTool

    Sub New()
        Me.OutputFolder = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    End Sub


    ''' <summary>msgファイルを書き出す既定のフォルダ</summary>
    Public Property OutputFolder As String

    Private Const DataFormats_RenPrivateSourceFolder As String = "RenPrivateSourceFolder"
    Private Const DataFormats_RenPrivateLatestMessages As String = "RenPrivateLatestMessages"
    Private Const DataFormats_RenPrivateMessages As String = "RenPrivateMessages"
    Private Const DataFormats_RenPrivateItem As String = "RenPrivateItem"

    Private Const DataFormats_FileGroupDescriptor As String = "FileGroupDescriptor"
    Private Const DataFormats_FileGroupDescriptorW As String = "FileGroupDescriptorW"


    Public Sub OnDragOver(e As DragEventArgs)
        Const CopyOrLink As DragDropEffects = DragDropEffects.Copy Or DragDropEffects.Link

        'For Each format As String In e.Data.GetFormats()
        '    System.Diagnostics.Debug.WriteLine(format)
        'Next

        If (e.AllowedEffect And CopyOrLink) = CopyOrLink _
            AndAlso e.Data.GetDataPresent(DataFormats_RenPrivateSourceFolder) Then
            'たぶんOutlookからのDragとおもわれる場合

            '受付可能に
            e.Effect = CopyOrLink

        End If
    End Sub

    Public Sub OnDragDrop(e As DragEventArgs)
        If Not DataObjectTool.HasFileContents(e.Data) Then
            Return
        End If

        Dim fileList As List(Of DescriptorItem) = DataObjectTool.ReadFileDescriptor(e.Data)
        If fileList.Count = 0 Then
            Return
        End If

        Dim ed As New DropEventArgs(fileList)
        ed.OutputFolder = Me.OutputFolder

        RaiseEvent Drop(Me, ed)
        If ed.Cancel Then
            Return
        End If

        For Each item As DescriptorItem In fileList
            If String.Equals(System.IO.Path.GetExtension(item.Name), ".msg", StringComparison.OrdinalIgnoreCase) Then
                '拡張子がmsgになっているデータのみ対象とする

                '出力先
                Dim path As String = System.IO.Path.Combine(ed.OutputFolder, item.Name)

                DataObjectTool.SaveFileContentsToFile(e.Data, item.Index, path)
            End If
        Next
    End Sub

    ''' <summary>ドロップされた時に一覧を通知し、保存するか判定させるイベント</summary>
    Public Event Drop As EventHandler(Of DropEventArgs)
End Class

Class DropEventArgs
    Inherits System.ComponentModel.CancelEventArgs
    Public Sub New(ByVal item As List(Of DescriptorItem))
        Me.Items = item.ToList()
    End Sub

    Public ReadOnly Items As ICollection(Of DescriptorItem)

    ''' <summary>保存先</summary>
    ''' <returns></returns>
    Public Property OutputFolder As String
End Class


