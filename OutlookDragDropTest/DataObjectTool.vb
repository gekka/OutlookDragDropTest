Friend Class DataObjectTool

    Public Const DataFormats_FileGroupDescriptorW As String = "FileGroupDescriptorW"
    Public Const DataFormats_FileFontents As String = "FileContents"

    Public Shared Function HasFileGroupDescriptor(ByVal data As IDataObject) As Boolean
        Return data.GetFormats().Contains(DataFormats_FileGroupDescriptorW)
    End Function

    Public Shared Function HasFileContents(ByVal data As IDataObject) As Boolean
        Return data.GetFormats().Contains(DataFormats_FileFontents)
    End Function

    ''' <summary>IDataObjectからファイル名の一覧を抽出する</summary>
    Public Shared Function ReadFileDescriptor(ByVal data As IDataObject) As List(Of DescriptorItem)
        Dim fileList As New List(Of DescriptorItem)

        If Not HasFileGroupDescriptor(data) Then
            Return fileList
        End If

        Dim bs As Byte() = GetBytes(data, DataFormats_FileGroupDescriptorW)

        Dim gh = System.Runtime.InteropServices.GCHandle.Alloc(bs, System.Runtime.InteropServices.GCHandleType.Pinned)
        Try
            Dim mem As IntPtr = gh.AddrOfPinnedObject()

            Dim fileGroup = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(mem, GetType(Win32.FILEGROUPDESCRIPTOR)), Win32.FILEGROUPDESCRIPTOR)

            Dim count As UInteger = fileGroup.cItems
            If count > 0 Then
                Dim size_citem = System.Runtime.InteropServices.Marshal.SizeOf(fileGroup.cItems)
                Dim size_descriptor = System.Runtime.InteropServices.Marshal.SizeOf(fileGroup.data)
                Dim p As IntPtr = mem + size_citem

                Dim desc As Win32.FILEDESCRIPTOR

                For i As Integer = 0 To CInt(fileGroup.cItems - 1)
                    desc = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(p, GetType(Win32.FILEDESCRIPTOR)), Win32.FILEDESCRIPTOR)

                    fileList.Add(New DescriptorItem(i, desc.cFileName))

                    p += size_descriptor
                Next
            End If

        Finally
            gh.Free()
        End Try
        Return fileList
    End Function

    ''' <summary>FileContentsに複数の情報が入っているのをIndex指定して取り出してファイルに保存する</summary>
    ''' <param name="data"></param>
    ''' <param name="index">FileContentsから取り出したいインデックス</param>
    ''' <param name="filePath">取り出した内容の保存先のファイルパス</param>
    ''' <remarks>FileContentsからインデックスで取り出す方法が標準ではないので、APIを使って取り出す。</remarks>
    Public Shared Sub SaveFileContentsToFile(ByVal data As IDataObject, ByVal index As Integer, ByVal filePath As String)

        If Not data.GetFormats().Contains(DataFormats_FileFontents) Then
            Throw New ApplicationException("FileContentsが含まれていません")
        End If

        If index < 0 Then
            Throw New ArgumentOutOfRangeException(NameOf(index))
        End If

        If System.IO.File.Exists(filePath) Then
            Throw New ApplicationException("ファイルが既に存在してます")
        End If


        Dim format As New System.Runtime.InteropServices.ComTypes.FORMATETC
        format.cfFormat = BitConverter.ToInt16(BitConverter.GetBytes(CType(DataFormats.GetFormat(DataFormats_FileFontents).Id, UShort)), 0)
        format.ptd = IntPtr.Zero
        format.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT
        format.lindex = index '取り出すインデックスを指定
        format.tymed = Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE
        '=  System.Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTREAM _
        'Or Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL _
        'Or Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE

        Dim medium As New System.Runtime.InteropServices.ComTypes.STGMEDIUM

        Dim id As System.Runtime.InteropServices.ComTypes.IDataObject = CType(data, System.Runtime.InteropServices.ComTypes.IDataObject)
        id.GetData(format, medium)
        Select Case (medium.tymed)
            Case Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE
                Dim storage As Win32.IStorage = DirectCast(System.Runtime.InteropServices.Marshal.GetTypedObjectForIUnknown(medium.unionmember, GetType(Win32.IStorage)), Win32.IStorage)
                If storage IsNot Nothing Then

                    Try

                        Dim IID_IStorage As Guid = GetType(Win32.IStorage).GUID
                        Dim copy As Win32.IStorage = Nothing
                        Dim obj As Object = Nothing

                        'ファイルをIStorageとして開く
                        Dim ret As Integer = Win32.StgCreateStorageEx(filePath, Win32.STGM.STGM_CREATE Or Win32.STGM.STGM_READWRITE Or Win32.STGM.STGM_SHARE_EXCLUSIVE, Win32.STGFMT.STGFMT_STORAGE, 0, IntPtr.Zero, IntPtr.Zero, IID_IStorage, obj)
                        If ret = 0 AndAlso obj IsNot Nothing Then
                            Try
                                copy = CType(obj, Win32.IStorage)
                                storage.CopyTo(0, Nothing, IntPtr.Zero, copy) 'ファイルにコピー
                                copy.Commit(Win32.STGC.STGC_DEFAULT)
                            Catch ex As Exception
                                Throw
                            Finally
                                Release(copy)
                                Release(obj)
                            End Try
                        End If
                    Finally
                        Release(storage)
                    End Try
                End If
            Case Else
                'IStorageで来るはず
                Throw New NotSupportedException("IStorege以外のデータは処理してません")
        End Select

    End Sub

    Private Shared Sub Release(Of T As Class)(ByRef o As T)
        If o Is Nothing OrElse Not o.GetType().IsCOMObject Then
            Return
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        o = Nothing
    End Sub

    Private Shared Function GetBytes(data As IDataObject, format As String) As Byte()
        Dim ms = TryCast(data.GetData(format), System.IO.MemoryStream)
        If ms Is Nothing Then
            Return Nothing
        End If

        Dim bs(CInt(ms.Length)) As Byte
        ms.Read(bs, 0, bs.Length)
        Return bs
    End Function


End Class
