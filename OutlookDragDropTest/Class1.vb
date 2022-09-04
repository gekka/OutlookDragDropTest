'IStorageの中身を見てみる


'Private Sub SaveMsgFromOutlook(ByVal data As System.Windows.Forms.DataObject, ByVal count As Integer)
'    Dim app As Microsoft.Office.Interop.Outlook.Application = Nothing

'    Dim ins As Outlook.Inspector = Nothing
'    Dim exp As Outlook.Explorer = Nothing
'    Dim sel As Outlook.Selection = Nothing
'    Dim w As Object = Nothing


'    Try
'        app = New Microsoft.Office.Interop.Outlook.Application()
'        w = app.ActiveWindow
'        ins = app.ActiveInspector
'        exp = app.ActiveExplorer

'        If w Is ins Then
'            Throw New NotImplementedException
'        ElseIf w Is exp Then

'            sel = exp.Selection
'            If sel.Count = count Then
'                For i As Integer = 1 To sel.Count

'                    Dim outputPath As String = GetTempFilePath()

'                    Dim obj As Object
'                    obj = sel.Item(i)
'                    Try
'                        If TypeOf obj Is Outlook.AppointmentItem Then
'                            Dim ai As Outlook.AppointmentItem = CType(obj, Outlook.AppointmentItem)
'                            Try
'                                ai.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(ai)
'                            End Try

'                        ElseIf TypeOf obj Is Outlook.ContactItem Then
'                            Dim ci As Outlook.ContactItem = CType(obj, Outlook.ContactItem)
'                            Try
'                                ci.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(ci)
'                            End Try

'                        ElseIf TypeOf obj Is Outlook.DistListItem Then
'                            Dim di As Outlook.DistListItem = CType(obj, Outlook.DistListItem)
'                            Try
'                                di.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(di)
'                            End Try

'                        ElseIf TypeOf obj Is Outlook.JournalItem Then
'                            Dim ji As Outlook.JournalItem = CType(obj, Outlook.JournalItem)
'                            Try
'                                ji.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(ji)
'                            End Try

'                        ElseIf TypeOf obj Is Outlook.MailItem Then
'                            Dim mi As Outlook.MailItem = CType(obj, Outlook.MailItem)
'                            Try
'                                mi.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(mi)
'                            End Try

'                        ElseIf TypeOf obj Is Outlook.NoteItem Then
'                            Dim ni As Outlook.NoteItem = CType(obj, Outlook.NoteItem)
'                            Try
'                                ni.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(ni)
'                            End Try

'                        ElseIf TypeOf obj Is Outlook.PostItem Then
'                            Dim pi As Outlook.PostItem = CType(obj, Outlook.PostItem)
'                            Try
'                                pi.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(pi)
'                            End Try

'                        ElseIf TypeOf obj Is Outlook.TaskItem Then
'                            Dim ti As Outlook.TaskItem = CType(obj, Outlook.TaskItem)
'                            Try
'                                ti.SaveAs(outputPath, Outlook.OlSaveAsType.olMSG)
'                            Finally
'                                Release(ti)
'                            End Try
'                        End If

'                    Finally
'                        Release(obj)
'                    End Try

'                Next
'            End If

'        Else
'            Throw New NotImplementedException
'        End If
'    Finally
'        Release(w)
'        Release(ins)
'        Release(exp)

'        Release(app)
'    End Try
'End Sub

'Private Function GetTempFilePath() As String
'    Dim temp As String = System.IO.Path.GetTempFileName()
'    System.IO.File.Delete(temp)
'    temp = System.IO.Path.ChangeExtension(temp, ".msg")
'    Return temp
'End Function

'Private Sub Release(ByRef o As Object)
'    If o Is Nothing OrElse Not o.GetType().IsCOMObject Then
'        Return
'    End If
'    System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
'    o = Nothing
'End Sub

'Private Function GetBytes(data As IDataObject, format As String) As Byte()
'    Dim ms = TryCast(data.GetData(format), System.IO.MemoryStream)
'    If ms Is Nothing Then
'        Return Nothing
'    End If

'    Dim bs(ms.Length) As Byte
'    ms.Read(bs, 0, bs.Length)
'    Return bs
'End Function

'''' <summary>FileContentsに複数の情報が入っているのをIndex指定して取り出してファイルに保存する</summary>
'''' <param name="data"></param>
'''' <param name="index">0始まりのインデックス</param>
'''' <param name="filePath"></param>
'''' <returns></returns>
'Private Function SaveToFile(data As IDataObject, index As Integer, filePath As String)
'    If index < 0 Then
'        Throw New ArgumentOutOfRangeException(NameOf(index))
'    End If

'    If System.IO.File.Exists(filePath) Then
'        Throw New ApplicationException("ファイルが既に存在してます")
'    End If

'    Dim format As New System.Runtime.InteropServices.ComTypes.FORMATETC
'    format.cfFormat = BitConverter.ToInt16(BitConverter.GetBytes(CType(DataFormats.GetFormat(Formats_FileFontents).Id, UShort)), 0)
'    format.ptd = IntPtr.Zero
'    format.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT
'    format.lindex = index '取り出すインデックス
'    format.tymed = Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE
'    '=  System.Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTREAM _
'    'Or Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL _
'    'Or Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE

'    Dim medium As New System.Runtime.InteropServices.ComTypes.STGMEDIUM

'    Dim id As System.Runtime.InteropServices.ComTypes.IDataObject = data
'    id.GetData(format, medium)
'    Select Case (medium.tymed)
'        Case Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE
'            Dim storage As Win32.IStorage = System.Runtime.InteropServices.Marshal.GetTypedObjectForIUnknown(medium.unionmember, GetType(Win32.IStorage))
'            If storage IsNot Nothing Then

'                Try

'                    Dim IID_IStorage As Guid = GetType(Win32.IStorage).GUID
'                    Dim copy As Win32.IStorage
'                    Dim o As Object
'                    Dim ret As Integer = Win32.StgCreateStorageEx(filePath, Win32.STGM.STGM_CREATE Or Win32.STGM.STGM_READWRITE Or Win32.STGM.STGM_SHARE_EXCLUSIVE, Win32.STGFMT.STGFMT_STORAGE, 0, 0, IntPtr.Zero, IID_IStorage, o)
'                    If ret = 0 AndAlso o IsNot Nothing Then
'                        Try
'                            copy = o
'                            storage.CopyTo(0, Nothing, IntPtr.Zero, copy)
'                            copy.Commit(Win32.STGC.STGC_DEFAULT)
'                        Catch ex As Exception
'                            System.Diagnostics.Debug.WriteLine(ex.Message)
'                        Finally
'                            Release(copy)
'                            Release(o)
'                        End Try
'                    End If
'                Finally
'                    Release(storage)
'                End Try
'            End If
'        Case Else
'            'IStorageで来るはず
'            Throw New NotSupportedException("IStorege以外のデータは処理してません")
'    End Select

'    Return Nothing
'End Function



'Private Function GetFileContents_(data As DataObject, index As Integer)
'    Dim format As New System.Runtime.InteropServices.ComTypes.FORMATETC
'    format.cfFormat = BitConverter.ToInt16(BitConverter.GetBytes(CType(DataFormats.GetFormat(Formats_FileFontents).Id, UShort)), 0)
'    format.ptd = IntPtr.Zero
'    format.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT
'    format.lindex = index
'    format.tymed _
'        = System.Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTREAM _
'        Or Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL _
'        Or Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE

'    Dim medium As New System.Runtime.InteropServices.ComTypes.STGMEDIUM

'    Dim id As System.Runtime.InteropServices.ComTypes.IDataObject = data
'    id.GetData(format, medium)
'    Select Case (medium.tymed)
'        Case Runtime.InteropServices.ComTypes.TYMED.TYMED_ISTORAGE
'            Dim storage As IStorage = System.Runtime.InteropServices.Marshal.GetTypedObjectForIUnknown(medium.unionmember, GetType(IStorage))
'            If storage IsNot Nothing Then
'                Try

'                    Dim ip As IntPtr = IntPtr.Zero
'                    storage.EnumElements(0, 0, 0, ip)

'                    Dim es As IEnumSTATSTG
'                    es = System.Runtime.InteropServices.Marshal.GetTypedObjectForIUnknown(ip, GetType(IEnumSTATSTG))

'                    es.Reset()

'                    'Dim statstg As New System.Runtime.InteropServices.ComTypes.STATSTG



'                    Dim idx As UInt32 = 1
'                    Do While True
'                        Dim statstg As New STATSTG
'                        Dim count As UInteger = 0
'                        Dim hresult As UInteger
'                        hresult = es.Next(1, statstg, count)
'                        If hresult <> 0 Then
'                            Throw New System.ComponentModel.Win32Exception()
'                        ElseIf count = 0 Then
'                            Exit Do
'                        Else
'                            Try
'                                System.Diagnostics.Debug.WriteLine(statstg.ToString())

'                                Dim name As String = System.Runtime.InteropServices.Marshal.PtrToStringUni(statstg.pwcsName)


'                                Dim lastaccess As DateTime = FILETIME2DateTime(statstg.atime)
'                                Dim creationTime As DateTime = FILETIME2DateTime(statstg.ctime)
'                                Dim modifyTime As DateTime = FILETIME2DateTime(statstg.mtime)

'                                Const STGM_SHARE_EXCLUSIVE As Integer = 16
'                                Select Case (CType(statstg.type, STGTY))
'                                    Case STGTY.STGTY_STORAGE

'                                        Dim subIS As IStorage
'                                        storage.OpenStorage(name, Nothing, STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, subIS)
'                                        Release(subIS)

'                                    Case STGTY.STGTY_STREAM
'                                        Dim st As System.Runtime.InteropServices.ComTypes.IStream
'                                        storage.OpenStream(name, IntPtr.Zero, STGM_SHARE_EXCLUSIVE, 0, st)

'                                        Dim bs(statstg.cbSize) As Byte

'                                        Dim p As IntPtr = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(4)
'                                        Try
'                                            st.Read(bs, bs.Length, p)
'                                            Dim read As Integer = System.Runtime.InteropServices.Marshal.ReadInt32(p)

'                                            Dim s = System.Text.Encoding.Unicode.GetString(bs)
'                                            'System.Diagnostics.Debug.WriteLine(s)
'                                        Finally
'                                            System.Runtime.InteropServices.Marshal.FreeCoTaskMem(p)
'                                            Release(st)
'                                        End Try


'                                    Case STGTY.STGTY_LOCKBYTES
'                                    Case STGTY.STGTY_PROPERTY

'                                End Select
'                            Finally
'                                System.Runtime.InteropServices.Marshal.FreeCoTaskMem(statstg.pwcsName)
'                            End Try



'                        End If
'                        idx += 1
'                    Loop
'                    'storage.OpenStream()
'                Finally
'                    System.Runtime.InteropServices.Marshal.ReleaseComObject(storage)
'                End Try
'            End If

'    End Select

'    Return Nothing
'End Function


'<Guid("0000000d-0000-0000-C000-000000000046")>
'<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'<ComImport>
'Interface IEnumSTATSTG

'    Function [Next](ByVal celt As UInteger, <Out> ByRef regelt As STATSTG, <Out> ByRef pceltFetched As UInteger) As UInteger

'    Sub Skip(ByVal celt As UInteger)

'    Sub Reset()

'    Sub Clone(<Out> ByRef ppenum As IEnumSTATSTG)
'End Interface