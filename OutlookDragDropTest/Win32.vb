Imports System.Runtime.InteropServices

Friend Class Win32

#Region "Enum"
    <Flags>
    Enum FileDescriptorFlags As UInteger
        FD_CLSID = &H1
        FD_SIZEPOINT = &H2
        FD_ATTRIBUTES = &H4
        FD_CREATETIME = &H8
        FD_ACCESSTIME = &H10
        FD_WRITESTIME = &H20
        FD_FILESIZE = &H40
        FD_PROGRESSUI = &H4000
        FD_LINKUI = &H8000
        FD_UNICODE = &H80000000L
    End Enum

    <Flags>
    Enum STGM As UInteger
        STGM_READ = &H0
        STGM_WRITE = &H1
        STGM_READWRITE = &H2
        STGM_SHARE_EXCLUSIVE = &H10
        STGM_SHARE_DENY_WRITE = &H20
        STGM_SHARE_DENY_READ = &H30
        STGM_SHARE_DENY_NONE = &H40
        STGM_CREATE = &H1000
    End Enum

    Enum STGFMT As UInteger
        STGFMT_STORAGE = 0
        STGFMT_FILE = 3
        STGFMT_ANY = 4
        STGFMT_DOCFILE = 5
    End Enum

    <Flags>
    Enum STGC As UInteger
        STGC_DEFAULT = 0
        STGC_OVERWRITE = 1
        STGC_ONLYIFCURRENT = 2
        STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 4
        STGC_CONSOLIDATE = 8
    End Enum


    Enum STGTY As Integer
        STGTY_STORAGE = 1
        STGTY_STREAM = 2
        STGTY_LOCKBYTES = 3
        STGTY_PROPERTY = 4
    End Enum

#End Region

#Region "Structure"

    <Runtime.InteropServices.StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Structure FILEDESCRIPTOR
        Public dwFlags As FileDescriptorFlags
        Public clsid As Guid
        Public sizel As System.Drawing.Size
        Public pointl As System.Drawing.Point
        Public dwFileAttributes As UInt32
        Public ftCreationTime As ComTypes.FILETIME
        Public ftLastAccessTime As ComTypes.FILETIME
        Public ftLastWriteTime As ComTypes.FILETIME
        Public nFileSizeHigh As UInt32
        Public nFileSizeLow As UInt32
        <Runtime.InteropServices.MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
    End Structure

    <Runtime.InteropServices.StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Structure FILEGROUPDESCRIPTOR
        Public cItems As UInteger
        Public data As FILEDESCRIPTOR
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Structure STATSTG

        Public pwcsName As IntPtr 'String

        Public type As Integer
        Public cbSize As Long
        Public mtime As ComTypes.FILETIME
        Public ctime As ComTypes.FILETIME
        Public atime As ComTypes.FILETIME
        Public grfMode As Integer
        Public grfLocksSupported As Integer
        Public clsid As Guid
        Public grfStateBits As Integer
        Public reserved As Integer


        Public Overrides Function ToString() As String
            Dim s As String = ""
            If pwcsName = IntPtr.Zero Then
                s = ""
            Else
                s = Marshal.PtrToStringUni(pwcsName)
            End If
            s = s & vbTab & CType(type, STGTY).ToString() & vbTab & cbSize & vbTab & mtime.ToString() & vbTab & ctime.ToString() & vbTab & atime.ToString() & vbTab & grfMode & grfLocksSupported & clsid.ToString() & grfStateBits & vbTab & reserved
            Return s
        End Function
    End Structure

#End Region

#Region "Interface"
    <ComImportAttribute()>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("0000000b-0000-0000-c000-000000000046")>
    Interface IStorage
        Sub CreateStream(ByVal pwcsName As String, ByVal grfMode As UInteger, ByVal reserved1 As UInteger, ByVal reserved2 As UInteger, ByRef ppstm As ComTypes.IStream)
        Sub OpenStream(ByVal pwcsName As String, ByVal reserved1 As IntPtr, ByVal grfMode As UInteger, ByVal reserved2 As UInteger, ByRef ppstm As ComTypes.IStream)
        Sub CreateStorage(ByVal pwcsName As String, ByVal grfMode As UInteger, ByVal reserved1 As UInteger, ByVal reserved2 As UInteger, ByRef ppstg As IStorage)
        Sub OpenStorage(ByVal pwcsName As String, ByVal pstgPriority As IStorage, ByVal grfMode As UInteger, ByVal snbExclude As IntPtr, ByVal reserved As UInteger, ByRef ppstg As IStorage)
        Sub CopyTo(ByVal ciidExclude As UInteger, ByVal rgiidExclude() As Guid, ByVal snbExclude As IntPtr, ByVal pstgDest As IStorage)
        Sub MoveElementTo(ByVal pwcsName As String, ByVal pstgDest As IStorage, ByVal pwcsNewName As String, ByVal grfFlags As UInteger)
        Sub Commit(ByVal grfCommitFlags As UInteger)
        Sub Revert()
        Sub EnumElements(ByVal reserved1 As UInteger, ByVal reserved2 As IntPtr, ByVal reserved3 As UInteger, <Out> ByRef ppenum As IntPtr) ' As IEnumSTATSTG)
        Sub DestroyElement(ByVal pwcsName As String)
        Sub RenameElement(ByVal pwcsOldName As String, ByVal pwcsNewName As String)
        Sub SetElementTimes(ByVal pwcsName As String, ByVal pctime As ComTypes.FILETIME, ByVal patime As ComTypes.FILETIME, ByVal pmtime As ComTypes.FILETIME)
        Sub SetClass(ByVal clsid As Guid)
        Sub SetStateBits(ByVal grfStateBits As UInteger, ByVal grfMask As UInteger)
        Sub Stat(ByRef pstatstg As ComTypes.STATSTG, ByVal grfStatFlag As UInteger)
    End Interface
#End Region

#Region "Function"
    <DllImport("ole32.dll", SetLastError:=True, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Winapi)>
    Shared Function StgCreateStorageEx _
        (<InAttribute, MarshalAs(UnmanagedType.LPWStr)> pwcsName As String _
        , grfMode As STGM _
        , stgfmt As STGFMT _
        , grfAttrs As Integer _
        , pStgOptions As IntPtr _
        , reserved2 As IntPtr _
        , <InAttribute> ByRef riid As Guid _
        , <OutAttribute, MarshalAs(Runtime.InteropServices.UnmanagedType.IUnknown)> ByRef ppObjectOpen As Object) As Integer
    End Function

    Private Shared Function FILETIME2DateTime(ByVal ftime As ComTypes.FILETIME) As DateTime
        Return DateTime.FromFileTime((CType(ftime.dwHighDateTime, Int64) << 32) Or ftime.dwLowDateTime)
    End Function
#End Region
End Class
