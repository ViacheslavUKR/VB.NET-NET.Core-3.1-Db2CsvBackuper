Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip
Imports System.IO

Public Class Zip
    Public Shared Sub ZipFolder(ByVal OutPathname As String, ByVal FolderName As String)
        If File.Exists(OutPathname) Then
            File.Delete(OutPathname)
        End If
        Dim FsOut As FileStream = File.Create(OutPathname)
        Dim ZipStream As ZipOutputStream = New ZipOutputStream(FsOut)
        ZipStream.SetLevel(3)
        Dim FolderOffset As Integer = FolderName.Length + (If(FolderName.EndsWith("\"), 0, 1))
        CompressFolder(FolderName, ZipStream, FolderOffset)
        ZipStream.IsStreamOwner = True
        ZipStream.Close()
    End Sub

    Private Shared Sub CompressFolder(ByVal Path As String, ByVal ZipStream As ZipOutputStream, ByVal FolderOffset As Integer)
        Dim Files As String() = Directory.GetFiles(Path)

        For Each Filename As String In Files
            Dim FI As FileInfo = New FileInfo(Filename)
            Dim EntryName As String = Filename.Substring(FolderOffset)
            EntryName = ZipEntry.CleanName(EntryName)
            Dim NewEntry As ZipEntry = New ZipEntry(EntryName)
            NewEntry.DateTime = FI.LastWriteTime
            NewEntry.Size = FI.Length
            ZipStream.PutNextEntry(NewEntry)
            Dim Buffer As Byte() = New Byte(4095) {}

            Using StreamReader As FileStream = File.OpenRead(Filename)
                StreamUtils.Copy(StreamReader, ZipStream, Buffer)
            End Using
            ZipStream.CloseEntry()
        Next

        Dim Folders As String() = Directory.GetDirectories(Path)
        For Each folder As String In Folders
            CompressFolder(folder, ZipStream, FolderOffset) 'recursion
        Next
    End Sub

    Public Shared Sub AddFileToZip(ByVal zipFile As String, ByVal fileToAdd As String)
        If File.Exists(zipFile) Then
            File.Delete(zipFile)
        End If
        Using ZipStream As ZipOutputStream = New ZipOutputStream(File.Create(zipFile))
            Dim Buffer As Byte() = New Byte(4095) {}
            Dim Entry As ZipEntry = New ZipEntry(System.IO.Path.GetFileName(fileToAdd))
            Entry.DateTime = DateTime.Now
            Entry.Size = New IO.FileInfo(fileToAdd).Length
            ZipStream.PutNextEntry(Entry)
            Using FS As FileStream = File.OpenRead(fileToAdd)
                Dim SourceBytes As Integer = 0
                While FS.Position < Entry.Size
                    SourceBytes = FS.Read(Buffer, 0, Buffer.Length)
                    If SourceBytes > 0 Then ZipStream.Write(Buffer, 0, SourceBytes)
                End While
            End Using
            ZipStream.Finish()
            ZipStream.Close()
            ZipStream.Dispose()
        End Using
    End Sub
End Class
