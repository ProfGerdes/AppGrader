Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports SharpCompress.Archive
Imports SharpCompress.Common

Namespace ArchiveExtractor
    'Extension to allow us to filter files in a folder using multiple extensions
    NotInheritable Class MyExtension
        Private Sub New()
        End Sub
        '   <System.Runtime.CompilerServices.Extension> _
        Public Shared Function GetFilesByExtensions(dir As DirectoryInfo, ParamArray extensions As String()) As IEnumerable(Of FileInfo)
            If extensions Is Nothing Then
                Throw New ArgumentNullException("extensions")
            End If
            Dim files As IEnumerable(Of FileInfo) = dir.EnumerateFiles()
            Return files.Where(Function(f) extensions.Contains(f.Extension))
        End Function
    End Class
    Class Program
        'Dictionary to keep track which files already processed
        Private Shared _processDict As Dictionary(Of [String], List(Of [String]))

        Private Shared Sub Main(args As String())
            _processDict = New Dictionary(Of String, List(Of String))()
            'repeat until user decide to call quit
            While True
                Console.Write("Directory to Process (x to Exit, Enter to process current directory): ")
                Dim cmd = Console.ReadLine()

                'Setup the initial directory
                Dim curDir As DirectoryInfo = Nothing
                If [String].IsNullOrEmpty(cmd) Then
                    curDir = New DirectoryInfo(Directory.GetCurrentDirectory())
                ElseIf cmd.ToLower() = "x" Then
                    Return
                Else
                    curDir = New DirectoryInfo(cmd)
                    If Not curDir.Exists Then
                        Console.WriteLine("Invalid directory")
                        Continue While
                    End If
                End If
                'Process the directory
                ProcessDirectory(curDir, "")
            End While
        End Sub

        Private Shared Sub ProcessDirectory(dir As DirectoryInfo, prefix As [String])
            Console.WriteLine("{0}[{1}]", prefix, dir.FullName)
            _processDict.Add(dir.FullName, New List(Of String)())

            Dim subdirs = dir.GetDirectories()
            For Each subdir As var In subdirs
                'Recursively process the sub-directories
                ProcessDirectory(subdir, "   ")
            Next
            'In case the .ZIP extracts to .RAR,
            'process the files twice
            ProcessFiles(dir, prefix)
            ProcessFiles(dir, prefix)
        End Sub

        Private Shared Sub ProcessFiles(dir As DirectoryInfo, prefix As [String])
            For Each f As var In IO.Directory.GetFiles(dir.ToString, ".zip", SearchOption.AllDirectories)      ' dir.GetFilesByExtensions(".zip", ".rar")
                'have we processed this file before?
                If Not _processDict(dir.FullName).Contains(f.Name) Then
                    Console.WriteLine("{0}  {1}", prefix, f.Name)
                    'nope, mark it as processed
                    _processDict(dir.FullName).Add(f.Name)

                    'open Archive
                    Dim archive = ArchiveFactory.Open(f)
                    'sort the entries
                    Dim sortedEntries = archive.Entries.OrderBy(Function(x) x.FilePath.Length)
                    For Each entry As var In sortedEntries
                        If entry.IsDirectory AndAlso Not Directory.Exists(dir.FullName + "\" + entry.FilePath) Then
                            'create sub-directory
                            dir.CreateSubdirectory(entry.FilePath)
                        ElseIf Not File.Exists(dir.FullName + "\" + entry.FilePath) Then
                            'extract the file
                            entry.WriteToFile(dir.FullName + "\" + entry.FilePath, ExtractOptions.Overwrite)
                        End If
                    Next
                End If
            Next
        End Sub
    End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================

