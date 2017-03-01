'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "cBinaryFileStream"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit
Public Class cBinaryFileStream


    Private m_sFile As String
    Private m_iFile As Integer
    Private m_iLen As Long
    Private m_iOffset As Long

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As VariantType, lpvSource As VariantType, ByVal cbCopy As Long)

    Property File As String = m_sFile

    'Public Property Get File() As String
    '   File = m_sFile
    'End Property
    'Public Property Let File(ByVal sFile As String)
    '   Dispose
    '   m_sFile = sFile
    '    Dim lErr As Long

    '   If (FileExists(m_sFile, lErr)) Then
    '      m_iFile = FreeFile
    '      Open m_sFile For Binary Access Read Lock Write As #m_iFile
    '      m_iLen = LOF(m_iFile)
    '   Else
    '      Err.Raise lErr, App.EXEName & ".File"
    '   End If

    'End Property

    Private Function FileExists(ByVal sFile As String, ByRef lErr As Long) As Boolean

        lErr = 0
        On Error Resume Next
        Dim sDir As String
        sDir = Dir(sFile)
        lErr = Err.Number
        On Error GoTo 0

        If (lErr = 0) Then
            If (Len(sDir) > 0) Then
                FileExists = True
            Else
                lErr = 53
            End If
        End If

    End Function

    Property Length As Long = m_iLen
    'Public Property Get Length() As Long
    '   Length = m_iLen
    'End Property

    Public Function Read(buffer() As Byte, ByVal readSize As Long) As Long

        Dim lReadSize As Long
        Dim newBuffer() As Byte

        lReadSize = readSize

        If (m_iOffset + lReadSize >= m_iLen) Then
            readSize = m_iLen - m_iOffset
            If (readSize > 0) Then
                ReDim newBuffer(0 To CInt(readSize - 1))

         Get #m_iFile, , newBuffer
                CopyMemory(buffer(0), newBuffer(0), readSize)
            Else
                Dispose()
            End If
            m_iOffset = m_iOffset + readSize
        Else
            ' Can read
      Get #m_iFile, , buffer
            m_iOffset = m_iOffset + readSize
        End If
        Read = readSize

    End Function

    Public Sub Dispose()
        If (m_iFile) Then
      Close #m_iFile
            m_iFile = 0
        End If
    End Sub

    Private Sub Class_Terminate()
        Dispose()
    End Sub
End Class