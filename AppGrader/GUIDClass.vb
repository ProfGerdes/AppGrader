Option Strict On
Option Explicit On

Public Class GUIDData
    ' Student, GUID, FilePath properties
    Private _Student As String
    Public Property Student() As String
        Get
            Return _Student
        End Get
        Set(ByVal value As String)
            _Student = value
        End Set
    End Property

    Private _GUID As String
    Public Property GUID() As String
        Get
            Return _GUID
        End Get
        Set(ByVal value As String)
            _GUID = value
        End Set
    End Property

    Private _FilePath As String
    Public Property FilePath() As String
        Get
            Return _FilePath
        End Get
        Set(ByVal value As String)
            _FilePath = value
        End Set
    End Property

End Class

