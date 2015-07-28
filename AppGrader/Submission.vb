Option Strict On
Option Explicit On

Public Class Submission

    'Structure crcdatum
    '    Dim UserID As String
    '    Dim Filename As String
    '    Dim vbCRC As String
    'End Structure

    'City UserID, vbCRC, Filename properties
    Private _UserID As String
    Public Property UserID() As String
        Get
            Return _UserID
        End Get
        Set(ByVal value As String)
            _UserID = value
        End Set
    End Property

    Private _Filename As String
    Public Property Filename() As String
        Get
            Return _Filename
        End Get
        Set(ByVal value As String)
            _Filename = value
        End Set
    End Property

    Private _vbCRC As String
    Public Property vbCRC() As String
        Get
            Return _vbCRC
        End Get
        Set(ByVal value As String)
            _vbCRC = value
        End Set
    End Property
    'Public Property vbMD5() As String
    '    Get
    '        Return _vbMD5
    '    End Get
    '    Set(ByVal value As String)
    '        _vbMD5 = value
    '    End Set
    'End Property
End Class

