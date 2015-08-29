Imports System.IO

Module Settings

    Public Class MySettings
        '  Public ID As Integer
        Public nm As String
        Public DVG As String
        Public Name As String
        Public Text As String
        Public Req As Boolean
        Public ShowVar As Boolean
        Public PtsPerError As Decimal
        Public MaxPts As Decimal
        Public Feedback As String
    End Class

    Public Settings As New List(Of MySettings)

    Public Sub LoadCfgFile(fn As String)
        Dim s As String
        Dim ss() As String
        Dim sr As StreamReader
        Dim sw As StreamWriter


        ' This reads the config file into the ns array, but where does it save the settings?????   jhg
        Try
            Settings.Clear()                            ' delete existing settings

            sr = File.OpenText(fn)
            s = sr.ReadLine
            If s <> "AppGrader Assignment Configuration File" Then
                MessageBox.Show("The selected file (" & fn & ") is not a valid configuration file for this application. Either load a different file, or use the default settings.")
            Else
                sw = File.CreateText(Application.StartupPath & "\Settings.txt")
                sw.WriteLine("DVG" & vbTab & "Name" & vbTab & "Text" & vbTab & "Req" & vbTab & "Show Var" & vbTab & "Pts Per Error" & vbTab & "Maxpts" & vbTab & "Feedback")

                Do While sr.Peek <> -1
                    s = sr.ReadLine
                    ss = s.Split(CChar(vbTab))

                    If ss.GetUpperBound(0) < 8 Then
                        ReDim Preserve ss(8)
                    End If

                    Dim ns As New MySettings

                    If ss.Length = 2 Then
                        ns.DVG = ss(0)
                        ns.Name = ss(0)
                        ns.Text = ss(1)
                        ns.Req = Nothing
                        ns.ShowVar = Nothing
                        ns.PtsPerError = -1
                        ns.MaxPts = -1
                        ns.Feedback = ss(8)
                    ElseIf ss.Length >= 9 Then
                        ns.DVG = ss(0)
                        ns.nm = ss(1)
                        ns.Name = ss(2)
                        ns.Text = ss(3)
                        If ss(4) = "" Then
                            ns.Req = False
                        Else
                            ns.Req = CBool(ss(4))
                        End If

                        If ss(5) = "" Then
                            ns.ShowVar = False
                        Else
                            ns.ShowVar = CBool(ss(5))
                        End If

                        ns.PtsPerError = CDec(ss(6))
                        ns.MaxPts = CDec(ss(7))
                        ns.Feedback = ss(8)
                    End If
                    sw.WriteLine(ns.DVG & vbTab & ns.Name & vbTab & ns.Req & vbTab & ns.ShowVar & vbTab & ns.PtsPerError & vbTab & ns.MaxPts & vbTab & ns.Feedback)
                    Settings.Add(ns)
                Loop
                sw.Close()
            End If

        Catch ex As Exception
            MessageBox.Show("Error Loading Config file. " & ex.Message, "Error Loading Config File")
        End Try
    End Sub


    Public Function Find_Setting(nm As String, calledfrom As String) As MySettings

        Dim p As New MySettings
        Dim p2 As New MySettings

        Dim sw As StreamWriter

        p = Settings.Find(Function(item As MySettings) item.nm = nm)
        If p Is Nothing Then
            '  MessageBox.Show("Find_Setting could not find <" & nm & ">")
            sw = File.CreateText(Application.StartupPath & "\CantFind.txt")

            sw.WriteLine("Find_Setting could not find <" & nm & "> called from " & calledfrom)
            sw.Close()

            p2.ShowVar = False
            p2.Req = False

            p2.MaxPts = 0
            Return p2
        Else
            Return p
        End If
    End Function

End Module
