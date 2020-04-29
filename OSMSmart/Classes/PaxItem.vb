Option Strict On
Option Explicit On
Public Class PaxItem
    Public ReadOnly Property ElementNo As Integer = 0
    Public ReadOnly Property LastName As String = ""
    Public ReadOnly Property FirstName As String = ""
    Public ReadOnly Property Remarks As String = ""
    Public ReadOnly Property RawText As String = ""
    Public ReadOnly Property PaxName() As String
        Get
            Return LastName & "/" & FirstName
        End Get
    End Property
    Public Sub New(ByVal pRawText As String)
        RawText = pRawText.TrimEnd
        ElementNo = CInt(RawText.Substring(0, 3))
        Dim iSep1 As Integer = RawText.IndexOf("/", 4)
        LastName = RawText.Substring(4, iSep1 - 4)
        Dim iSep2 As Integer = RawText.IndexOf("(", iSep1)
        If iSep2 = -1 Then
            FirstName = RawText.Substring(iSep1 + 1)
            Remarks = ""
        Else
            FirstName = RawText.Substring(iSep1 + 1, iSep2 - iSep1 - 1)
            Remarks = RawText.Substring(iSep2 + 1, RawText.Length - iSep2 - 2)
        End If
    End Sub
End Class
