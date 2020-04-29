Option Strict On
Option Explicit On
Imports System.Text.RegularExpressions
Public Class PNR
    Private mRPIndex As Integer = -1
    Private PNRLines() As String
    Public ReadOnly Property PNRRaw As String
    Public ReadOnly Property PNRCode As String
    Public Property Pax As PaxList
    Public Property Seg As SegList
    Public Sub New(ByVal pPNRRaw As String)
        PNRRaw = pPNRRaw
        PNRLines = PNRRaw.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        FindRPLine()
        FindPaxSegs()
    End Sub
    Private Sub FindRPLine()
        For i As Integer = 0 To PNRLines.GetUpperBound(0)
            If PNRLines(i).StartsWith("RP/") Then
                If mRPIndex >= 0 Then
                    Throw New Exception("Input contains more than one RP/ line")
                End If
                mRPIndex = i
            End If
        Next
        If mRPIndex = -1 Then
            Throw New Exception("Input does not contain RP/ line")
        End If
    End Sub
    Private Sub FindPaxSegs()
        Dim pPaxPattern As String = "^[ \d]{2}[\d]\.[A-Z].*$"
        Dim pSegPattern As String = "^[ \d]{2}[\d]\s{2}[A-Z].*$"
        Pax = New PaxList
        Seg = New SegList
        For i As Integer = mRPIndex + 1 To PNRLines.GetUpperBound(0)
            If Regex.IsMatch(PNRLines(i), pPaxPattern) Then
                Pax.AddPNRPax(PNRLines(i))
            ElseIf Regex.IsMatch(PNRLines(i), pSegPattern) Then
                Seg.AddPNRPax(PNRLines(i))
            End If
        Next
    End Sub
End Class
