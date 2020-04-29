Option Strict On
Option Explicit On
Public Class SegList
    Inherits Collections.Generic.Dictionary(Of Integer, SegItem)
    Public Sub AddPNRPax(ByVal SegLine As String)
        Dim pItem As New SegItem(SegLine)
        MyBase.Add(pItem.ElementNo, pItem)
    End Sub
End Class
