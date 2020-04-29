Option Strict On
Option Explicit On
Public Class PaxList
    Inherits Collections.Generic.Dictionary(Of Integer, PaxItem)
    Public Sub AddPNRPax(ByVal PaxLine As String)
        Dim pItem As New PaxItem(PaxLine)
        MyBase.Add(pItem.ElementNo, pItem)
    End Sub
End Class
