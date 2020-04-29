Option Strict On
Option Explicit On
Public Class SegItem
    Public ReadOnly Property ElementNo As Integer
    Public ReadOnly Property Flown As Boolean = False
    Public ReadOnly Property Airline As String
    Public ReadOnly Property FlightNumber As String
    Public ReadOnly Property ClassOfService As String
    Public ReadOnly Property FlightDate As String
    Public ReadOnly Property Weekday As Integer
    Public ReadOnly Property BoardPoint As String
    Public ReadOnly Property OffPoint As String
    Public ReadOnly Property Status As String
    Public ReadOnly Property DepartureTime As String
    Public ReadOnly Property ArrivalTime As String
    Public ReadOnly Property ArrivalDate As String
    Public ReadOnly Property RawText As String = ""
    Public Sub New(ByVal pRawText As String)
        RawText = pRawText.TrimEnd
        ElementNo = CInt(RawText.Substring(0, 3))
        '  8  FB 808 R 27APR 1 ATHSOF         FLWN                                       
        '  9  FB 431 S 28APR 2 SOFCDG HK7  0710 0910  28APR  E  FB/UZU6RE                
        ' 10  B2 866 M 28APR 2 CDGMSQ HK7  1115 1515  28APR  E  B2/JVOAJI                
        '0123456789012345678901234567890123456789012345678901234567890123
        Airline = RawText.Substring(5, 2)
        FlightNumber = RawText.Substring(7, 4)
        ClassOfService = RawText.Substring(12, 1)
        FlightDate = RawText.Substring(14, 5)
        Weekday = CInt(RawText.Substring(20, 1))
        BoardPoint = RawText.Substring(22, 3)
        OffPoint = RawText.Substring(25, 3)
        If RawText.Substring(37).StartsWith("FLWN") Then
            Flown = True
        Else
            Flown = False
            Status = RawText.Substring(29, 4)
            DepartureTime = RawText.Substring(34, 4)
            ArrivalTime = RawText.Substring(39, 4)
            ArrivalDate = RawText.Substring(45, 5)
        End If
    End Sub
End Class
