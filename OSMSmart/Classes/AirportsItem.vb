Option Strict On
Option Explicit On
Public Class AirportsItem
    Public ReadOnly Property AirportCode As String
    Public ReadOnly Property AirportName As String
    Public ReadOnly Property CityName As String
    Public ReadOnly Property Country As String
    Public Sub New(ByVal pAirpoprtCode As String, ByVal pAirportName As String, ByVal pCityName As String, ByVal pCountry As String)
        AirportCode = pAirpoprtCode
        AirportName = pAirportName
        CityName = pCityName
        Country = pCountry
    End Sub
End Class
