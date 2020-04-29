Option Strict On
Option Explicit On
Public Class PortAgentList
    Inherits Collections.Generic.Dictionary(Of Integer, PortAgentItem)

    'Public Sub Load()

    '    Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
    '    Dim pobjComm As New SqlClient.SqlCommand
    '    Dim pobjReader As SqlClient.SqlDataReader
    '    Dim pobjClass As OSMEmailItem

    '    pobjConn.Open()
    '    pobjComm = pobjConn.CreateCommand

    '    With pobjComm
    '        .CommandType = CommandType.Text
    '        .CommandText = "SELECT osmeId  
    '                              ,ISNULL(osmeVessel_FK,0) AS osmeVessel_FK  
    '                              ,ISNULL(osmeName, '') AS osmeName  
    '                              ,ISNULL(osmeDetails, '') AS osmeDetails  
    '                              ,ISNULL(osmeType, '') AS osmeType  
    '                              ,ISNULL(osmeEmail, '') AS osmeEmail  
    '                              ,ISNULL(osmvVesselName, '') AS osmvVesselName  
    '                          FROM AmadeusReports.dbo.osmEmailDetails  
    '                         LEFT JOIN AmadeusReports.dbo.osmVessels  
    '                           ON osmeVessel_FK = osmVessels.osmvID  
    '                          WHERE osmeType = 'AGENT'  
    '                          ORDER BY  osmeName"
    '        pobjReader = .ExecuteReader
    '    End With

    '    MyBase.Clear()
    '    With pobjReader
    '        Do While .Read
    '            pobjClass = New OSMEmailItem
    '            pobjClass.SetValues(CInt(.Item("osmeId")), CInt(.Item("osmeVessel_FK")), CStr(.Item("osmeName")), CStr(.Item("osmeDetails")), CStr(.Item("osmeType")), CStr(.Item("osmeEmail")), CStr(.Item("osmvVesselName")))
    '            MyBase.Add(pobjClass.Id, pobjClass)
    '        Loop
    '        .Close()
    '    End With
    '    pobjConn.Close()

    'End Sub
    Public Sub Load()
        Dim pobjClass As PortAgentItem

        pobjClass = New PortAgentItem()
        MyBase.Add(pobjClass.Id, pobjClass)

    End Sub
End Class