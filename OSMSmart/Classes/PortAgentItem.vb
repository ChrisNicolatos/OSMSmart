Option Strict On
Option Explicit On
Public Class PortAgentItem

    Private Structure Classprops
        Dim Id As Integer
        Dim Vessel_FK As Integer
        Dim Name As String
        Dim Details As String
        Dim EmailType As String
        Dim Email As String
        Dim VesselName As String
        Dim isNew As Boolean
        Dim isValid As Boolean
    End Structure
    Dim mudtprops As Classprops

    Public Sub New()

        With mudtprops
            .Id = 0
            .Vessel_FK = 0
            .Name = ""
            .Details = ""
            .EmailType = ""
            .Email = ""
            .VesselName = ""
            .isNew = True
        End With

    End Sub
    Public Sub New(ByVal pType As String)

        With mudtprops
            .Id = 0
            .Vessel_FK = 0
            .Name = ""
            .Details = ""
            .EmailType = pType
            .Email = ""
            .VesselName = ""
            .isNew = True
            CheckValid()
        End With

    End Sub
    Public Sub New(ByVal pType As String, ByVal pVessel_FK As Integer)

        With mudtprops
            .Id = 0
            .Vessel_FK = pVessel_FK
            .Name = ""
            .Details = ""
            .EmailType = pType
            .Email = ""
            .VesselName = ""
            .isNew = True
            CheckValid()
        End With

    End Sub
    Private Sub CheckValid()

        With mudtprops
            .isValid = False
            If (.Vessel_FK <> 0 Or (.EmailType = "AGENT" And .Vessel_FK = 0)) And .Name <> "" And .EmailType <> "" And .Email <> "" Then
                .isValid = True
            End If
        End With
    End Sub
    Public Overrides Function ToString() As String

        ToString = Chr(34) & mudtprops.Name & Chr(34) & " <" & mudtprops.Email & ">"
        Dim pEmail() As String = mudtprops.Email.Split({";"}, StringSplitOptions.RemoveEmptyEntries)
        If IsArray(pEmail) Then
            Dim pString As String = ""
            For i As Integer = 0 To pEmail.GetUpperBound(0)

                If pString <> "" Then
                    pString &= ";"
                End If
                pString &= Chr(34) & mudtprops.Name & Chr(34) & " <" & pEmail(i) & ">"
            Next
            ToString = pString
        End If

    End Function

    Public ReadOnly Property Id As Integer
        Get
            Return mudtprops.Id
        End Get
    End Property
    Public ReadOnly Property Vessel_FK As Integer
        Get
            Return mudtprops.Vessel_FK
        End Get
    End Property
    Public Property Name As String
        Get
            Return mudtprops.Name
        End Get
        Set(value As String)
            mudtprops.Name = value
            CheckValid()
        End Set
    End Property
    Public Property Details As String
        Get
            Return mudtprops.Details
        End Get
        Set(value As String)
            mudtprops.Details = value
            CheckValid()
        End Set
    End Property
    Public ReadOnly Property EmailType As String
        Get
            Return mudtprops.EmailType
        End Get
    End Property
    Public Property Email As String
        Get
            Return mudtprops.Email
        End Get
        Set(value As String)
            mudtprops.Email = value
            CheckValid()
        End Set
    End Property
    Public ReadOnly Property VesselName As String
        Get
            Return mudtprops.VesselName
        End Get
    End Property
    Public ReadOnly Property isValid As Boolean
        Get
            Return mudtprops.isValid
        End Get
    End Property
    Public ReadOnly Property isNew As Boolean
        Get
            Return mudtprops.isNew
        End Get
    End Property
    Public Sub SetValues(ByVal pId As Integer, ByVal pVessel_FK As Integer, ByVal pName As String, ByVal pDetails As String, ByVal pEmailType As String, ByVal pEmail As String, ByVal pVesselName As String)

        With mudtprops
            .Id = pId
            .Vessel_FK = pVessel_FK
            .Name = pName
            .Details = pDetails
            .EmailType = pEmailType
            .Email = pEmail
            .VesselName = pVesselName
            .isNew = False
            CheckValid()
        End With
    End Sub
End Class