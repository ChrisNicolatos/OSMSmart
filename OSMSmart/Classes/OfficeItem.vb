Option Strict On
Option Explicit On
Imports System.IO
Public Class OfficeItem
    Public Overrides Function ToString() As String
        Return Office
    End Function
    Public Property Id As Integer
    Public Property Office As String
    Public Property SignedByName As String
    Public Property Title As String
    Public Property CompanyName As String
    Public Property Address As String
    Public Property PCArea As String
    Public Property Country As String
    Public Property Telephone As String
    Public Property SignatorysExpenses As Boolean
    Public Property LogoImage_fk As Integer
    Public Property SignatureImage_fk As Integer
    Public ReadOnly Property LogoImageFromStream As Image
        Get
            Return Image.FromStream(New MemoryStream(LogoImage))
        End Get
    End Property
    Public ReadOnly Property LogoImage As Byte()
    Public ReadOnly Property SignatureImageFromStream As Image
        Get
            Return Image.FromStream(New MemoryStream(SignatureImage))
        End Get
    End Property
    Public ReadOnly Property SignatureImage As Byte()
    Public Sub New(ByVal pId As Integer, ByVal pOffice As String, ByVal pSignedByName As String, ByVal pTitle As String _
                       , ByVal pCompanyName As String, ByVal pAddress As String, ByVal pPCArea As String, ByVal pCountry As String _
                       , ByVal pTelephone As String, ByVal pLogoImage_fk As Integer, ByVal pSignatureImage_fk As Integer, ByVal pSignatorysExpenses As Boolean)
        Id = pId
        Office = pOffice
        SignedByName = pSignedByName
        Title = pTitle
        CompanyName = pCompanyName
        Address = pAddress
        PCArea = pPCArea
        Country = pCountry
        Telephone = pTelephone
        LogoImage_fk = pLogoImage_fk
        SignatureImage_fk = pSignatureImage_fk
        If LogoImage_fk <> 0 Then
            'LogoImage = LoadImage(LogoImage_fk)
        End If
        If SignatureImage_fk <> 0 Then
            'SignatureImage = LoadImage(SignatureImage_fk)
        End If
        SignatorysExpenses = pSignatorysExpenses

    End Sub
    'Private Shared Function LoadImage(ByVal pImage_FK As Integer) As Byte()
    '    Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
    '    Dim pobjComm As New SqlClient.SqlCommand

    '    pobjConn.Open()
    '    pobjComm = pobjConn.CreateCommand
    '    With pobjComm
    '        .CommandType = CommandType.Text
    '        .Parameters.Add("@Id", SqlDbType.Int).Value = pImage_FK
    '        .CommandText = "SELECT imageData 
    '                        FROM [AmadeusReports].[dbo].[Images] 
    '                        WHERE imageId = @Id"
    '        Return CType(.ExecuteScalar, Byte())
    '    End With

    '    pobjConn.Close()
    'End Function
End Class
