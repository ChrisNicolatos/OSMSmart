Option Strict On
Option Explicit On
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO
Public Class OSMLog
    Public Enum EnumLoGLanguage
        English = 0
        Brazil = 1
    End Enum

    Private mobjPNR As PNR
    Public ReadOnly Property LogPerPax As Boolean
    Public ReadOnly Property LogLanguage As EnumLoGLanguage ' 0 = English, 1 = Brazil

    Private mobjPortAgent As PortAgentItem
    Private mflgNoPortAgent As Boolean
    Private mobjAddressItem As OfficeItem
    Public Sub New(ByRef pPNR As PNR, ByVal pLogPerPax As Boolean, ByVal pLogLanguage As EnumLoGLanguage, ByVal LogPath As String, ByVal LogOnsigner As Boolean)
        mobjPNR = pPNR
        LogPerPax = pLogPerPax
        LogLanguage = pLogLanguage
        ReadOptions(LogPath, LogOnsigner)
    End Sub
    Private Function CreateDocs(ByVal LogPath As String, ByVal LogOnsigner As Boolean) As String
        Dim pFileName As String = ""
        Dim pstrTextCrewMembers As String = "crew member listed below is scheduled" ' default text is for one pax
        CreateDocs = ""
        If LogPerPax Then
            For Each pPax As PaxItem In mobjPNR.Pax.Values
                pFileName = GetPDFFileName(mobjPNR.PNRCode & "-" & pPax.ElementNo & pPax.LastName, LogPath)
                MakePDFDocument(LogLanguage, pFileName, pstrTextCrewMembers, LogOnsigner, pPax)
                CreateDocs &= pFileName & vbCrLf
            Next pPax
        Else
            If mobjPNR.Pax.Count > 1 Then
                pstrTextCrewMembers = "crew members listed below are scheduled"
            End If
            pFileName = GetPDFFileName(mobjPNR.PNRCode, LogPath)
            MakePDFDocument(LogLanguage, pFileName, pstrTextCrewMembers, LogOnsigner)
            CreateDocs = pFileName
        End If
    End Function
    Private Sub MakePDFDocument(ByVal LoGLanguage As Integer, ByVal pFileName As String, ByVal CrewMembersText As String, ByVal LogOnsigner As Boolean, Optional ByRef pPax As PaxItem = Nothing)
        Select Case LoGLanguage
            Case 1
                PDFDocumentLangBrazil(pFileName, pPax)
            Case Else
                PDFDocument(pFileName, CrewMembersText, LogOnsigner, pPax)
        End Select
    End Sub
    Private Sub PDFDocument(ByVal pFileName As String, ByVal CrewMembersText As String, ByVal LogOnsigner As Boolean, Optional ByRef pPax As PaxItem = Nothing)

        Dim pLogoImage As Image = Image.GetInstance(mobjAddressItem.LogoImage)
        Dim pDoc As New Document(PageSize.A4, 36, 36, 36, 36)
        Dim pArial11 As Font = FontFactory.GetFont("arial", 11, FontStyle.Regular)
        Dim pArial11b As Font = FontFactory.GetFont("arial", 11, FontStyle.Bold)
        Dim pArial12 As Font = FontFactory.GetFont("arial", 12, FontStyle.Regular)
        Dim pArial16b As Font = FontFactory.GetFont("arial", 16, FontStyle.Bold)

        Dim pExpensesBy As String = ""

        If mobjAddressItem.SignatorysExpenses Then
            pExpensesBy = mobjAddressItem.CompanyName
        Else
            pExpensesBy = mobjPNR.ExistingElements.ClientName
        End If

        PdfWriter.GetInstance(pDoc, New FileStream(pFileName, FileMode.Create))
        pDoc.Open()
        pLogoImage.ScalePercent(40)
        pDoc.Add(pLogoImage)

        pDoc.Add(AddParagraph("LETTER OF GUARANTEE", pArial16b, 14, 14, "Center"))
        pDoc.Add(AddParagraph(Format(Now, "dd/MM/yyyy"), pArial12, 0, 14, "Right"))

        pDoc.Add(AddParagraph("To Whom It May Concern", pArial11, 0, 14, "Left"))

        If LogOnsigner Then
            pDoc.Add(AddParagraph("We hereby declare that our " & CrewMembersText & " to arrive on " & mobjPNR.Seg.Values.Last.Arrival.SegDate & " to embark the vessel/rig " & mobjPNR.ExistingElements.VesselName & " in " & mobjPNR.Seg.Values.Last.Destination.CityName & If(mobjPNR.Seg.Values.Last.Destination.CountryName <> "", ", " & mobjPNR.Seg.Values.Last.Destination.CountryName, ""), pArial11, 0, 6, "Left"))
        Else
            pDoc.Add(AddParagraph("We hereby declare that our " & CrewMembersText & " to depart on " & mobjPNR.Seg.First.Value.Departure.SegDate & " to disembark the vessel/rig " & mobjPNR.ExistingElements.VesselName & " in " & mobjPNR.Seg.First.Value.Origin.CityName & If(mobjPNR.Seg.First.Value.Origin.CountryName <> "", ", " & mobjPNR.Seg.First.Value.Origin.CountryName, ""), pArial11, 0, 6, "Left"))
        End If
        pDoc.Add(AddParagraph("By carrying this letter, the crew is entitled to travel on maritime/offshore fares.", pArial11, 0, 0, "Left"))

        If pPax Is Nothing Then
            pDoc.Add(OSMElements.MakePaxTable(mobjPNR.Pax, pArial11, pArial11b))
        Else
            pDoc.Add(OSMElements.MakePaxTable(pPax, pArial11, pArial11b))
        End If

        pDoc.Add(AddParagraph("Travel Itinerary (subject to change):", pArial11b, 0, 0, "Left"))
        pDoc.Add(MakeSegTable(mobjPNR.Seg, pArial11))

        If Not mflgNoPortAgent And Not mobjPortAgent Is Nothing Then
            pDoc.Add(AddParagraph("PORT AGENT", pArial11b, 7, 0, "Left"))
            pDoc.Add(AddParagraph(mobjPortAgent.Name, pArial11, 0, 0, "Left"))
            pDoc.Add(AddParagraph(mobjPortAgent.Details, pArial11, 0, 0, "Left"))
            pDoc.Add(AddParagraph(mobjPortAgent.Email, pArial11, 0, 7, "Left"))
        End If
        pDoc.Add(AddParagraph("We ask you kindly to render all necessary assistance.", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("We confirm that " & pExpensesBy & " will cover all expenses that may occur in connection with our employee's travel.", pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph("If you need any further information, please contact our employer as stated below.", pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph("Sincerely,", pArial11, 0, 14, "Left"))
        If mobjAddressItem.SignatureImage_fk <> 0 Then
            Dim pSignatureImage As Image = Image.GetInstance(mobjAddressItem.SignatureImage)
            pSignatureImage.ScalePercent(50)
            pDoc.Add(pSignatureImage)
        End If

        pDoc.Add(AddParagraph(mobjAddressItem.SignedByName, pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph(mobjAddressItem.Title, pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph("On behalf of", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph(mobjAddressItem.CompanyName, pArial11b, 0, 0, "Left"))
        pDoc.Add(AddParagraph("Address: " & mobjAddressItem.Address & " " & mobjAddressItem.PCArea & " " & mobjAddressItem.Country, pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("Phone: " & mobjAddressItem.Telephone, pArial11, 0, 0, "Left"))


        pDoc.Close()

    End Sub
    Private Sub PDFDocumentLangBrazil(ByVal pFileName As String, Optional ByRef pPax As PaxItem = Nothing)

        Dim pLogoImage As Image = Image.GetInstance(mobjAddressItem.LogoImage)
        Dim pDoc As New Document(PageSize.A4, 36, 36, 36, 36)
        Dim pArial11 As Font = FontFactory.GetFont("arial", 11, FontStyle.Regular)
        Dim pArial11b As Font = FontFactory.GetFont("arial", 11, FontStyle.Bold)
        Dim pArial12 As Font = FontFactory.GetFont("arial", 12, FontStyle.Regular)

        PdfWriter.GetInstance(pDoc, New FileStream(pFileName, FileMode.Create))
        pDoc.Open()
        pLogoImage.ScalePercent(40)
        pDoc.Add(pLogoImage)

        pDoc.Add(AddParagraph("A QUEM POSSA INTERESSAR", pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph(mobjPNR.ExistingElements.VesselName, pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph("Atenas," & MyMonthName(Now, EnumLoGLanguage.Brazil), pArial12, 0, 14, "Right"))

        pDoc.Add(AddParagraph("Isso é para aconselhá-lo, que o representante do escritório o seguinte vai velejar com o navio legenda em " & mobjPNR.Seg.Values.Last.Destination.CityName & If(mobjPNR.Seg.Values.Last.Destination.CountryName <> "", ", " & mobjPNR.Seg.Values.Last.Destination.CountryName, ""), pArial11, 0, 6, "Left"))

        If pPax Is Nothing Then
            pDoc.Add(MakePaxTableLangBrazil(mobjPNR.Pax, pArial11))
        Else
            pDoc.Add(MakePaxTableLangBrazil(pPax, pArial11))
        End If

        pDoc.Add(AddParagraph("Detalhes do vôo :", pArial11b, 7, 0, "Left"))
        pDoc.Add(MakeSegTable(mobjPNR.Seg, pArial11))

        If Not mflgNoPortAgent And Not mobjPortAgent Is Nothing Then
            pDoc.Add(AddParagraph("Agente Portuário", pArial11b, 7, 0, "Left"))
            pDoc.Add(AddParagraph(mobjPortAgent.Name, pArial11, 0, 0, "Left"))
            pDoc.Add(AddParagraph(mobjPortAgent.Details, pArial11, 0, 0, "Left"))
            pDoc.Add(AddParagraph(mobjPortAgent.Email, pArial11, 0, 7, "Left"))
        End If
        pDoc.Add(AddParagraph("Por favor, dê-lhe toda a assistência possível, a fim de que ele deve chegar ao seu destino com o mínimo de atraso possível. Nós garantir que, no entanto, que será responsável para pagar todas as suas despesas de repatriamento, se as Autoridades Portuárias recusar a sua entrada no " & mobjPNR.ExistingElements.VesselName & " por qualquer motivo.", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("Agradecendo antecipadamente.", pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph("Com os melhores cumprimentos.", pArial11, 0, 14, "Left"))
        If mobjAddressItem.SignatureImage_fk <> 0 Then
            Dim pSignatureImage As Image = Image.GetInstance(mobjAddressItem.SignatureImage)
            pSignatureImage.ScalePercent(50)
            pDoc.Add(pSignatureImage)
        End If
        pDoc.Add(AddParagraph(mobjAddressItem.SignedByName, pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph(mobjAddressItem.Title, pArial11, 0, 14, "Left"))
        pDoc.Add(AddParagraph("On behalf of", pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph(mobjAddressItem.CompanyName, pArial11b, 0, 0, "Left"))
        pDoc.Add(AddParagraph("Address: " & mobjAddressItem.Address & " " & mobjAddressItem.PCArea & " " & mobjAddressItem.Country, pArial11, 0, 0, "Left"))
        pDoc.Add(AddParagraph("Phone: " & mobjAddressItem.Telephone, pArial11, 0, 0, "Left"))
        pDoc.Close()

    End Sub
    Private Shared Function GetPDFFileName(ByVal FileNameDetails As String, ByVal LogPath As String) As String
        GetPDFFileName = System.IO.Path.Combine(LogPath, FileNameDetails & ".pdf")
        Dim pTemp As Integer = 0
        Do While System.IO.File.Exists(GetPDFFileName)
            pTemp += 1
            GetPDFFileName = System.IO.Path.Combine(LogPath, FileNameDetails & "-" & pTemp & ".pdf")
        Loop

    End Function
    Private Shared Function AddParagraph(ByVal pText As String, ByVal pFont As Font, ByVal pSpacingBefore As Integer, ByVal pSpacingAfter As Integer, ByVal pAlignment As String) As Paragraph

        Dim x2 As New Paragraph(pText, pFont) With {
            .SpacingBefore = pSpacingBefore,
            .SpacingAfter = pSpacingAfter
        }
        x2.SetAlignment(pAlignment)
        AddParagraph = x2

    End Function
    Private Function MakePaxTableLangBrazil(ByRef pPassengers As PaxList, ByVal pFont As Font) As PdfPTable

        Dim Table As New PdfPTable(2) With {
            .LockedWidth = False,
            .HorizontalAlignment = 0,
            .SpacingBefore = 14,
            .SpacingAfter = 14
        }
        For Each pPax As PaxItem In pPassengers.Values
            If pPax.Remarks <> "" Then
                Exit For
            End If
        Next pPax
        'relative col widths in proportions - 1/2 And 1/2
        Dim widths() As Single = {1, 1}
        Table.SetWidths(widths)

        For Each pPax As PaxItem In pPassengers.Values
            'Sobrenome: Carlos Naia
            Table.AddCell(OSMElements.AddCell("Sobrenome:", pFont))
            Table.AddCell(OSMElements.AddCell(pPax.LastName, pFont))
            'Nome:   Hugo Miguel
            Table.AddCell(OSMElements.AddCell("Nome:", pFont))
            Table.AddCell(OSMElements.AddCell(pPax.FirstName, pFont))
            Table.AddCell(OSMElements.AddCell("Posição:", pFont))
            Table.AddCell(OSMElements.AddCell(pPax.Remarks, pFont))
            If mobjPNR.ExistingElements.SSRDocsExists Then
                For Each pDocs As ApisPaxItem In mobjPNR.ExistingElements.SSRDocsCollection.Values
                    If pDocs.Surname.Replace(" ", "") = pPax.LastName.Replace(" ", "") And pPax.FirstName.Replace(" ", "").StartsWith(pDocs.FirstName.Replace(" ", "")) Then
                        'Nacionalidade: Portuguese
                        Table.AddCell(OSMElements.AddCell("Nacionalidade:", pFont))
                        Table.AddCell(OSMElements.AddCell(pDocs.Nationality, pFont))
                        'DOB:  18/03/1988
                        If pDocs.BirthDate <> Date.MinValue Then
                            Table.AddCell(OSMElements.AddCell("DOB:", pFont))
                            Table.AddCell(OSMElements.AddCell(MyMonthName(pDocs.BirthDate, EnumLoGLanguage.Brazil), pFont))
                        End If
                        'Número do passaporte:  P876285
                        Table.AddCell(OSMElements.AddCell("Número do passaporte:", pFont))
                        Table.AddCell(OSMElements.AddCell(pDocs.PassportNumber, pFont))
                        'Data de validade:   06/07/2022 
                        If pDocs.ExpiryDate <> Date.MinValue Then
                            Table.AddCell(OSMElements.AddCell("Data de validade:", pFont))
                            Table.AddCell(OSMElements.AddCell(MyMonthName(pDocs.ExpiryDate, EnumLoGLanguage.Brazil), pFont))
                        End If
                    End If
                Next
            End If

            Table.AddCell(OSMElements.AddCell(" ", pFont))
            Table.AddCell(OSMElements.AddCell(" ", pFont))
        Next pPax

        Return Table

    End Function
    Private Function MakePaxTableLangBrazil(ByRef pPax As PaxItem, ByVal pFont As Font) As PdfPTable

        Dim Table As New PdfPTable(2) With {
            .LockedWidth = False,
            .HorizontalAlignment = 0,
            .SpacingBefore = 14,
            .SpacingAfter = 14
        }

        'relative col widths in proportions - 1/2 And 1/2
        Dim widths() As Single = {1, 1}
        Table.SetWidths(widths)

        'Sobrenome: Carlos Naia
        Table.AddCell(OSMElements.AddCell("Sobrenome:", pFont))
        Table.AddCell(OSMElements.AddCell(pPax.LastName, pFont))
        'Nome:   Hugo Miguel
        Table.AddCell(OSMElements.AddCell("Nome:", pFont))
        Table.AddCell(OSMElements.AddCell(pPax.FirstName, pFont))
        Table.AddCell(OSMElements.AddCell("Posição:", pFont))
        Table.AddCell(OSMElements.AddCell(pPax.Remarks, pFont))
        If mobjPNR.ExistingElements.SSRDocsExists Then
            For Each pDocs As ApisPaxItem In mobjPNR.ExistingElements.SSRDocsCollection.Values
                If pDocs.Surname = pPax.LastName And pDocs.FirstName = pPax.FirstName Then
                    'Nacionalidade: Portuguese
                    Table.AddCell(OSMElements.AddCell("Nacionalidade:", pFont))
                    Table.AddCell(OSMElements.AddCell(pDocs.Nationality, pFont))
                    'DOB:  18/03/1988
                    If pDocs.BirthDate <> Date.MinValue Then
                        Table.AddCell(OSMElements.AddCell("DOB:", pFont))
                        Table.AddCell(OSMElements.AddCell(MyMonthName(pDocs.BirthDate, EnumLoGLanguage.Brazil), pFont))
                    End If
                    'Número do passaporte:  P876285
                    Table.AddCell(OSMElements.AddCell("Número do passaporte:", pFont))
                    Table.AddCell(OSMElements.AddCell(pDocs.PassportNumber, pFont))
                    'Data de validade:   06/07/2022 
                    If pDocs.ExpiryDate <> Date.MinValue Then
                        Table.AddCell(OSMElements.AddCell("Data de validade:", pFont))
                        Table.AddCell(OSMElements.AddCell(MyMonthName(pDocs.ExpiryDate, EnumLoGLanguage.Brazil), pFont))
                    End If
                End If
            Next
        End If
        Return Table

    End Function
    Private Function MakeSegTable(ByRef pSegs As SegList, ByVal pFont As Font) As PdfPTable

        Dim pWidths(6) As Single
        Dim pVBFont As New Drawing.Font(pFont.Familyname, pFont.Size, If(pFont.IsBold, FontStyle.Bold, FontStyle.Regular))
        Dim pfrm As New frmOSMLoG(mobjPNR)
        Dim g As Graphics = pfrm.CreateGraphics

        Dim Table As New PdfPTable(7) With {
            .LockedWidth = False,
            .HorizontalAlignment = 0,
            .SpacingBefore = 14,
            .SpacingAfter = 14
        }
        For Each pSeg As SegItem In pSegs.Values
            With pSeg

                pWidths(0) = Math.Max(pWidths(0), g.MeasureString(.Airline, pVBFont).Width)
                pWidths(1) = Math.Max(pWidths(1), g.MeasureString(.FlightNumber, pVBFont).Width)
                pWidths(2) = Math.Max(pWidths(2), g.MeasureString(.Departure.DateIATA, pVBFont).Width)
                pWidths(3) = Math.Max(pWidths(3), g.MeasureString(.Origin.CityName, pVBFont).Width)
                pWidths(4) = Math.Max(pWidths(4), g.MeasureString(.Destination.CityName, pVBFont).Width)
                pWidths(5) = Math.Max(pWidths(5), g.MeasureString(.Departure.TimeShort, pVBFont).Width)
                pWidths(6) = Math.Max(pWidths(6), g.MeasureString(.Arrival.TimeShort, pVBFont).Width)
            End With
        Next
        'relative col widths in proportions - 2/3 And 1/3

        Table.SetWidths(pWidths)

        For Each pSeg As SegItem In pSegs.Values
            With pSeg
                Table.AddCell(OSMElements.AddCell(.Airline, pFont))
                Table.AddCell(OSMElements.AddCell(.FlightNumber, pFont))
                Table.AddCell(OSMElements.AddCell(.Departure.DateIATA, pFont))
                Table.AddCell(OSMElements.AddCell(.Origin.CityName, pFont))
                Table.AddCell(OSMElements.AddCell(.Destination.CityName, pFont))
                Table.AddCell(OSMElements.AddCell(.Departure.TimeShort, pFont))
                Table.AddCell(OSMElements.AddCell(.Arrival.TimeShort, pFont))
            End With
        Next

        MakeSegTable = Table

    End Function

    Private Sub ReadOptions(ByVal LogPath As String, ByVal LogOnsigner As Boolean)

        Dim pFrm As New frmOSMLoG(mobjPNR)
        If pFrm.ShowDialog() = DialogResult.OK Then
            mflgNoPortAgent = pFrm.NoPortAgent
            mobjPortAgent = pFrm.PortAgent
            mobjAddressItem = pFrm.AddressItem
            pFrm.Close()
            Dim pStatus As String = CreateDocs(LogPath, LogOnsigner)
            MessageBox.Show(pStatus, "Create PDF File(s)", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Cancelled")
        End If

    End Sub
    Public Function MyMonthName(ByVal pDate As Date, ByVal Language As EnumLoGLanguage) As String
        Static Dim pNamesLang1() As String = {"janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"}

        If Language = EnumLoGLanguage.Brazil Then
            Return pDate.Day & " de " & pNamesLang1(pDate.Month - 1) & " de " & pDate.Year
        Else
            Return pDate.Day & " " & MonthName(pDate.Month) & " " & pDate.Year
        End If

    End Function
End Class
