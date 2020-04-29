Option Strict On
Option Explicit On
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO

Public Class OSMElements
    Public Shared Function MakePaxTable(ByRef pPassengers As PaxList, ByVal pFont As Font, ByVal pHeaderFont As Font) As PdfPTable

        Dim Table As New PdfPTable(2) With {
            .LockedWidth = False,
            .HorizontalAlignment = 0,
            .SpacingBefore = 14,
            .SpacingAfter = 14
        }

        Dim pPosition As Boolean = False
        For Each pPax As PaxItem In pPassengers.Values
            If pPax.Remarks <> "" Then
                pPosition = True
                Exit For
            End If
        Next pPax
        'relative col widths in proportions - 2/3 And 1/3
        Dim widths() As Single = {2, 1}
        Table.SetWidths(widths)
        Table.AddCell(AddCell("Name", pHeaderFont))
        If pPosition Then
            Table.AddCell(AddCell("Position", pHeaderFont))
        Else
            Table.AddCell(AddCell(" ", pHeaderFont))
        End If

        For Each pPax As PaxItem In pPassengers.Values
            Table.AddCell(AddCell(pPax.PaxName, pFont))
            Table.AddCell(AddCell(pPax.Remarks, pFont))
        Next pPax

        MakePaxTable = Table

    End Function
    Public Shared Function MakePaxTable(ByRef pPax As PaxItem, ByVal pFont As Font, ByVal pHeaderFont As Font) As PdfPTable

        Dim Table As New PdfPTable(2) With {
            .LockedWidth = False,
            .HorizontalAlignment = 0,
            .SpacingBefore = 14,
            .SpacingAfter = 14
        }

        'relative col widths in proportions - 2/3 And 1/3
        Dim widths() As Single = {1, 1}
        Table.SetWidths(widths)

        Table.AddCell(AddCell("Name", pHeaderFont))
        If pPax.Remarks = "" Then
            Table.AddCell(AddCell(" ", pHeaderFont))
        Else
            Table.AddCell(AddCell("Position", pHeaderFont))
        End If
        Table.AddCell(AddCell(pPax.PaxName, pFont))
        Table.AddCell(AddCell(pPax.Remarks, pFont))

        MakePaxTable = Table

    End Function
    Public Shared Function AddCell(ByVal pText As String, ByVal pFont As Font) As PdfPCell
        Dim c1 As New PdfPCell(New Phrase(pText, pFont)) With {
                    .Border = Rectangle.NO_BORDER
                }
        AddCell = c1
    End Function
End Class
