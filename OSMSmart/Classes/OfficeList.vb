Option Strict On
Option Explicit On
Public Class OfficeList
    Inherits System.Collections.Generic.Dictionary(Of Integer, OfficeItem)
    Public Sub Load()
        Dim pobjClass As OfficeItem

        pobjClass = New OfficeItem(1, "Australia ", "Casey Munyard ", "Managing Director", "OSM Australia Pty Ltd.", "AMP Level 28, 140 St Georges Terrace", "Perth 6000 WA Australia", "Australia", "+61 497 578 079", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)
        pobjClass = New OfficeItem(2, "Brazil ", "Andreia Tinoco ", "Logistic Assistant", "OSM Maritime Group", "Rua da Assembléia, 10 – 2213 – Centro ", "Rio de Janeiro – RJ – 20011-901", "Brasil", "+5521990333076", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)
        pobjClass = New OfficeItem(3, "Croatia ", "Katarina Gudelj", "Office Manager ", "OSM Croatia", "Uvala baluni 9, ", "21000 Split, Croatia", "Croatia", "+385912610103", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)
        pobjClass = New OfficeItem(4, "Manila ", "Cherryl Rose Nemenzo ", "Head of Crewing Support", "OSM Maritime Group", "OSM Bldg., 479 Pedro Gil St., Ermita", "Manila 1000", "Philippines", "M: +63 9175593234 D: +63 2 5248879", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)
        pobjClass = New OfficeItem(5, "Poland ", "Miroslaw Baska ", "General Manager", "OSM Poland Sp. z o.o.", "Al. Grunwaldzka 472D", "80-309 Gdansk", "Poland", "+ 48 58 661 59 61", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)
        pobjClass = New OfficeItem(6, "Riga", "Maris Grigorovics", "General Manager", "OSM Crew Management Latvia, SIA", "3-105 Republikas laukums", "Riga LV-1010, Latvia", "Riga", "+37129258996", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)
        pobjClass = New OfficeItem(7, "Russia", "Aleksei Reznikov ", "General Manager", "OSM Crew Management St.Petersburg LLC", "3A Konstitutsii Square, 6th floor, Piramida Business Centre", "196247, St. Petersburg", "Russia", "+78126435532", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)
        pobjClass = New OfficeItem(8, "Ukraine", "Igor Gezha", "General Manager", "OSM Crew Management Ukraine Ltd.", "33 Zhukovskogo Street, Office 204", "Odessa 65045, Ukraine", "Ukraine", "+38 067 5569409", 0, 0, True)
        MyBase.Add(pobjClass.Id, pobjClass)

    End Sub
End Class
