Imports System
Imports System.IO
Imports TallComponents.PDF.Layout
Imports TallComponents.PDF.Layout.Paragraphs
Imports TallComponents.PDF.Layout.Fonts

Namespace txt2pdf

    Class Program

        Private Shared Sub Main(ByVal args() As String)
            Dim document As Document = New Document
            Dim section As Section = New Section
            section.PageSize = PageSize.Letter
            document.Sections.Add(section)

            Dim lines() As String = File.ReadAllLines(pers)
            For Each line As String In lines
                Dim text As TextParagraph = New TextParagraph
                text.SpacingAfter = 12
                section.Paragraphs.Add(text)
                Dim fragment As Fragment = New Fragment
                fragment.FontSize = 12
                fragment.Font = Font.Courier
                fragment.Text = line
                text.Fragments.Add(fragment)
            Next
            Dim file1 As FileStream = New FileStream("test.pdf", FileMode.Create, FileAccess.Write)
            document.Write(file1)
        End Sub
    End Class
End Namespace
