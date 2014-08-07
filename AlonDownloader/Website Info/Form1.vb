Imports System
Imports System.IO
Imports TallComponents.PDF.Layout
Imports TallComponents.PDF.Layout.Paragraphs
Imports TallComponents.PDF.Layout.Fonts
Public Class Form1
    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer
    Dim dGreg As Date
    Dim nYearH As Integer
    Dim nMonthH As Integer
    Dim nDateH As Integer

    Dim times As Integer

    Dim W1R As Boolean
    Dim W2R As Boolean
    Dim W3R As Boolean
    Dim W4R As Boolean
    Dim ver As String = Environment.OSVersion.ToString
    Dim path As String
    Private Sub FadeIn()
        Dim iCount As Integer
        For iCount = 10 To 100 Step 2
            Me.Opacity = iCount / 100
            Me.Refresh()
            Threading.Thread.Sleep(50)
        Next
        Me.ShowInTaskbar = True
    End Sub
    Public Sub WebBrowser2_DocumentCompleted(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser2.DocumentCompleted
        Try
            Dim strongs As HtmlElementCollection = WebBrowser2.Document.GetElementsByTagName("strong")
            For Each strong As HtmlElement In strongs
                If strong.InnerText IsNot Nothing AndAlso strong.InnerText.Contains("Parshas ") Then
                    Dim Parsha As String = strong.InnerText
                    Label2.Text = Parsha & ", " & HYear
                    Label2.Visible = True
                    Label3.Text = Parsha.Replace("Parshas ", "")
                End If
            Next
            W2R = False
        Catch
            W2R = True
        End Try
    End Sub
    Public Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate("http://www.ladaat.info/showgil.aspx?par=" & DateTime.Now.AddDays(1).ToString("yyyyMMdd") & "&gil=1368")
        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width - 425
        y = Screen.PrimaryScreen.WorkingArea.Height - 425
        Me.Location = New Point(x, y)
        FadeIn() 'When done, delete.
        If ver = "6.1" Then
            path = "C:/Users/" & Environment.UserName.ToString & "/Downloads/AlonDownloader"
        ElseIf ver = "6" Then
            path = "C:/Users/" & Environment.UserName.ToString & "/Downloads/AlonDownloader"
        ElseIf ver = "5.1" Then
            path = "C:/Documents And Settings/" & Environment.UserName.ToString & "/Documents/Downloads/AlonDownloader"
        End If

    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Application.Exit()
    End Sub
    Private Sub LinkLabel2_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        drag = True
        mousex = Windows.Forms.Cursor.Position.X - Me.Left
        mousey = Windows.Forms.Cursor.Position.Y - Me.Top
    End Sub
    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        If drag Then
            Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            Me.Left = Windows.Forms.Cursor.Position.X - mousex
            Me.Cursor = Cursors.SizeAll
        End If

    End Sub
    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        drag = False
        Me.Cursor = Cursors.Arrow
    End Sub
    Public Sub WebBrowser3_DocumentCompleted(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser3.DocumentCompleted
        Try
            Dim links As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("A")
            For Each link As HtmlElement In links
                If link.InnerText IsNot Nothing AndAlso link.InnerText.Contains("Business Weekly -") Then
                    Dim BusinessURL As String = link.GetAttribute("href")
                    My.Computer.Network.DownloadFile(BusinessURL, "C:/Users/" & Environment.UserName & "Downloads/AlonDownloader/" & Label3.Text & "BW.pdf")
                End If
            Next
            W3R = False
            PictureBox3.Visible = False
        Catch
            W3R = True
            PictureBox3.Visible = True
        End Try
    End Sub
    'Public Sub WebBrowser1_DocumentCompleted(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
    'Try
    'Dim KavURL As String = WebBrowser1.Document.GetElementById("cphBottom_lnkGilayon").GetAttribute("href")
    ' If (Not System.IO.Directory.Exists("C:/Users/" & Environment.UserName & "/Downloads/AlonDownloader")) Then
    '   System.IO.Directory.CreateDirectory("C:/Users/" & Environment.UserName & "/Downloads/AlonDownloader")
    ' End If
    ' My.Computer.Network.DownloadFile(KavURL, "C:/Users/" & Environment.UserName & "/Downloads/AlonDownloader/" & Label3.Text & "KVN.pdf")
    'W1R = False
    'Catch
    'W1R = True
    'End Try
    'End Sub
    'Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
    '  If DateTime.Now.DayOfWeek = DayOfWeek.Friday Then
    '    If DateTime.Now.Hour > 10 Then
    '       FadeIn()
    '    End If
    '   End If
    ' End Sub
    'When done, uncomment.
    Public Function MonSinceFirstMolad(ByVal nYearH As Integer) _
                    As Integer
        Dim nMonSinceFirstMolad As Integer
        nYearH -= 1
        nMonSinceFirstMolad = Int(nYearH / 19) * 235
        nYearH = nYearH Mod 19
        nMonSinceFirstMolad += 12 * nYearH
        If nYearH >= 17 Then
            nMonSinceFirstMolad += 6
        ElseIf nYearH >= 14 Then
            nMonSinceFirstMolad += 5
        ElseIf nYearH >= 11 Then
            nMonSinceFirstMolad += 4
        ElseIf nYearH >= 8 Then
            nMonSinceFirstMolad += 3
        ElseIf nYearH >= 6 Then
            nMonSinceFirstMolad += 2
        ElseIf nYearH >= 3 Then
            nMonSinceFirstMolad += 1
        End If
        Return nMonSinceFirstMolad
    End Function
    Public Function IsLeapYear(ByVal nYearH As Integer) As Boolean
        Dim nYearInCycle As Integer
        nYearInCycle = nYearH Mod 19
        Return nYearInCycle = 3 Or _
               nYearInCycle = 6 Or _
               nYearInCycle = 8 Or _
               nYearInCycle = 11 Or _
               nYearInCycle = 14 Or _
               nYearInCycle = 17 Or _
               nYearInCycle = 0
    End Function
    Public Function Tishrei1(ByVal nYearH As Integer) As Date
        Dim nMonthsSinceFirstMolad As Integer
        Dim nChalakim As Integer
        Dim nHours As Integer
        Dim nDays As Integer
        Dim nDayOfWeek As Integer
        Dim dTishrei1 As Date

        nMonthsSinceFirstMolad = MonSinceFirstMolad(nYearH)
        nChalakim = 793 * nMonthsSinceFirstMolad
        nChalakim += 204
        nHours = Int(nChalakim / 1080)
        nChalakim = nChalakim Mod 1080
        nHours += nMonthsSinceFirstMolad * 12
        nHours += 5
        nDays = Int(nHours / 24)
        nHours = nHours Mod 24
        nDays += 29 * nMonthsSinceFirstMolad
        nDays += 2
        nDayOfWeek = nDays Mod 7

        If Not IsLeapYear(nYearH) And _
           nDayOfWeek = 3 And _
           (nHours * 1080) + nChalakim >= _
           (9 * 1080) + 204 Then
            nDayOfWeek = 5
            nDays += 2
        ElseIf IsLeapYear(nYearH - 1) And _
               nDayOfWeek = 2 And _
               (nHours * 1080) + nChalakim >= _
               (15 * 1080) + 589 Then
            nDayOfWeek = 3
            nDays += 1
        Else
            If nHours >= 18 Then
                nDayOfWeek += 1
                nDayOfWeek = nDayOfWeek Mod 7
                nDays += 1
            End If
            If nDayOfWeek = 1 Or _
               nDayOfWeek = 4 Or _
               nDayOfWeek = 6 Then
                nDayOfWeek += 1
                nDayOfWeek = nDayOfWeek Mod 7
                nDays += 1
            End If
        End If
        nDays -= 2067025
        dTishrei1 = #1/1/1900#
        dTishrei1 = dTishrei1.AddDays(nDays)
        Return dTishrei1
    End Function
    Public Function LengthOfYear(ByVal nYearH As Integer) As Integer
        Dim dThisTishrei1 As Date
        Dim dNextTishrei1 As Date
        Dim diff As TimeSpan
        dThisTishrei1 = Tishrei1(nYearH)
        dNextTishrei1 = Tishrei1(nYearH + 1)
        diff = dNextTishrei1.Subtract(dThisTishrei1)
        Return diff.Days
    End Function
    Public Function GregToHeb(ByVal dGreg As Date, _
                              ByRef nYearH As Integer, _
                              ByRef nMonthH As Integer, _
                              ByRef nDateH As Integer) As String
        Dim nOneMolad As Double
        Dim nAvrgYear As Double
        Dim nDays As Integer
        Dim dTishrei1 As Date
        Dim nLengthOfYear As Integer
        Dim bLeap As Boolean
        Dim bHaser As Boolean
        Dim bShalem As Boolean
        Dim nMonthLen As Integer
        Dim bWhile As Boolean

        nOneMolad = 29 + (12 / 24) + (793 / (1080 * 24))
        nAvrgYear = nOneMolad * (235 / 19)
        nDays = dGreg.Subtract(#1/1/1900#).Days
        nDays += 2067025

        nYearH = Int(CDbl(nDays) / nAvrgYear) + 1

        dTishrei1 = Tishrei1(nYearH)
        If dTishrei1 = dGreg Then
            nMonthH = 1
            nDateH = 1
        Else

            If dTishrei1 < dGreg Then

                Do While Tishrei1(nYearH + 1) <= dGreg
                    nYearH += 1
                Loop
            Else

                nYearH -= 1
                Do While Tishrei1(nYearH) > dGreg
                    nYearH -= 1
                Loop
            End If

            nDays = dGreg.Subtract(Tishrei1(nYearH)).Days

            nLengthOfYear = LengthOfYear(nYearH)
            bHaser = nLengthOfYear = 353 Or nLengthOfYear = 383
            bShalem = nLengthOfYear = 355 Or nLengthOfYear = 385
            bLeap = IsLeapYear(nYearH)

            nMonthH = 1
            Do
                Select Case nMonthH
                    Case 1, 5, 8, 10, 12 ' 30 day months
                        nMonthLen = 30
                    Case 4, 7, 9, 11, 13 ' 29 day months
                        nMonthLen = 29
                    Case 6 ' Adar A (6) will be skipped on non-leap years
                        nMonthLen = 30
                    Case 2
                        nMonthLen = IIf(bShalem, 30, 29)
                    Case 3
                        nMonthLen = IIf(bHaser, 29, 30)
                End Select
                If nDays >= nMonthLen Then
                    bWhile = True
                    If bLeap Or nMonthH <> 5 Then
                        nMonthH += 1
                    Else
                        nMonthH += 2
                    End If
                    nDays -= nMonthLen
                Else
                    bWhile = False
                End If
            Loop While bWhile
            nDateH = nDays + 1
        End If
        Return CStr(nYearH)
    End Function
    Public Sub WebBrowser4_DocumentCompleted(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser4.DocumentCompleted
        Dim TDs As HtmlElementCollection = WebBrowser4.Document.GetElementsByTagName("td")
        For Each TD As HtmlElement In TDs
            If TD.InnerText IsNot Nothing AndAlso TD.InnerText.Contains(HYear) Then
                Dim links As HtmlElementCollection = TD.GetElementsByTagName("a")
                For Each a As HtmlElement In links
                    Dim matchingLinks = (From l In links Where l.InnerText IsNot Nothing AndAlso l.Parent.InnerText.Contains(HYear)).ToArray
                    Select Case matchingLinks.Count
                        Case 0
                            W4R = True
                            PictureBox4.Visible = True
                        Case 1
                            PictureBox4.Visible = False
                            W4R = False
                            LabelPerceptions.Text = "Perceptions: " & matchingLinks(0).ToString()
                            WebBrowser4.Navigate(matchingLinks(0).GetAttribute("href") & "?print=1")
                            Do While WebBrowser4.ReadyState = WebBrowserReadyState.Complete
                                Dim pers = WebBrowser4.Document.Body.InnerText
                                Dim document As New Document()
                                Dim section As New Section()
                                section.PageSize = PageSize.Letter
                                document.Sections.Add(section)
                                Dim lines() As String = File.ReadAllLines(pers)
                                For Each line As String In lines
                                    Dim text As New TextParagraph()
                                    text.SpacingAfter = 12
                                    section.Paragraphs.Add(text)
                                    Dim fragment As New Fragment()
                                    fragment.FontSize = 12
                                    fragment.Text = line
                                    text.Fragments.Add(fragment)
                                Next line
                                Using file As New FileStream(path & Label3.Text & "P.pdf", FileMode.Create, FileAccess.Write)
                                    document.Write(file)
                                End Using
                            Loop
                        Case 2
                            Form3.Show()
                            Form3.CheckBox1.Text = matchingLinks(0).ToString()
                            Form3.CheckBox2.Text = matchingLinks(1).ToString()
                            PictureBox4.Visible = False
                            W4R = False
                        Case 3
                            Form3.Show()
                            Form3.CheckBox3.Visible = True
                            Form3.CheckBox1.Text = matchingLinks(0).ToString()
                            Form3.CheckBox2.Text = matchingLinks(1).ToString()
                            Form3.CheckBox3.Text = matchingLinks(2).ToString()
                            PictureBox4.Visible = False
                            W4R = False
                        Case Is > 3 : MsgBox("This program cannot handle this many 'Perceptions'. Sorry.")

                    End Select
                Next
            End If
        Next


    End Sub
    Public HYear As Integer = GregToHeb(DateTime.Now.Date, nYearH, nMonthH, nDateH)

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        If W1R = True Then
            WebBrowser1.Refresh()
        End If
        If W2R = True Then
            WebBrowser2.Refresh()
        End If
        If W3R = True Then
            WebBrowser3.Refresh()
        End If
        If W4R = True Then
            WebBrowser4.Refresh()
        End If
    End Sub
End Class
