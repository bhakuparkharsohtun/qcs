Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports System.Diagnostics
Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        '  ' Get the current presentation
        '  Dim pres As Presentation = Globals.ThisAddIn.Application.ActivePresentation
        '
        '  ' Loop through all the slides
        '  For Each slide As Slide In pres.Slides
        '
        '      ' Loop through all the shapes on the slide
        '      For Each shape As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
        '
        '          ' Check if the shape is hidden
        '          If shape.Visible = MsoTriState.msoFalse Then
        '
        '              ' Do something with the hidden shape
        '              ' For example, print its name to the debug window
        '              Debug.WriteLine(shape.Name)
        '
        '          End If
        '
        '      Next
        '
        '  Next

        Dim sld As PowerPoint.Slide
        Dim shp As PowerPoint.Shape
        Dim count As Integer
        Dim msg As String 'declare a variable to store the message
        count = 0
        msg = "" 'initialize the message as an empty string
        For Each sld In Globals.ThisAddIn.Application.ActivePresentation.Slides
            For Each shp In sld.Shapes
                If shp.Visible = Microsoft.Office.Core.MsoTriState.msoFalse Then
                    count = count + 1
                    msg = msg & "Slide " & sld.SlideIndex & ": " & shp.Name & " is hidden" & vbNewLine 'append the new information to the message with a new line
                End If
            Next shp
        Next sld
        MsgBox("Total number of hidden layers: " & count & vbNewLine & msg) 'display the final message with the count and the details
    End Sub

End Class
