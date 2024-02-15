Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' Add a handler for the PresentationOpen event.
        AddHandler Application.PresentationOpen, AddressOf Application_PresentationOpen
    End Sub

    Private Sub Application_PresentationOpen(ByVal Pres As PowerPoint.Presentation)
        ' Loop through all the slides in the presentation.
        For Each sld As PowerPoint.Slide In Pres.Slides
            ' Loop through all the shapes in the slide.
            For Each shp As PowerPoint.Shape In sld.Shapes
                ' Set the visible property of the shape to true.
                shp.Visible = Office.MsoTriState.msoTrue
            Next
        Next
    End Sub
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
