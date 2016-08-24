Attribute VB_Name = "Module1"
'Copy A Picture To Every Slide Of A PowerPoint
'From http://www.pptfaq.com/FAQ00780_Copy_a_picture_or_other_shape_to_every_slide_in_a_presentation.htm

Sub SuperDuper()

    Dim oSh As Shape
    Dim x As Long

    Set oSh = ActiveWindow.Selection.ShapeRange(1)
    oSh.Copy

    For x = 2 To ActivePresentation.Slides.Count
        ActivePresentation.Slides(x).Shapes.Paste
    Next

End Sub
