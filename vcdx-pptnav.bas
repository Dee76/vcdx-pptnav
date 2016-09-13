Attribute VB_Name = "Module1"
'''
' Set the slide index value for your Index slide.
'   If PowerPoint says your Index slide is slide 50 of 50, set to 50.
'   If PowerPoint says your Index slide is slide 10 of 100, set to 10.
'''
Public Const IndexSlideIndex = 7

'''''''''''''''''''''''''''''''''''
''' DO NOT EDIT BELOW THIS LINE '''
'''''''''''''''''''''''''''''''''''

' Define variables
Public SlideBeforeIndex As Integer

Sub GotoIndex()
    ' Store the current slide index value.
    SlideBeforeIndex = SlideShowWindows(1).View.Slide.SlideIndex
    ' Go to Index slide.
    SlideShowWindows(1).View.GotoSlide (IndexSlideIndex)
End Sub

Sub ReturnFromIndex()
    ' Return to slide viewed before Index slide.
    If (SlideBeforeIndex <> 0) Then
        If (SlideBeforeIndex <> SlideShowWindows(1).View.Slide.SlideIndex) Then
            SlideShowWindows(1).View.GotoSlide (SlideBeforeIndex)
        End If
    End If
End Sub

