Attribute VB_Name = "Module1"
Public Sub Progressbar(ctlContainer As Control, ctlPgb As Control, ByVal Percentage As Integer)
    
    '***************************************
    ' Usage:
    ' Step 1:  place a picturebox (picture1) on the form
    ' Step 2:  place a second picturebox (picture2) inside the first picturebox
    ' Step 3:  set the visibility of the second picturebox to False
    ' Step 4:  set the Backcolor of the second picturebox to "Highlight"
    ' Step 5:  set the Borderstyle of the second picturebox to "None"
    ' Step 5:  call the sub
    ' Example: Progressbar picture1, picture2, 10
    ' Note:    Labels or frames might work as well (I didn't try)
    '***************************************

    ' May not be needed, just in case
    ctlPgb.Move 0, 0, ctlPgb.Width, ctlContainer.Height

    If Percentage <= 0 Then ctlPgb.Visible = False And Percentage = 0
    If Percentage > 0 Then ctlPgb.Visible = True
    If Percentage > 100 Then Percentage = 100
    
    ctlPgb.Width = (ctlContainer.Width / 100) * Percentage

End Sub
