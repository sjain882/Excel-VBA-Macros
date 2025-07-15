Option Explicit
Dim rng As Range
Dim cell As Range
Dim Filename As String

' *** WARNING: No undo available after running this script! ***
' Source: https://superuser.com/a/1404121
' Modified to preserve aspect ratio and centre image in the cell.

Sub URLPictureInsert()
    Dim theShape As Shape
    Dim xRg As Range
    Dim xCol As Long
    On Error Resume Next
    
    ' Prevent flashing and UI updates from slowing down application
    Application.ScreenUpdating = False
    
    ' Set to the range of cells you want to change to pictures
    Set rng = Selection
    
    ' Iterate all cells in this range
    For Each cell In rng
        Filename = cell
        
        ' Use Shapes instead so that we can force it to save with the document
        ' Open .xlsm with 7-Zip and go to xl\media to verify
        Set theShape = ActiveSheet.Shapes.AddPicture( _
            Filename:=Filename, linktofile:=msoFalse, _
            savewithdocument:=msoCTrue, _
            Left:=cell.Left, Top:=cell.Top, Width:=-1, Height:=-1)
            
        ' If no valid image...
        If theShape Is Nothing Then GoTo isnill
        
        ' Position image
        With theShape
        
            ' Doesn't seem to work but we'll set it anyway
            .LockAspectRatio = msoTrue
            
            ' Origin of image is top-left of cell
            .Top = cell.Top + 1
            .Left = cell.Left + 1
            
            ' Set image to move with the cell (and size, though that is likely buggy)
            .Placement = xlMove
        End With
            
            
        ' If the image is bigger than the cell (based on ratios)
        If ((theShape.Width / theShape.Height) > (cell.Width / cell.Height)) Then
            ' Use the width of both as a scale
            theShape.Width = theShape.Width / (theShape.Width / cell.Width)
            theShape.Height = theShape.Height / (theShape.Width / cell.Width)
        Else
            ' Otherwise, use the height of both as a scale
            theShape.Width = theShape.Width / (theShape.Height / cell.Height)
            theShape.Height = theShape.Height / (theShape.Height / cell.Height)
        End If
        
        ' Centre the image in the cell
        theShape.Top = cell.Top + ((cell.Height - theShape.Height) / 2)
        theShape.Left = cell.Left + ((cell.Width - theShape.Width) / 2)
        
        ' Remove the link from the cell
        cell.ClearContents
        
        ' If image is null
isnill:
        Set theShape = Nothing
        Range("A2").Select

    Next
    
    ' Present final contents to user
    Application.ScreenUpdating = True

    Debug.Print "Done " & Now

End Sub

