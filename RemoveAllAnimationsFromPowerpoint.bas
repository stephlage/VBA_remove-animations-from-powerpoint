Attribute VB_Name = "Module1"
Sub RemoveAllAnimations()
'created by stephanie frantz

Dim sld As Slide
Dim x As Long
Dim animation As Long

'Loop thru slides
  For Each sld In ActivePresentation.Slides
    
    'Loop all animations
      For x = sld.TimeLine.MainSequence.Count To 1 Step -1
        
        'remove the animation
          sld.TimeLine.MainSequence.Item(x).Delete
        
        'Maintain Deletion Stat
          animation = animation + 1
          
      Next x
  
  Next sld

'Completion Notification
MsgBox animation & " Animation(s) were removed from you PowerPoint presentation!"

End Sub
