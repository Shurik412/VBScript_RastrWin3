Sub main
  ' Create the dialog box
  Set Dialog = CreateObject("TPFSoftware.ScriptDialog")
 ' Size the dialog box
  Dialog.SetBounds "center", "center", 231, 249
  ' Set the dialog box title
  Dialog.Title = "Three Bears Inn"
 ' Add a group of radio buttons
  Dialog.AddRadios "Porridge", "Porridge:", _
    Array("too hot", "too cold", "just right"), "too cold"
  ' Add a select list
  Dialog.AddSelectList "Chair", "Rocking Chair:", _
    Array("too fast", "too slow", "just right"), _
    "too slow", 3
  ' Add a select menu
  Dialog.AddSelectMenu "Bed", "Bed:", _
    Array("too hard", "too soft", "just right"), "just right"
  ' Add a row of buttons
  Dialog.AddButtons "button", Array("OK", "Cancel")
 ' Display the dialog box
  Set Result = Dialog.Execute
 ' Find out which button was clicked
  If Result.ValueOf("button") = "OK" Then
    PorridgeQuality = Result.ValueOf( "Porridge" )
    ChairQuality = Result.ValueOf( "Chair" )
    BedQuality = Result.ValueOf( "Bed" ) 


    MsgBox "The porridge was " & PorridgeQuality & _
           " and the rocking chair was " & ChairQuality & _
           " and the bed was " & BedQuality
  Else
    MsgBox "You did not press OK"
  End If
End Sub

Call main()