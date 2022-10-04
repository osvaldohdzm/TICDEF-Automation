Attribute VB_Name = "NewMacros"
Sub Replacer()
'
' Replacer Macro
'
'
Application.ScreenUpdating = False
Dim searchedWord As String
Dim newWord As String

searchedWord = "Los eventos"
newWord = " eventos"
With ActiveDocument.Range
  .Find.Text = searchedWord
  Do While .Find.Execute
    .Text = newWord
    .Collapse wdCollapseEnd
  Loop
End With

searchedWord = "  "
newWord = " "
With ActiveDocument.Range
  .Find.Text = searchedWord
  Do While .Find.Execute
    .Text = newWord
    .Collapse wdCollapseEnd
  Loop
End With


searchedWord = "Common web attack"
newWord = "ataque web común"
With ActiveDocument.Range
  .Find.Text = searchedWord
  Do While .Find.Execute
    .Text = newWord
    .Collapse wdCollapseEnd
  Loop
End With

searchedWord = "Blacklisted user agent"
newWord = "Agente de usuario en lista negra"
With ActiveDocument.Range
  .Find.Text = searchedWord
  Do While .Find.Execute
    .Text = newWord
    .Collapse wdCollapseEnd
  Loop
End With

searchedWord = "known malicious user agent"
newWord = "agente de usuario reconocido como malicioso"
With ActiveDocument.Range
  .Find.Text = searchedWord
  Do While .Find.Execute
    .Text = newWord
    .Collapse wdCollapseEnd
  Loop
End With

searchedWord = "Shellshock attack detected"
newWord = "ataque Shellshock"
With ActiveDocument.Range
  .Find.Text = searchedWord
  Do While .Find.Execute
    .Text = newWord
    .Collapse wdCollapseEnd
  Loop
End With

searchedWord = "A web attack returned code 200 (success)"
newWord = "ataque web común con código 200 (acceso exitoso a recurso)"
With ActiveDocument.Range
  .Find.Text = searchedWord
  Do While .Find.Execute
    .Text = newWord
    .Collapse wdCollapseEnd
  Loop
End With


Application.ScreenUpdating = True


End Sub
