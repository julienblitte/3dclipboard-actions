'Author: Jesse 
'Date: Jul 26, 2012

Set s_reTrimWS = New RegExp
s_reTrimWS.Pattern = "[^0-9]"
s_reTrimWS.Global  = True
 
Clipboard.Value = s_reTrimWS.Replace( Clipboard.Value, "" )

If(Len(Clipboard.Value) = 7) Then Clipboard.Value = "541" & Clipboard.Value

If(Len(Clipboard.Value) = 10) Then
  s_reTrimWS.Pattern = "(...)(...)(....)" 
  Clipboard.Value = s_reTrimWS.Replace( Clipboard.Value, "$1-$2-$3" )
End If

