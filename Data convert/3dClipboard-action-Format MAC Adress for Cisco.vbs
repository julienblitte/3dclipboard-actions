' Author: Jesse
' Date: Jul 27, 2011

Set s_reTrimWS = New RegExp
s_reTrimWS.Pattern = "[^0-9a-f]"
s_reTrimWS.Global  = True
 
Clipboard.Value = s_reTrimWS.Replace( LCase(Clipboard.Value), "" )
 
s_reTrimWS.Pattern = "(....)(....)(....)"
 
Clipboard.Value = s_reTrimWS.Replace( Clipboard.Value, "$1.$2.$3" )
