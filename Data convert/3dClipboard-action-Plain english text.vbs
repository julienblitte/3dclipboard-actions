' Action: plain text
' Author: Bill King
' Date: 04-11-14
' Code: VBscript
'
' Some notes:
'    This only works on Ascii characters
'    The resulting string does NOT include "extended ascii" characters (e.g. non-english or mathematical characters)
'    The resulting string DOES include the 'linefeed', 'carriage return', and 'tab' characters.
'

Tcase1 = ""
strCBI = clipboard.value
intlencbi = len(strcbi)
For Indx = 1 To intlencbi
   ch = Asc(Mid(strCBI, Indx, 1))

   If ((ch > 31) and (ch < 127)) Then			' letter is printable
      Tcase1 = Tcase1 & Chr(ch)					' Keep building the new string
   ElseIf (ch=10) Then							' letter is "line feed"
		Tcase1 = Tcase1 & Chr(ch)					' Keep building the new string
'		End If
   ElseIf (ch=13) Then							' letter is "carriage return"
		Tcase1 = Tcase1 & Chr(ch)					' Keep building the new string
'		End If
   ElseIf (ch=9) Then							' letter is "tab"
		Tcase1 = Tcase1 & Chr(ch)					' Keep building the new string
'		End If
   End If
   
' --- if the character is not one of the above, then it is skipped
Next
'
clipboard.value = TCase1