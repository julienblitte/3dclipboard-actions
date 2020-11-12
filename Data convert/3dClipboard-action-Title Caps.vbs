' Action: Title Case
' Author: Bill King
' Date:   05-07-16
' Code:   VBscript
'
' This code makes sure the first letter of each word is capitalized like a title;
' Even minor words such as "a", "the", "by", etc. will be made upper case.
' Capitalized letters in the middle of a word will remain capitalized
'
' This works for Ascii characters
' lower case letters have an ascii value from 97 to 122
' Upper case letters have an ascii value from 65 to 90
'
' B. King
'
Tcase1 = ""
strCBI = clipboard.value
intlencbi = len(strcbi)
CapThis = 1                                   ' 1 - cap the current letter; 2 - leave the letter alone
                                              ' the first character should be capitalized If it is a letter
For Indx = 1 To intlencbi
   ch = Asc(Mid(strCBI, Indx, 1))             ' get the decimal value of the current letter in the string
   If ch = 32 then CapThis = 1                ' the next character should be capitalized If it is a letter
   If ch <>; 32 then 
      If CapThis = 1 then                     ' If the CapThis flag isn't set, don't do anything with the character
 	     If ((ch > 96) and (ch < 123)) Then   ' we only need to consider lower case characters
	        ch = ch - 32                      ' change ascii value to upper case
		    CapThis = 0                       ' turn off the CapThis flag
         end if
         CapThis = 0                          ' turn off the CapThis flag
      End If 
   End If
   Tcase1 = Tcase1 & Chr(ch)                  ' Keep building the new string with changed cases
Next
'
clipboard.value = TCase1
