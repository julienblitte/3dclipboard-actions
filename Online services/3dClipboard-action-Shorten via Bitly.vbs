'Austin James
'7/29/2011

 
Const BITLY_USERNAME = "<YOUR_BITLY_USERNAME>"
Const BITLY_API_KEY = "<YOUR_BITLY_API_KEY>"
'Character Set used was found at http://www.blooberry.com/indexdot/html/topics/urlencoding.htm
Const VALID_CHAR_SET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890$-_.+!*'(),"
 
Dim i
Dim c
Dim bitlyURL
Dim encodedURL
Dim inputURL
Dim objHttp
 
inputURL = clipboard.value
 
encodedURL = ""
 
'Encode the URL.  This may have logical errors in it.  This is a rather quick and dirty encode routine.
If Len(inputURL) > 0 Then
' Loop through each character in the URL and encode it
i = 1
Do While i <= Len(inputURL)
        c = Mid(inputURL, i, 1)
        If c = "%" Then
       'If we find an encoded character, keep it
       encodedURL = encodedURL & c & mid(inputURL, i+1,1) & mid(inputURL, i+2,1)
       i = i + 2
        ElseIf InStr(VALID_CHAR_SET, c) = 0 Then 'If current character is not in the valid character set 
       ' convert current character to HEX
                  c = Hex(Asc(c))
                
                  'Prepend % and leading zeros (if necessary) to hex code
                  If Len(c) = 1 Then
                     c = "%0" & c
                  Else
                     c = "%" & c
                  End If
   End If
   encodedURL = encodedURL & c
        
   i = i + 1
Loop
 
End If
 
'Format URL to send to Bitly
bitlyURL = "http://api.bitly.com/v3/shorten?login=" & BITLY_USERNAME & "&apiKey=" & BITLY_API_KEY & "&longUrl=" & encodedURL & "&format=txt"
 
Set objHttp = CreateObject("Microsoft.XmlHttp")
'Send request to bitly and place short url on clipboard
objHttp.open "GET", bitlyURL, False
objHttp.send ""
clipboard.value =  objHttp.responseText