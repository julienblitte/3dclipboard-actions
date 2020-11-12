'Author: Austin James
'Date:   7/10/2011

'This action will take a single tracking number from UPS, FedEX, or the United States Postal Service (USPS) and
'perform a lookup of the status of the package.  The lookup will be displayed in a browser window.

'Tracking number format definitions derived from http://answers.google.com/answers/threadview/id/207899.html

'Initialize Variables
Dim wShell
Dim trackURL
Dim trackNum
Dim re
 
Dim isUPSnum
Dim isFedEXnum
Dim isUSPSnum
Dim trackNumTypeFound
 
'Remove spaces from tracking number for URL
trackNum = Replace(Clipboard.Value," ","")
 
'Set initial values for boolean check variables
isupsnum = False
isfedexnum = False
isuspsnum = False
trackNumTypeFound = False
 
'If the tracking number is numeric, determine which carrier the tracking number belongs to via length
If isnumeric(trackNum) Then
   Select Case len(trackNum)
      Case 11 '11 character numeric tracking numbers are from UPS
         isUPSnum = True
         trackNumTypeFound = True
      Case 12 '12 character numeric tracking numbers are from FedEX
         isFedEXnum = True
         trackNumTypeFound = True
      Case 15 '15 character numeric tracking numbers are from FedEX
         isFedEXnum = True
         trackNumTypeFound = True
      Case Else 'All other tracking numbers are probably from USPS
         isUSPSnum = True
         trackNumTypeFound = True
   End Select
Else 'If there are non numeric characters in tracking number check for UPS '1Z...' pattern
   Set re = New RegExp
   re.pattern = "^1Z"
   If re.test(trackNum) Then
      isUPSnum = True
      trackNumTypeFound = True
   End If
   Set re = Nothing
End If
 
'If we know who the carrier is, format the URL for lookup
If isUPSnum Then
   trackURL = "http://wwwapps.ups.com/WebTracking/track?track.y=10&trackNums="
ElseIf isFedEXnum Then
   trackURL = "http://www.fedex.com/Tracking/Detail?language=english&trackNum="
ElseIf isuspsnum Then
   trackURL = "http://trkcnfrm1.smi.usps.com/PTSInternetWeb/InterLabelInquiry.do?origTrackNum="
Else
End If
 
'If we know the carrier, perform the lookup
'If we do not know the carrier, alert the user.
If trackNumTypeFound Then
   Set wShell = CreateObject("Shell.Application")
   wShell.Open trackURL & trackNum
   Set wShell = Nothing
Else
   MsgBox "Cannot determine what carrier the supplied tracking number belongs to."
End If