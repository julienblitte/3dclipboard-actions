' Author: Julien Blitte

reg = "HKEY_CURRENT_USER\Control Panel\International"

sListEn =  ","
sDecimalEn = "."
sThousandEn = ","
sDateEn = "/"
sTimeEn = ":"
sString = """"

Set sh = CreateObject("WScript.Shell")
sListLc =  sh.RegRead (reg & "\sList")
sDecimalLc = sh.RegRead (reg & "\sDecimal")
sThousandLc = sh.RegRead (reg & "\sThousand")
sDateLc = sh.RegRead (reg & "\sDate")
sTimeLc = sh.RegRead (reg & "\sTime")

convert = MsgBox("Convert to local CSV format (YES)" & chr(13) & "or convert to English format (NO)?", vbYesNoCancel + vbQuestion, "CSV convert")

If convert = vbYes Then
  f = Clipboard.Value
  l = len(f)
  s = False
  t = ""
  For i = 1 To l
    c = Mid(f, i, 1)
	If c = sString Then
	  s = Not s
	End If
	If Not s Then
      Select Case (c)
      Case sListEn
        c = sListLc
      Case sDecimalEn
        c = sDecimalLc
      Case sThousandEn
        c = sThousandLc
      Case sDateEn
        c = sDateLc
      Case sTimeEn
        c = sTimeLc
      End Select
	End If
    t = t & c
  Next
  Clipboard.Value = t
ElseIf convert = vbNo Then 
  f = Clipboard.Value
  l = len(f)
  s = False
  t = ""
  For i = 1 To l
    c = Mid(f, i, 1)
	If c = sString Then
	  s = Not s
	End If
	If Not s Then
      Select Case (c)
      Case sListLc
        c = sListEn
      Case sDecimalLc
        c = sDecimalEn
      Case sThousandLc
        c = sThousandEn
      Case sDateLc
        c = sDateEn
      Case sTimeLc
        c = sTimeEn
      End Select
	End If
    t = t & c
  Next
  Clipboard.Value = t
End If