' Author: Julien Blitte
' Show online DHL Tracking page
' Date: 2020-12-01

Set wShell = CreateObject("Shell.Application")
link = "https://www.dhl.fr/en/express/tracking.html?brand=DHL&AWB=" & Trim(Clipboard.Value)
wShell.Open link
