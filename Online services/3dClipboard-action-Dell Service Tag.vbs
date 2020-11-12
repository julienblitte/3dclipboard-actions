' Author: Julien Blitte
' Show online Dell service tag information

Set wShell = CreateObject("Shell.Application")
link = "http://www.dell.com/support/home/US/Us/USbsdt1/product-support/servicetag/" & Trim(Clipboard.Value)
wShell.Open link