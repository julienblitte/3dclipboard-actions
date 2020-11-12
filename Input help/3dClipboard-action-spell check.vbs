'Author: Austin James
'Date:   7/27/2011

' This works with MS Word 2007 and 2010.  It is untested on other versions of MS Word.
' MS Word is required for this to work.
' Concept modified from posts by MartinLiss at http://www.vbforums.com/archive/index.php/t-283064.html
' MS Word is a product of Microsoft Corporatioon.

Dim objWord
Dim objDoc
 
' Create a Word Object and hide window
Set objWord = CreateObject("Word.Application")
objWord.WindowState = 2
objWord.Visible = False
 
' Create a new document
Set objDoc = objWord.Documents.Add(,,1,True)
 
' Add current Clipboard item to document
objDoc.Content = Clipboard.Value
 
' Make the document visible and begin the spell check
objWord.Visible = True
objWord.Activate
objDoc.CheckSpelling
 
' Get the spell checked text back from MS Word.  If the text is 
' correclty spelled, Then the item is duplicated on the clipboard
' unless the Replace Original Item checkbox is checked.
Clipboard.Value = Trim(Left(objDoc.Content, Len(objDoc.Content) - 1))
 
' Clean up
objDoc.Close(False)
Set objDoc = Nothing
 
objWord.Quit(False)
Set objWord = Nothing
