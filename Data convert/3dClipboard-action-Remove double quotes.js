// Author: SpiroC
// Date: Thu May 22, 2014

var strClip = Clipboard.Value;
if (strClip.substring(0, 1) == '"')
	strClip = strClip.substring(1, strClip.length-1);
if (strClip.substring(strClip.length-1, 1) == '"')
	strClip = strClip.substring(0, strClip.length-1);
Clipboard.Value = strClip;
