// Author: Vince
// Date: Fri Jul 22, 2011

Clipboard.Value = Clipboard.Value
  .split("\r\n")
  .sort()
  .join("\r\n");
