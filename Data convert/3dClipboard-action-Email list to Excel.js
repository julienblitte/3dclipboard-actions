// Author: Julien Blitte
// Date: 2020-11-12

var emails = Clipboard.Value.split(";");
var contact_format = /^ *'?([^']+)'? *<([^>]+)>$/;
var email_format = /^([^@]+)@(.+)\.([^.]+)$/;

var result = "";
for(i=0; i < emails.length; i++)
{
	var email = emails[i];
	var company = "";
	var names = {};

	var match = contact_format.exec(email);
	if (match != null)
	{
		var names = match[1].split(/[ ,]+/);
		email = match[2];
	}

	match = email_format.exec(email);
	if (match != null)
	{
		if (names != "")
		{
			names = match[1].split(/[\._]+/);
		}

		company += match[2];
	}

	while (names.length < 3)
	{
		names.push("");
	}

	result += names.join("\t") + "\t" + company + "\t" + email + "\r\n";
}
Clipboard.Value = result;