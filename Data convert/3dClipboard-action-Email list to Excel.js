// Author: Julien Blitte
// Date: 2020-12-01

var emails = Clipboard.Value.replace(/[\r|\n|\t]+/mg, "").split(";");

var contact_format = /^ *'?([^']+)'? *<([^>]+)>$/;
var email_format = /^'?([^'@]+)@(.+)\.([^.']+)'?$/;

function cap(s)
{
	if (s.length < 1)
		return s;

	return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

function trim(s)
{
  return s.replace(/^ */, "").replace(/ *$/, "");
}

var result = "";
for(i=0; i < emails.length; i++)
{
	var email = emails[i];
	var company = "";
	var names = [];

	var match = contact_format.exec(email);
	if (match != null)
	{
		names = match[1].split(/[ ,]+/);
		email = match[2];
	}

	match = email_format.exec(email);
	if (match != null)
	{
		if (names.length == 0)
		{
			names = match[1].split(/[\._]+/);
		}

		company = cap(trim(match[2]));
	}

	while (names.length < 3)
	{
		names.push("");
	}

	for(n=0; n < names.length; n++)
	{
		names[n] = cap(trim(names[n]));
	}

	result += names.join("\t") + "\t" + company + "\t" + email + "\r\n";
}
Clipboard.Value = result;
