// Author: Julien Blitte
// Date: 2020-11-12

var url = Clipboard.Value.split("?");

if (url.length == 2)
{
	var param_list = url[1].split("&");
	for(i = 0; i < param_list.length; i++)
	{
		var param = param_list[i].split("=");
		if (param[0] == "url")
		{
			Clipboard.Value = decodeURIComponent(param[1]);
			break;
		}
	}
}