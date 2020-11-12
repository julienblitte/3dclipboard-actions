// Author: Julien Blitte
// Date: 2020-11-12

// DISCLAMER: THIS IS NOT CRYPTOGRAPHIC PASSWORD GENERATION, DO NOT USE TO GENERATE A NUCLEAR WEAPON ACTIVIATION CODE PLEASE :)

var length = 12;

var alphabets = [
"abcdefghijklmnopqrstuvwxyz",
"ABCDEFGHIJKLMNOPQRSTUVWXYZ",
"0123456789",
"!@#$%^&*()_+[]{};:<>,./?"
];

password = "";
for (var i = 0; i < length; ++i)
{
	charset = alphabets[Math.floor(alphabets.length * Math.random())];
	password += charset.charAt(Math.floor(Math.random() * charset.length));
}

Clipboard.value = password;