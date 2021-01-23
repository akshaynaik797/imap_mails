from email.header import decode_header
a = decode_header('=?utf-8?B?Q2FzaGxlc3MgQXV0aG9yaXphdGlvbiBMZXR0ZXIgLSBNUlMuIE1JUkEgREVCICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAk=?=')[0][0].decode("utf-8")
pass