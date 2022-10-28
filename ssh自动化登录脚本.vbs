Dim WshShell 
Set WshShell=WScript.CreateObject("WScript.Shell") 
WshShell.Run "cmd.exe"
WScript.Sleep 1500 
WshShell.SendKeys "ssh -p ¶Ë¿ÚºÅ ÕËºÅ@µØÖ·"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1500 
WshShell.SendKeys "ÃÜÂë"
WshShell.SendKeys "{ENTER}"                                                                                                                                                                      