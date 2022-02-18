Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
Set sapi.Voice = sapi.GetVoices.Item(1)
 dim str
 if hour(time) < 12 then
 Sapi.speak "Good Morning Sir, Welcome Back To Your Computer."
 else
 if hour(time) > 12 then
 if hour(time) > 16 then
 Sapi.speak "Good evening sir, Welcome Back To Your Computer."
 else
 Sapi.speak "Good afternoon Sir, Welcome Back To Your Computer."
 end if
 end if
 end if