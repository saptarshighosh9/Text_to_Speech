Dim msg, spk

msg=InputBox("txt to speech converter :gooblu :Enter Your Message")

Set spk=CreateObject("sapi.spvoice")

spk.Speak msg
