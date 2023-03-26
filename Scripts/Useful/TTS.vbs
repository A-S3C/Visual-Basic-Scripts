Dim msg, sapi
msg=InputBox("Enter your text for conversion                                                                                                                                                                                          Created by FG" ,"Simpe Text to Speech" ,"Type Here")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak msg
msgbox("Your Text has been Spoken: ") + msg
msgbox("Thank You for Using")
