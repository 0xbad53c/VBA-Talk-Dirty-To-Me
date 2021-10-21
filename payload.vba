Sub run()
    Dim objVOICE As Object

    'create SpVoice interface COM object instance
    Set objVOICE = CreateObject("SAPI.SpVoice.1")
    
    Dim i As Integer
    i = 1
    
    'Scare the user into thinking this is ransomware
    objVOICE.Speak "Phishing alert! Hacker Detected. Pay 1 bitcoin to unlock your computer. Encrypting your files in 10 9 8 7 6 5 4 3 2 1."
    
    'repeat some stuff 10 times
    Do While i < 10
        objVOICE.Speak "Phishing alert! Hacker Detected."
        i = i + 1
    Loop

    'Inform the user about the test, which is the responsible thing to do
    objVOICE.Speak "This was a test, you failed."
End Sub

