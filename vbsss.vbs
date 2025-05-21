Dim nome
nome = InputBox("Inserisci il tuo nome:", "Benvenuto")

If nome <> "" Then
    MsgBox "Ciao " & nome & "! Benvenuto nel mondo di VBScript!", vbInformation, "Saluto"
Else
    do
    X=MsgBox ("Non hai inserito un nome!",0+16, "Attenzione")
    loop
End If
