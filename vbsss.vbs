Dim nome
nome = InputBox("Inserisci il tuo nome:", "Benvenuto")

If nome <> "" Then
    MsgBox "Ciao " & nome & "! Benvenuto nel mondo di VBScript!", vbInformation, "Saluto"
Else
    MsgBox "Non hai inserito un nome!", vbExclamation, "Attenzione"
End If
