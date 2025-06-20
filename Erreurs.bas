Public Sub DefinitionErreurs()

With BDE
    
    .ExplicationErreur(1) = "Aucun scénario coché!"
    .ExplicationErreur(2) = "Vous avez entré un nom de produit inconnu!"
    .ExplicationErreur(3) = "Vous avez entré un model point inconnu!"
    'Et ainsi de suite pour chaque erreur
    
End With

End Sub

Public Sub TraitErreur()

Dim ComptErreur As Integer
Dim Message As String

Message = ""

For ComptErreur = 1 To NbErreurs
    Message = Message & IIf(BDE.ErreurPresente(ComptErreur), BDE.ExplicationErreur(ComptErreur) & vbCrLf, "")
Next ComptErreur

MsgBox Message, vbCritical + vbOKOnly, "Erreur"

End Sub
