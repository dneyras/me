Sub Simu_chocs()
Dim Debut As Date, Fin As Date

Debut = Now

If ThisWorkbook.Worksheets("PARAMETRES").Range("G9").Value = "Normal" Then
    Main.Principale
Else
    PRIIPS.MainPriips
End If

Fin = Now

ThisWorkbook.Worksheets("PARAMETRES").Range("O28") = Fin - Debut

End Sub
