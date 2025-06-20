'*******************************************************************************
Sub EcritRésultatsCatRachat()
'*******************************************************************************
Dim FeuilResultat As Worksheet

Set FeuilResultat = ThisWorkbook.Worksheets("MODEL POINT")

For NumLgn = 1 To NbContrats
    With Donnees(NumLgn)
        FeuilResultat.Cells(NumLgn + 1, 45) = .CatRachatTot
    End With
Next NumLgn

End Sub

'Afficher le txRachatTot pour verifier qu'on sélectionne bien le bon tx en fct des 5 catégories
'*******************************************************************************
Sub EcritRésultatsTxRachat()
'*******************************************************************************
Dim FeuilResultat As Worksheet
'afficher le TxRachat avec l'ancienneté 0 pour tous les contrats
Set FeuilResultat = ThisWorkbook.Worksheets("MODEL POINT")

For NumLgn = 1 To NbContrats
        For CompteurCat = 1 To NbCat
            With Donnees(NumLgn)
            ' si la valeur de la categorie (ds les Hyp) est égale à la valeur de la categorie dans le MP alors on lui assigne le Tx de Rachat correpondant à sa catégorie
                If (ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurCat + 802).Value) = .CatRachatTot Then
                    If .CatRachatTot = 1 Then FeuilResultat.Cells(NumLgn + 1, 48) = TxRachatTot(1, 0) 'CompteurCat=1
                    ElseIf .CatRachatTot = 2 Then FeuilResultat.Cells(NumLgn + 1, 48) = TxRachatTot(2, 0) 'CompteurCat=2
                    ElseIf .CatRachatTot = 3 Then FeuilResultat.Cells(NumLgn + 1, 48) = TxRachatTot(3, 0)
                    ElseIf .CatRachatTot = 4 Then FeuilResultat.Cells(NumLgn + 1, 48) = TxRachatTot(4, 0)
                    ElseIf .CatRachatTot = 5 Then FeuilResultat.Cells(NumLgn + 1, 48) = TxRachatTot(5, 0)
                    End If
            End With
        Next CompteurCat
Next NumLgn

End Sub


'*******************************************************************************
Sub EcritRésultatsTxRachat2()
'*******************************************************************************
Dim FeuilResultat As Worksheet
'afficher le TxRachat avec l'ancienneté 100 pour tous les contrats
Set FeuilResultat = ThisWorkbook.Worksheets("MODEL POINT")

For NumLgn = 1 To NbContrats
        For CompteurCat = 1 To NbCat
            With Donnees(NumLgn)
            ' si la valeur de la categorie (ds les Hyp) est égale à la valeur de la categorie dans le MP alors on lui assigne le Tx de Rachat correpondant à sa catégorie
                If (ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurCat + 802).Value) = .CatRachatTot Then
                    If .CatRachatTot = 1 Then FeuilResultat.Cells(NumLgn + 1, 49) = TxRachatTot(1, 100) 'CompteurCat=1
                    ElseIf .CatRachatTot = 2 Then FeuilResultat.Cells(NumLgn + 1, 49) = TxRachatTot(2, 100) 'CompteurCat=2
                    ElseIf .CatRachatTot = 3 Then FeuilResultat.Cells(NumLgn + 1, 49) = TxRachatTot(3, 100)
                    ElseIf .CatRachatTot = 4 Then FeuilResultat.Cells(NumLgn + 1, 49) = TxRachatTot(4, 100)
                    ElseIf .CatRachatTot = 5 Then FeuilResultat.Cells(NumLgn + 1, 49) = TxRachatTot(5, 100)
                    End If
            End With
        Next CompteurCat
Next NumLgn

End Sub


'*******************************************************************************
Sub EcritRésultats()
'*******************************************************************************
Dim NumLgn As Integer, FeuilResultat As Worksheet, CompteurAnnee As Integer

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP EURO")

For CompteurAnnee = 1 To Horizon
    With Totaux_Pour_Affichage(CompteurAnnee)
        
        FeuilResultat.Cells(49 * NumChoc + 6, 5 + CompteurAnnee) = .Somm_NbOuverturesEuro
        FeuilResultat.Cells(49 * NumChoc + 7, 5 + CompteurAnnee) = .Somm_NbDecesEuro
        FeuilResultat.Cells(49 * NumChoc + 8, 5 + CompteurAnnee) = .Somm_NbTiragesEuro
        FeuilResultat.Cells(49 * NumChoc + 9, 5 + CompteurAnnee) = .Somm_NbRachatsPartEuro
        FeuilResultat.Cells(49 * NumChoc + 10, 5 + CompteurAnnee) = .Somm_NbRachatsTotEuro
        FeuilResultat.Cells(49 * NumChoc + 11, 5 + CompteurAnnee) = .Somm_NbTermesEuro
        
        If CompteurAnnee = 1 And Perimetre = "NB" Then
            FeuilResultat.Cells(49 * NumChoc + 14, 5) = Totaux_Pour_Affichage(0).Somm_PrimeCommEuro
            FeuilResultat.Cells(49 * NumChoc + 15, 5) = Totaux_Pour_Affichage(0).Somm_ChargPrimeEuro
            FeuilResultat.Cells(49 * NumChoc + 16, 5) = Totaux_Pour_Affichage(0).Somm_Commissions_PrimesEuro
        ElseIf CompteurAnnee = 1 Then
            FeuilResultat.Cells(49 * NumChoc + 14, 5) = ""
            FeuilResultat.Cells(49 * NumChoc + 15, 5) = ""
            FeuilResultat.Cells(49 * NumChoc + 16, 5) = ""
        End If
        FeuilResultat.Cells(49 * NumChoc + 14, 5 + CompteurAnnee) = .Somm_PrimeCommEuro
        FeuilResultat.Cells(49 * NumChoc + 15, 5 + CompteurAnnee) = .Somm_ChargPrimeEuro
        FeuilResultat.Cells(49 * NumChoc + 16, 5 + CompteurAnnee) = .Somm_Commissions_PrimesEuro
        
        FeuilResultat.Cells(49 * NumChoc + 18, 5 + CompteurAnnee) = .Somm_PM_OuvertureEuro
        FeuilResultat.Cells(49 * NumChoc + 19, 5 + CompteurAnnee) = .Somm_InteretsPM_MiPeriodeEuro + .Somm_InteretsFinPeriodeEuro
        FeuilResultat.Cells(49 * NumChoc + 20, 5 + CompteurAnnee) = .Somm_BonusRetEuro
        
        FeuilResultat.Cells(49 * NumChoc + 22, 5 + CompteurAnnee) = .Somm_SinDecesEuro
        FeuilResultat.Cells(49 * NumChoc + 23, 5 + CompteurAnnee) = .Somm_SinTirageEuro
        FeuilResultat.Cells(49 * NumChoc + 24, 5 + CompteurAnnee) = .Somm_SinRachatPartEuro
        FeuilResultat.Cells(49 * NumChoc + 25, 5 + CompteurAnnee) = .Somm_SinRachatTotEuro
        FeuilResultat.Cells(49 * NumChoc + 26, 5 + CompteurAnnee) = .Somm_SinTermeEuro
        
        FeuilResultat.Cells(49 * NumChoc + 28, 5 + CompteurAnnee) = .Somm_ChargDecesEuro
        FeuilResultat.Cells(49 * NumChoc + 29, 5 + CompteurAnnee) = .Somm_ChargTirageEuro
        FeuilResultat.Cells(49 * NumChoc + 30, 5 + CompteurAnnee) = .Somm_ChargRachatPartEuro
        FeuilResultat.Cells(49 * NumChoc + 31, 5 + CompteurAnnee) = .Somm_ChargRachatTotEuro
        FeuilResultat.Cells(49 * NumChoc + 32, 5 + CompteurAnnee) = .Somm_ChargTermeEuro
        FeuilResultat.Cells(49 * NumChoc + 33, 5 + CompteurAnnee) = .Somm_ChargPM_Euro
        FeuilResultat.Cells(49 * NumChoc + 34, 5 + CompteurAnnee) = .Somm_Commissions_PMeuro
        
        
        FeuilResultat.Cells(49 * NumChoc + 37, 5 + CompteurAnnee) = .Somm_PM_ClotureEuro
    End With
Next CompteurAnnee

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP UC")

For CompteurAnnee = 1 To Horizon
    With Totaux_Pour_Affichage(CompteurAnnee)
        FeuilResultat.Cells(49 * NumChoc + 6, 5 + CompteurAnnee) = .Somm_NbOuverturesUC
        FeuilResultat.Cells(49 * NumChoc + 7, 5 + CompteurAnnee) = .Somm_NbDecesUC
        FeuilResultat.Cells(49 * NumChoc + 8, 5 + CompteurAnnee) = .Somm_NbTiragesUC
        FeuilResultat.Cells(49 * NumChoc + 9, 5 + CompteurAnnee) = .Somm_NbRachatsPartUC
        FeuilResultat.Cells(49 * NumChoc + 10, 5 + CompteurAnnee) = .Somm_NbRachatsTotUC
        FeuilResultat.Cells(49 * NumChoc + 11, 5 + CompteurAnnee) = .Somm_NbTermesUC
        
        If CompteurAnnee = 1 And Perimetre = "NB" Then
            FeuilResultat.Cells(49 * NumChoc + 14, 5) = Totaux_Pour_Affichage(0).Somm_PrimeCommUC
            FeuilResultat.Cells(49 * NumChoc + 15, 5) = Totaux_Pour_Affichage(0).Somm_ChargPrimeUC
            FeuilResultat.Cells(49 * NumChoc + 16, 5) = Totaux_Pour_Affichage(0).Somm_Commissions_PrimesUC
        ElseIf CompteurAnnee = 1 Then
            FeuilResultat.Cells(49 * NumChoc + 14, 5) = ""
            FeuilResultat.Cells(49 * NumChoc + 15, 5) = ""
            FeuilResultat.Cells(49 * NumChoc + 16, 5) = ""
        End If
        FeuilResultat.Cells(49 * NumChoc + 14, 5 + CompteurAnnee) = .Somm_PrimeCommUC
        FeuilResultat.Cells(49 * NumChoc + 15, 5 + CompteurAnnee) = .Somm_ChargPrimeUC
        FeuilResultat.Cells(49 * NumChoc + 16, 5 + CompteurAnnee) = .Somm_Commissions_PrimesUC
        
        FeuilResultat.Cells(49 * NumChoc + 18, 5 + CompteurAnnee) = .Somm_PM_OuvertureUC
        FeuilResultat.Cells(49 * NumChoc + 19, 5 + CompteurAnnee) = 0
        FeuilResultat.Cells(49 * NumChoc + 20, 5 + CompteurAnnee) = .Somm_BonusRetUC
        
        FeuilResultat.Cells(49 * NumChoc + 22, 5 + CompteurAnnee) = .Somm_SinDecesUC
        FeuilResultat.Cells(49 * NumChoc + 23, 5 + CompteurAnnee) = .Somm_SinTirageUC
        FeuilResultat.Cells(49 * NumChoc + 24, 5 + CompteurAnnee) = .Somm_SinRachatPartUC
        FeuilResultat.Cells(49 * NumChoc + 25, 5 + CompteurAnnee) = .Somm_SinRachatTotUC
        FeuilResultat.Cells(49 * NumChoc + 26, 5 + CompteurAnnee) = .Somm_SinTermeUC
        
        FeuilResultat.Cells(49 * NumChoc + 28, 5 + CompteurAnnee) = .Somm_ChargDecesUC
        FeuilResultat.Cells(49 * NumChoc + 29, 5 + CompteurAnnee) = .Somm_ChargTirageUC
        FeuilResultat.Cells(49 * NumChoc + 30, 5 + CompteurAnnee) = .Somm_ChargRachatPartUC
        FeuilResultat.Cells(49 * NumChoc + 31, 5 + CompteurAnnee) = .Somm_ChargRachatTotUC
        FeuilResultat.Cells(49 * NumChoc + 32, 5 + CompteurAnnee) = .Somm_ChargTermeUC
        FeuilResultat.Cells(49 * NumChoc + 33, 5 + CompteurAnnee) = .Somm_ChargPM_UC
        FeuilResultat.Cells(49 * NumChoc + 34, 5 + CompteurAnnee) = .Somm_Commissions_PMuc
        FeuilResultat.Cells(49 * NumChoc + 35, 5 + CompteurAnnee) = .Somm_RetroGlobalPM_UC
        FeuilResultat.Cells(49 * NumChoc + 36, 5 + CompteurAnnee) = .Somm_RetroAEPM_UC
        
        FeuilResultat.Cells(49 * NumChoc + 37, 5 + CompteurAnnee) = .Somm_PM_ClotureUC
    End With
Next CompteurAnnee


End Sub

'Nouveau format d'onglet de synthèse

'*******************************************************************************
Sub EcritRésultatsNew()
'*******************************************************************************
Dim NumLgn As Integer, FeuilResultat As Worksheet, CompteurAnnee As Integer

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP EURO (NEW)")

For CompteurAnnee = 1 To Horizon - 10
    With Totaux_Pour_Affichage(CompteurAnnee)
        
        FeuilResultat.Cells(63 * NumChoc + 23, 6 + CompteurAnnee) = .Somm_NbOuverturesEuro
        FeuilResultat.Cells(63 * NumChoc + 24, 6 + CompteurAnnee) = .Somm_PrimeCommEuro
        
        FeuilResultat.Cells(63 * NumChoc + 26, 6 + CompteurAnnee) = .Somm_Commissions_PrimesEuro
        FeuilResultat.Cells(63 * NumChoc + 27, 6 + CompteurAnnee) = .Somm_Commissions_PMeuro
        
        'Rachats structurels :rachats totaux
        FeuilResultat.Cells(63 * NumChoc + 31, 6 + CompteurAnnee) = .Somm_SinRachatTotEuro
        FeuilResultat.Cells(63 * NumChoc + 33, 6 + CompteurAnnee) = .Somm_SinDecesEuro
        FeuilResultat.Cells(63 * NumChoc + 34, 6 + CompteurAnnee) = (.Somm_SinTermeEuro + .Somm_SinTirageEuro)
        
        FeuilResultat.Cells(63 * NumChoc + 37, 6 + CompteurAnnee) = .Somm_FraisAdmEuro
        FeuilResultat.Cells(63 * NumChoc + 38, 6 + CompteurAnnee) = .Somm_FraisPrestEuro
        
        FeuilResultat.Cells(63 * NumChoc + 40, 6 + CompteurAnnee) = .Somm_InteretsPM_MiPeriodeEuro + .Somm_InteretsFinPeriodeEuro
        FeuilResultat.Cells(63 * NumChoc + 41, 6 + CompteurAnnee) = .Somm_PB_MiPeriodeEuro + .Somm_PBFinPeriodeEuro
        
        FeuilResultat.Cells(63 * NumChoc + 47, 6 + CompteurAnnee) = .Somm_ChargPrimeEuro
        FeuilResultat.Cells(63 * NumChoc + 48, 6 + CompteurAnnee) = .Somm_ChargPM_Euro
        'Chargements sur prestations et autres : Décès + Tirage + Rachat partiel + Rachat total + Termes
        FeuilResultat.Cells(63 * NumChoc + 49, 6 + CompteurAnnee) = .Somm_ChargDecesEuro + .Somm_ChargRachatPartEuro + .Somm_ChargRachatTotEuro
        FeuilResultat.Cells(63 * NumChoc + 50, 6 + CompteurAnnee) = .Somm_ChargTermeEuro + .Somm_ChargTirageEuro
        
        FeuilResultat.Cells(63 * NumChoc + 51, 6 + CompteurAnnee) = .Somm_PM_OuvertureEuro
        FeuilResultat.Cells(63 * NumChoc + 52, 6 + CompteurAnnee) = .Somm_PM_ClotureEuro
        'Flux ?Résultats?
        
        FeuilResultat.Cells(63 * NumChoc + 60, 6 + CompteurAnnee) = .Somm_InteretsPM_MiPeriodeEuro + .Somm_InteretsFinPeriodeEuro + .Somm_BonusRetEuro _
                                                                    + .Somm_PB_MiPeriodeEuro + .Somm_PBFinPeriodeEuro
        
        'Déroule de contrats
        FeuilResultat.Cells(63 * NumChoc + 69, 6 + CompteurAnnee) = .Somm_NbOuverturesEuro
        FeuilResultat.Cells(63 * NumChoc + 70, 6 + CompteurAnnee) = .Somm_NbDecesEuro
        FeuilResultat.Cells(63 * NumChoc + 71, 6 + CompteurAnnee) = .Somm_NbTiragesEuro
        FeuilResultat.Cells(63 * NumChoc + 72, 6 + CompteurAnnee) = .Somm_NbRachatsPartEuro
        FeuilResultat.Cells(63 * NumChoc + 73, 6 + CompteurAnnee) = .Somm_NbRachatsTotEuro
        FeuilResultat.Cells(63 * NumChoc + 74, 6 + CompteurAnnee) = .Somm_NbTermesEuro
        'Quand on aura les frais : les mettre en négatif
    End With
Next CompteurAnnee


Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP UC (NEW)")

For CompteurAnnee = 1 To Horizon - 10
    With Totaux_Pour_Affichage(CompteurAnnee)
        
        FeuilResultat.Cells(59 * NumChoc + 23, 6 + CompteurAnnee) = .Somm_NbOuverturesUC
        FeuilResultat.Cells(59 * NumChoc + 24, 6 + CompteurAnnee) = .Somm_PrimeCommUC
        
        FeuilResultat.Cells(59 * NumChoc + 26, 6 + CompteurAnnee) = .Somm_Commissions_PrimesUC
        FeuilResultat.Cells(59 * NumChoc + 27, 6 + CompteurAnnee) = .Somm_Commissions_PMuc
   
        FeuilResultat.Cells(59 * NumChoc + 30, 6 + CompteurAnnee) = .Somm_RetroAEPM_UC
        
        'pas de ligne rachat dynamique
        FeuilResultat.Cells(59 * NumChoc + 32, 6 + CompteurAnnee) = .Somm_SinRachatTotUC
        FeuilResultat.Cells(59 * NumChoc + 33, 6 + CompteurAnnee) = .Somm_SinDecesUC
        FeuilResultat.Cells(59 * NumChoc + 34, 6 + CompteurAnnee) = (.Somm_SinTermeUC + .Somm_SinTirageUC)
        
        FeuilResultat.Cells(63 * NumChoc + 37, 6 + CompteurAnnee) = .Somm_FraisAdmUC
        FeuilResultat.Cells(63 * NumChoc + 38, 6 + CompteurAnnee) = .Somm_FraisPrestUC
        
        FeuilResultat.Cells(59 * NumChoc + 41, 6 + CompteurAnnee) = .Somm_ChargPrimeUC
        FeuilResultat.Cells(59 * NumChoc + 42, 6 + CompteurAnnee) = .Somm_ChargPM_UC
        FeuilResultat.Cells(59 * NumChoc + 43, 6 + CompteurAnnee) = .Somm_ChargDecesUC + .Somm_ChargRachatPartUC + .Somm_ChargRachatTotUC
        FeuilResultat.Cells(59 * NumChoc + 44, 6 + CompteurAnnee) = .Somm_ChargTermeUC + .Somm_ChargTirageUC
        'ACAV : ajustement dans le cas d'une assurance vie  pour comptabiliser les plus ou moins values
        FeuilResultat.Cells(59 * NumChoc + 45, 6 + CompteurAnnee) = .Somm_Rend_MiPeriodeUC + .Somm_RendFinPeriodeUC
        
        FeuilResultat.Cells(59 * NumChoc + 46, 6 + CompteurAnnee) = .Somm_PM_OuvertureUC
        FeuilResultat.Cells(59 * NumChoc + 47, 6 + CompteurAnnee) = .Somm_PM_ClotureUC
        
        'Rétrocession (rétrocommission globale)
        FeuilResultat.Cells(59 * NumChoc + 58, 6 + CompteurAnnee) = .Somm_RetroGlobalPM_UC * -1
        
        'Déroule de contrats
        FeuilResultat.Cells(59 * NumChoc + 65, 6 + CompteurAnnee) = .Somm_NbOuverturesUC
        FeuilResultat.Cells(59 * NumChoc + 66, 6 + CompteurAnnee) = .Somm_NbDecesUC
        FeuilResultat.Cells(59 * NumChoc + 67, 6 + CompteurAnnee) = .Somm_NbTiragesUC
        FeuilResultat.Cells(59 * NumChoc + 68, 6 + CompteurAnnee) = .Somm_NbRachatsPartUC
        FeuilResultat.Cells(59 * NumChoc + 69, 6 + CompteurAnnee) = .Somm_NbRachatsTotUC
        FeuilResultat.Cells(59 * NumChoc + 70, 6 + CompteurAnnee) = .Somm_NbTermesUC
        
    End With
Next CompteurAnnee


End Sub
'*******************************************************************************
Sub EcritRésultatsNewCourbeTx()
'*******************************************************************************
Dim NumLgn As Integer, FeuilResultat As Worksheet, CompteurAnnee As Integer

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP EURO (NEW)")

For CompteurAnnee = 1 To Horizon
        FeuilResultat.Cells(14, 6 + CompteurAnnee) = CbTaux(CompteurAnnee)
Next CompteurAnnee


Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP UC (NEW)")

For CompteurAnnee = 1 To Horizon
        FeuilResultat.Cells(14, 6 + CompteurAnnee) = CbTaux(CompteurAnnee)
Next CompteurAnnee

End Sub


'*******************************************************************************
Sub EcritRésultatsInflation()
'*******************************************************************************
Dim NumLgn As Integer, FeuilResultat As Worksheet, CompteurAnnee As Integer

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP EURO (NEW)")

For CompteurAnnee = 1 To Horizon
        FeuilResultat.Cells(524, 6 + CompteurAnnee) = Inflation(CompteurAnnee)
Next CompteurAnnee

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP UC (NEW)")

For CompteurAnnee = 1 To Horizon
        FeuilResultat.Cells(492, 6 + CompteurAnnee) = Inflation(CompteurAnnee)
Next CompteurAnnee

End Sub

'*******************************************************************************
Sub EcritInfosExécution()
'*******************************************************************************
Dim FeuilAct As Worksheet, NumAn As Integer

Set FeuilAct = ThisWorkbook.Worksheets("PARAMETRES")

With FeuilAct
    '*** Ecriture des temps de calcul de la dernière éxecution ***
'    .Cells(17, 3) = timer2 - timer1
'    .Cells(18, 3) = timer3 - timer2
'    .Cells(19, 3) = timer4 - timer3
'    .Cells(20, 3) = timer4 - timer1
    '*** Ecriture des informations de la dernière éxecution ***
    .Cells(30, 15) = Now
    .Cells(32, 15) = ThisWorkbook.BuiltinDocumentProperties("Last Author").Value
    .Cells(34, 15) = ThisWorkbook.BuiltinDocumentProperties("Last Save Time").Value
End With

End Sub


'*******************************************************************************
Sub EcritRésultatsFlexingModeling(NomChoc As String, DossierOutPut As String, FichierOutPut As String)
'*******************************************************************************
Dim NumLgn As Long, FeuilResultat As Worksheet
Dim NumAnnee As Integer, NumMP As Integer, NumLoB As Integer, NumPB As Integer, NumTMG As Integer

Set FeuilResultat = ThisWorkbook.Worksheets("ADDACTIS " & NomChoc)
FeuilResultat.Range("A2:BA1000").ClearContents

'Si pour un choc donné, dans le pramétrage il est indiqué qu'il garder le choc, on l'écrit
'Sinon, s'il est renseigné Central, on copie colle l'onglet central, puisqu'il est écrit en 1er
If ChocAddactis(NumChoc) = True Then
    NumLgn = 1
    ' 1 : Ecriture de l'Euro
    For NumPB = 1 To NbNiveauxPB '
        For NumTMG = 1 To NbTMG
            For NumAnnee = 0 To Horizon
                With FlexingEuro(NumAnnee, NumPB, NumTMG)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 1) = NumLgn
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 2) = NumAnnee + AnneeInitiale
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 3) = .LoB
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 4) = .TMG
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 5) = .RachDyn
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 6) = .Code_support
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 7) = Round(.Effectifs, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 8) = Round(.Presta_Deces, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 9) = Round(.Presta_Rachat, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 10) = Round(.Presta_Autres, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 11) = Round(.Cotisations, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 12) = Round(.PM, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 13) = Round(.IT, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 14) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 15) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 16) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 17) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 18) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 19) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 20) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 21) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 22) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 23) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 24) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 25) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 26) = 0
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 27) = Round(.Chgt_Cotis, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 28) = Round(.Chgt_Encours, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 29) = Round(.Chgt_Presta, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 30) = Round(.Chgt_Autres, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 31) = Round(.Frais_Cotis, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 32) = Round(.Frais_Encours, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 33) = Round(.Frais_Presta, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 34) = Round(.Frais_Autres, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 35) = Round(.Indemnites_Cotis, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 36) = Round(.Indemnites_Encours, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 37) = Round(.Indemnites_Presta, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 38) = Round(.Comm_Autres, 2)
                    FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 39) = Round(.RetroAE, 2)
    
    '                FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 41) = .Presta_Terme
    '                FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 42) = .Presta_Tirage
    '                FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 43) = .Bonus
    '                FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 44) = .Chgt_Presta_deces
    '                FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 45) = .Chgt_Presta_rachat
    '                FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 46) = .Chgt_Presta_terme
    '                FeuilResultat.Cells(2 + NumAnnee + (Horizon + 1) * (NumTMG - 1) + (Horizon + 1) * NbTMG * (NumPB - 1), 47) = .Chgt_Presta_tirage
                    NumLgn = NumLgn + 1
                End With
            Next NumAnnee
        Next NumTMG
    Next NumPB
    
    ' 2 : Ecriture de l'UC
    DecalageUC = NumLgn + 1 'La longueur de l'écriture du segment Euro
    For NumAnnee = 0 To Horizon
        With FlexingUC(NumAnnee)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 1) = NumLgn
            FeuilResultat.Cells(NumAnnee + DecalageUC, 2) = NumAnnee + AnneeInitiale
            FeuilResultat.Cells(NumAnnee + DecalageUC, 3) = .LoB
            FeuilResultat.Cells(NumAnnee + DecalageUC, 4) = .TMG
            FeuilResultat.Cells(NumAnnee + DecalageUC, 5) = .RachDyn
            FeuilResultat.Cells(NumAnnee + DecalageUC, 6) = .Code_support
            FeuilResultat.Cells(NumAnnee + DecalageUC, 7) = Round(.Effectifs, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 8) = Round(.Presta_Deces, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 9) = Round(.Presta_Rachat, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 10) = Round(.Presta_Autres, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 11) = Round(.Cotisations, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 12) = Round(.PM, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 13) = Round(.IT, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 14) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 15) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 16) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 17) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 18) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 19) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 20) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 21) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 22) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 23) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 24) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 25) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 26) = 0
            FeuilResultat.Cells(NumAnnee + DecalageUC, 27) = Round(.Chgt_Cotis, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 28) = Round(.Chgt_Encours, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 29) = Round(.Chgt_Presta, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 30) = Round(.Chgt_Autres, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 31) = Round(.Frais_Cotis, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 32) = Round(.Frais_Encours, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 33) = Round(.Frais_Presta, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 34) = Round(.Frais_Autres, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 35) = Round(.Indemnites_Cotis, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 36) = Round(.Indemnites_Encours, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 37) = Round(.Indemnites_Presta, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 38) = Round(.Comm_Autres, 2)
            FeuilResultat.Cells(NumAnnee + DecalageUC, 39) = Round(.RetroAE, 2)
     
            
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 41) = .Presta_Terme
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 42) = .Presta_Tirage
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 43) = .Bonus
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 44) = .Chgt_Presta_deces
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 45) = .Chgt_Presta_rachat
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 46) = .Chgt_Presta_terme
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 47) = .Chgt_Presta_tirage
    '        FeuilResultat.Cells(NumAnnee + DecalageUC, 48) = .Retro
            
            NumLgn = NumLgn + 1
        End With
    Next NumAnnee
Else
    ThisWorkbook.Sheets("ADDACTIS CENTRAL").Activate
    Cells.Select
    Selection.Copy
    FeuilResultat.Activate
    Cells.Select
    ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats
End If

If FichierSortiesModeling = True Then
    Call ExportResultats.ExportResultatsModeling(NomChoc, DossierOutPut, FichierOutPut)
End If

End Sub

'*******************************************************************************
Sub EcritDPMContrat()
'*******************************************************************************
Dim Feuil As Worksheet

Set Feuil = ThisWorkbook.Worksheets("DPM")


For NumLgn = 1 To NbContrats
    Feuil.Cells(1 + NumLgn, 1) = BDD(NumLgn).PM_Euro(0)
    Feuil.Cells(1 + NumLgn, 2) = BDD(NumLgn).PrimeCommEuro(1)
    Feuil.Cells(1 + NumLgn, 3) = BDD(NumLgn).SinDecesEuro(1) + BDD(NumLgn).SinRachatTotEuro(1) + BDD(NumLgn).SinTermeEuro(1) + BDD(NumLgn).SinTirageEuro(1)
    Feuil.Cells(1 + NumLgn, 4) = BDD(NumLgn).InteretsPrimeEuro(1) + BDD(NumLgn).InteretsFinPeriodeEuro(1) + BDD(NumLgn).InteretsPM_MiPeriodeEuro(1) + BDD(NumLgn).BonusRetEuro(1)
    Feuil.Cells(1 + NumLgn, 5) = BDD(NumLgn).ChargPM_Euro(1) + BDD(NumLgn).ChargPrimeEuro(1) + BDD(NumLgn).ChargDecesEuro(1) + BDD(NumLgn).ChargRachatTotEuro(1) + BDD(NumLgn).ChargTermeEuro(1) + BDD(NumLgn).ChargTirageEuro(1)
    Feuil.Cells(1 + NumLgn, 6) = BDD(NumLgn).PM_Euro(1)
Next NumLgn

End Sub


