'*******************************************************************************
Sub CalculsPreliminaires_parContrat()
'*******************************************************************************

Dim NumLgn As Double
Dim AnneeProj As Integer

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        Auxdateproj(0) = DateValorisation
        .IndicObsèque = Donnees(NumLgn).IndicObsèque
        .NbClotures(0) = Donnees(NumLgn).NbTetes
        .NbClotures_Euro(0) = Donnees(NumLgn).NbTetesEuro
        .NbClotures_UC(0) = Donnees(NumLgn).NbTetesUC
        .NbClotures_Prev(0) = Donnees(NumLgn).NbTetes '* IIf(Donnees(NumLgn).DureeOptionPrev > 0, 1, 0) '** seulement là où il y a bien l'option prévoyance
'        .ModelPoint = Donnees(NumLgn).ModelPoint
        .NomProd = Donnees(NumLgn).NomProd
        .NbTetes = Donnees(NumLgn).NbTetes
        .NbTetesCU = Donnees(NumLgn).NbTetesCU
        .Sexe = Donnees(NumLgn).Sexe
        .AnneeNaissance = Donnees(NumLgn).AnneeNaissance
        .AgeAssure(0) = AnneeValorisation - .AnneeNaissance
        .AnneeEffet = Donnees(NumLgn).AnneeEffet
        .DateEffet = Donnees(NumLgn).DateEffet
        .AncienneteContrat(0) = AnneeValorisation - .AnneeEffet
        .AnneeEcheance = Donnees(NumLgn).AnneeEcheance
        .DateEcheance = Donnees(NumLgn).DateEcheance
        .ResteContrat(0) = .AnneeEcheance - AnneeValorisation
    '           .IndTermeContrat(0) = IIf(.ResteContrat(0) = 0, 1, 0)
        .IndTermeContrat(0) = IIf(DateValorisation < .DateEcheance, 0, 1)
        .TMG(0) = Donnees(NumLgn).TMG
'        .TMGprev = Donnees(NumLgn).TMGprev      '**
'        .NbPartsUC = Donnees(NumLgn).NbPartsUC
        .Taux_Com_Euro = Donnees(NumLgn).Taux_Com_Euro
        .Taux_Com_UC = Donnees(NumLgn).Taux_Com_UC
'        .Taux_Com_Prev = Donnees(NumLgn).Taux_Com_Prev '**
        .TxComSurEncours_Euro = Donnees(NumLgn).TxComSurEncours_Euro
        .TxComSurEncours_UC = Donnees(NumLgn).TxComSurEncours_UC
'        .TxComSurEncours_Prev = Donnees(NumLgn).TxComSurEncours_Prev '**
        .TxRétroGlobal = Donnees(NumLgn).TxRetroGlobal
        .TxRétroAE = Donnees(NumLgn).TxRetroAE
        .PrimeCommEuro(0) = Donnees(NumLgn).PrimeCommEuro
        .PrimeCommUC(0) = Donnees(NumLgn).PrimeCommUC
'        .PrimeCommPrev(0) = Donnees(NumLgn).PrimeCommPrev '**
        .PM_Euro(0) = Donnees(NumLgn).PM_Euro
        .PM_ClotureEuro(0) = .PM_Euro(0)
        .PM_UC(0) = Donnees(NumLgn).PM_UC
        .PM_ClotureUC(0) = .PM_UC(0)
'        .PM_Prev(0) = Donnees(NumLgn).PM_Prev '**
        .Bonus1(0) = Donnees(NumLgn).Bonus1
        .Bonus2(0) = Donnees(NumLgn).Bonus2
        .DelaiBonus1(0) = Donnees(NumLgn).DelaiBonus1
        .DelaiBonus2(0) = Donnees(NumLgn).DelaiBonus2
        .Periodicite = Donnees(NumLgn).Periodicite
        .AgeMaxprev = IIf(Donnees(NumLgn).Periodicite = "Unique", 80, 73) '**
'        .CapitalDeces(0) = Donnees(NumLgn).CapitalDeces '**
        .DureeRestante(1) = Max(Year(.DateEcheance) - Year(Auxdateproj(0)) + IIf(Month(.DateEcheance) > Month(Auxdateproj(0)), 1, 0), 0)
        .DureeRestantePrime(1) = Donnees(NumLgn).DuréeRestantePrime
'        .DureeRestante_Prev(0) = Donnees(NumLgn).DureeOptionPrev - .AncienneteContrat(0) '**
            For AnneeProj = 1 To Horizon

                .AgeAssure(AnneeProj) = 1 + .AgeAssure(AnneeProj - 1)
                .AncienneteContrat(AnneeProj) = 1 + .AncienneteContrat(AnneeProj - 1)
                .ResteContrat(AnneeProj) = .ResteContrat(AnneeProj - 1) - 1
                        '.IndTermeContrat(AnneeProj) = IIf(.ResteContrat(AnneeProj) = 0, 1, 0)
                If NumLgn = 27432 Then
                    x = x
                End If
                .IndTermeContrat(AnneeProj) = IIf(DateSerial(Year(DateValorisation) + AnneeProj - 1, Month(DateValorisation), Day(DateValorisation)) < _
                    .DateEcheance And DateSerial(Year(DateValorisation) + AnneeProj, Month(DateValorisation), Day(DateValorisation)) >= .DateEcheance, 1, 0)
                .DateAnniversaire(AnneeProj) = IIf(Month(.DateEffet) <= Month(DateValorisation), DateSerial(Year(DateValorisation) + _
                    AnneeProj, Month(.DateEffet), Day(.DateEffet)), DateSerial(Year(DateValorisation) + AnneeProj - 1, Month(.DateEffet), Day(.DateEffet)))
                Auxdateproj(AnneeProj) = DateSerial(Year(Auxdateproj(AnneeProj - 1)) + 1, Month(Auxdateproj(AnneeProj - 1)), Day(Auxdateproj(AnneeProj - 1)))
                .DureeRestante_Prev(AnneeProj) = Max(.DureeRestante_Prev(AnneeProj - 1) - 1, 0)                         '**
                .CapitalDeces(AnneeProj) = IIf(.DureeRestante_Prev(AnneeProj - 1) > 0, .CapitalDeces(AnneeProj - 1), 0) '**
            Next AnneeProj
            For AnneeProj = 2 To Horizon
                .DureeRestante(AnneeProj) = Max(.DureeRestante(AnneeProj - 1) - 1, 0)
                .DureeRestantePrime(AnneeProj) = Max(.DureeRestantePrime(AnneeProj - 1) - 1, 0)
            Next AnneeProj
        .Taux_Chargement_PM_Euro = Donnees(NumLgn).Taux_Chargement_PM_Euro
        .Taux_Chargement_PM_UC = Donnees(NumLgn).Taux_Chargement_PM_UC
        .Taux_Chargement_Prime_Euro = Donnees(NumLgn).Taux_Chargement_PrimeEuro
        .Taux_Chargement_Prime_UC = Donnees(NumLgn).Taux_Chargement_PrimeUC
        .IndicTypeTauxMin = Donnees(NumLgn).IndicTypeTauxMin
        .CatRachatTot = Donnees(NumLgn).CatRachatTot
'        '***** Permet de savoir si le contrat est lillois ou strasbourgeois (savoir où lire les taux chargement PM) *****
'
'            .TEST_Contrat_STBG = Donnees(NumLgn).TEST_Contrat_STBG
'            .TEST_Contrat_LILLE = Donnees(NumLgn).TEST_Contrat_LILLE
'            .Taux_Chargement_PM_Lille_Euro = Donnees(NumLgn).Taux_Chargement_PM_Lille_Euro
'            .Taux_Chargement_PM_Lille_UC = Donnees(NumLgn).Taux_Chargement_PM_Lille_UC
'        '***************************************************************************************************
    End With
Next NumLgn
    
End Sub

'*******************************************************************************
Sub CalculsRedimQx()
'*******************************************************************************

Dim AnneeProj, CompteurSexe, CompteurAge As Integer, CptrChoc As Integer
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    For AnneeProj = 1 To Horizon
        With BDC(BDD(NumLgn).Sexe, BDD(NumLgn).AgeAssure(AnneeProj))
            qxChoques(NumLgn, AnneeProj, 0) = .qx
        End With
    Next AnneeProj
Next NumLgn

End Sub

'*******************************************************************************
'*****************************Calcul Dx,Cx et Mx
Sub CalculsTableCommut()
'*******************************************************************************

Dim NumTMG As Long
Dim age As Integer
Dim v As Double

For NumTMG = 1 To NbTMGprev
    v = (1 / (1 + TMGprev(NumTMG)))
    
    For age = 0 To (Age_Maxi - 1)
        With TableCommut(NumTMG, age)
        
             .Dx = .lx * v ^ age
             .Cx = (.lx - TableCommut(NumTMG, age + 1).lx) * v ^ (0.5 + age)
             
        End With
    Next age
    For age = 1 To (Age_Maxi)
    
        TableCommut(NumTMG, Age_Maxi - age).Mx = TableCommut(NumTMG, Age_Maxi - age + 1).Mx + TableCommut(NumTMG, Age_Maxi - age).Cx
    
    Next age
Next NumTMG

End Sub
'*****************************

'*******************************************************************************
Sub CalculsLancerChoc()
'*******************************************************************************

If ScenCent = True Then
    LancerChoc(0) = True
End If
If ScenMort = True Then
    LancerChoc(1) = True
End If
If ScenLong = True Then
    LancerChoc(2) = True
End If
If ScenCata = True Then
    LancerChoc(3) = True
End If
If ScenHausse = True Then
    LancerChoc(4) = True
End If
If ScenBaisse = True Then
    LancerChoc(5) = True
End If
If ScenMassif = True Then
    LancerChoc(6) = True
End If

End Sub


'*******************************************************************************
Sub CalculsChocs_Qx()
'*******************************************************************************

Dim AnneeProj, CompteurProd, CompteurAnciennete As Integer
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    For AnneeProj = 1 To Horizon
        If NumChoc = 1 Then         'Choc mortalité
            If qxChoques(NumLgn, AnneeProj, 0) = 1 Then
                qxChoques(NumLgn, AnneeProj, NumChoc) = 1
            Else
                qxChoques(NumLgn, AnneeProj, NumChoc) = Min(qxChoques(NumLgn, AnneeProj, 0) * (1 + ChocMortalite), 1)
            End If
        ElseIf NumChoc = 2 Then     'Choc longévité
            If qxChoques(NumLgn, AnneeProj, 0) = 1 Then
                qxChoques(NumLgn, AnneeProj, NumChoc) = 1
            Else
                qxChoques(NumLgn, AnneeProj, NumChoc) = Min(qxChoques(NumLgn, AnneeProj, 0) * (1 + ChocLongevite), 1)
            End If
        ElseIf NumChoc = 3 Or NumChoc = 4 Or NumChoc = 5 Or NumChoc = 6 Then     'Chocs catastrophe et rachat
            qxChoques(NumLgn, AnneeProj, NumChoc) = qxChoques(NumLgn, AnneeProj, 0)
        End If
        If NumChoc = 3 Then     'Revue du scénario catastrophe pour intégration du choc
            If qxChoques(NumLgn, 1, 0) = 1 Then
                qxChoques(NumLgn, 1, NumChoc) = 1
            Else
                qxChoques(NumLgn, 1, NumChoc) = Min(qxChoques(NumLgn, 1, 0) + ChocCatastrophe, 1)
            End If
        End If
    Next AnneeProj
Next NumLgn

If ScenCent = True Or ScenMort = True Or ScenLong = True Or ScenCata = True Or ScenMassif = True Then
    For CompteurProd = 1 To NbCat
        For CompteurAnciennete = 0 To Anciennete
            With BDR(CompteurProd, CompteurAnciennete)
                .ProbaRachatTot = TxRachatTot(CompteurProd, CompteurAnciennete)
                .ProbaRachatPart = TxRachatPart(CompteurProd, CompteurAnciennete)
                .ProbaRachatTot_Prev = TxRachatTot_Prev(CompteurProd, CompteurAnciennete) '**
            End With
        Next CompteurAnciennete
    Next CompteurProd
End If
    
End Sub


'*******************************************************************************
Sub CalculsChocs_TauxRachat()
'*******************************************************************************

Dim CompteurProd, CompteurAnciennete As Integer

If NumChoc = 4 Then         'Choc rachat hausse
    For CompteurProd = 1 To NbCat
        For CompteurAnciennete = 0 To Anciennete
            With BDR(CompteurProd, CompteurAnciennete)
                .ProbaRachatTot = Min(TxRachatTot(CompteurProd, CompteurAnciennete) * (1 + ChocHausse), 1)
                .ProbaRachatPart = Min(TxRachatPart(CompteurProd, CompteurAnciennete) * (1 + ChocHausse), 1)
                .ProbaRachatTot_Prev = Min(TxRachatTot_Prev(CompteurProd, CompteurAnciennete) * (1 + ChocHausse), 1)  '**
            End With
        Next CompteurAnciennete
     Next CompteurProd
ElseIf NumChoc = 5 Then         'Choc rachat baisse
    For CompteurProd = 1 To NbCat
        For CompteurAnciennete = 0 To Anciennete
            With BDR(CompteurProd, CompteurAnciennete)
                .ProbaRachatTot = Max(TxRachatTot(CompteurProd, CompteurAnciennete) * (1 + ChocBaisse), _
                                                        TxRachatTot(CompteurProd, CompteurAnciennete) - 0.2)
                .ProbaRachatPart = Max(TxRachatPart(CompteurProd, CompteurAnciennete) * (1 + ChocBaisse), _
                                                        TxRachatPart(CompteurProd, CompteurAnciennete) - 0.2)
                .ProbaRachatTot_Prev = Max(TxRachatTot_Prev(CompteurProd, CompteurAnciennete) * (1 + ChocBaisse), _
                                                        TxRachatTot_Prev(CompteurProd, CompteurAnciennete) - 0.2)    '**
            End With
        Next CompteurAnciennete
    Next CompteurProd
ElseIf NumChoc = 0 Or NumChoc = 1 Or NumChoc = 2 Or NumChoc = 3 Or NumChoc = 6 Then            'Chocs n'impactant pas le taux de rachat
    For CompteurProd = 1 To NbCat
        For CompteurAnciennete = 0 To Anciennete
            With BDR(CompteurProd, CompteurAnciennete)
                .ProbaRachatTot = TxRachatTot(CompteurProd, CompteurAnciennete)
                .ProbaRachatPart = TxRachatPart(CompteurProd, CompteurAnciennete)
                .ProbaRachatTot_Prev = TxRachatTot_Prev(CompteurProd, CompteurAnciennete) '**
            End With
        Next CompteurAnciennete
    Next CompteurProd
End If

End Sub


'*******************************************************************************
Sub CalculsNbDeces_NbTirages_NbRachatsTot_NbRachatsPart_NbTermes_NbClotures()
'*******************************************************************************

Dim NumLgn As Double

If NumChoc = 6 And LancerChoc(NumChoc) = True And CompteurAnnee = 1 Then ' Si le choc de rachat massif est activé
 'Choc appliqué sur les rachats de l'année de proj 1
    For NumLgn = 1 To NbContrats
        With BDD(NumLgn)
            .NbDeces(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Max(0, Min(1, qxChoques(NumLgn, CompteurAnnee, NumChoc)))
            .NbTirages(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(TxTirage(.NomProd, CompteurAnnee), Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc)))
            .NbRachatsTot(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(ChocMassif, _
                                               Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee)))
            .NbRachatsPart(CompteurAnnee) = 0
            .NbTermes(CompteurAnnee) = .IndTermeContrat(CompteurAnnee) * (.NbClotures(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - _
                                        .NbRachatsTot(CompteurAnnee))   ' Estelle 11/17 Indice(n) au lieu de (n-1)
            .NbClotures(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - .NbRachatsTot(CompteurAnnee) - _
                                         .NbTermes(CompteurAnnee)
            If .NbClotures_Euro(0) > 0 Then
                .NbClotures_Euro(CompteurAnnee) = .NbClotures_Euro(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - .NbRachatsTot(CompteurAnnee) - _
                                         .NbTermes(CompteurAnnee)
            End If
            If .NbClotures_UC(0) > 0 Then
                .NbClotures_UC(CompteurAnnee) = .NbClotures_UC(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - .NbRachatsTot(CompteurAnnee) - _
                                         .NbTermes(CompteurAnnee)
            End If
            .NbDeces_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) * Max(0, Min(1, qxChoques(NumLgn, CompteurAnnee, NumChoc))) '**
            .NbRachatsTot_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot_Prev, _
                                               Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc)))                                      '**
            .NbTermes_Prev(CompteurAnnee) = IIf(.DureeRestante_Prev(CompteurAnnee - 1) = 1, .NbClotures_Prev(CompteurAnnee - 1) - .NbDeces_Prev(CompteurAnnee) - _
                                        .NbRachatsTot_Prev(CompteurAnnee), 0)                                                              '**
            .NbClotures_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) - .NbDeces_Prev(CompteurAnnee) - .NbRachatsTot_Prev(CompteurAnnee) - _
                                         .NbTermes_Prev(CompteurAnnee)                                                                     '**
        End With
    Next NumLgn
Else
    For NumLgn = 1 To NbContrats
        With BDD(NumLgn)
            .NbDeces(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Max(0, Min(1, qxChoques(NumLgn, CompteurAnnee, NumChoc)))
            .NbTirages(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(TxTirage(.NomProd, CompteurAnnee), Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc)))
            .NbRachatsTot(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot, _
                                               Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee)))
            .NbRachatsPart(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatPart, _
                                                Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee) - _
                                                BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot))
            .NbTermes(CompteurAnnee) = .IndTermeContrat(CompteurAnnee) * (.NbClotures(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - _
                                        .NbRachatsTot(CompteurAnnee))   ' Estelle 11/17 Indice(n) au lieu de (n-1)
            .NbClotures(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - .NbRachatsTot(CompteurAnnee) - _
                                         .NbTermes(CompteurAnnee)
            If .NbClotures_Euro(0) > 0 Then
                .NbClotures_Euro(CompteurAnnee) = .NbClotures_Euro(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - .NbRachatsTot(CompteurAnnee) - _
                                         .NbTermes(CompteurAnnee)
            End If
            If .NbClotures_UC(0) > 0 Then
                .NbClotures_UC(CompteurAnnee) = .NbClotures_UC(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - .NbRachatsTot(CompteurAnnee) - _
                                         .NbTermes(CompteurAnnee)
            End If
            .NbDeces_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) * Max(0, Min(1, qxChoques(NumLgn, CompteurAnnee, NumChoc))) '**
            .NbRachatsTot_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot_Prev, _
                                               Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc)))                                      '**
            .NbTermes_Prev(CompteurAnnee) = IIf(.DureeRestante_Prev(CompteurAnnee - 1) = 1, .NbClotures_Prev(CompteurAnnee - 1) - .NbDeces_Prev(CompteurAnnee) - _
                                        .NbRachatsTot_Prev(CompteurAnnee), 0)                                                              '**
            .NbClotures_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) - .NbDeces_Prev(CompteurAnnee) - .NbRachatsTot_Prev(CompteurAnnee) - _
                                         .NbTermes_Prev(CompteurAnnee)                                                                     '**
        End With
    Next NumLgn
End If

End Sub


'*******************************************************************************
Sub CalculsBonus()
'*******************************************************************************
          
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .BonusRetUC(CompteurAnnee) = 0 'Mise à 0 des bonus retenus et remplacement par leur valeur le cas échéant
        .BonusRetEuro(CompteurAnnee) = 0
        'Bonus classé euro dans le cas mixte
            
            If .DelaiBonus1(CompteurAnnee - 1) = 1 Then
                .BonusRetEuro(CompteurAnnee) = .Bonus1(CompteurAnnee - 1) * .NbClotures(CompteurAnnee - 1)
                .Bonus1(CompteurAnnee) = 0
                .Bonus2(CompteurAnnee) = .Bonus2(CompteurAnnee - 1)
            ElseIf .DelaiBonus2(CompteurAnnee - 1) = 1 Then
                .BonusRetEuro(CompteurAnnee) = .Bonus2(CompteurAnnee - 1) * .NbClotures(CompteurAnnee - 1)
                .Bonus2(CompteurAnnee) = 0
                .Bonus1(CompteurAnnee) = .Bonus1(CompteurAnnee - 1)
            Else
                .Bonus1(CompteurAnnee) = .Bonus1(CompteurAnnee - 1)
                .Bonus2(CompteurAnnee) = .Bonus2(CompteurAnnee - 1)
            End If
            
            .DelaiBonus1(CompteurAnnee) = IIf(.DelaiBonus1(CompteurAnnee - 1) <> 0, .DelaiBonus1(CompteurAnnee - 1) - 1, 0)
            .DelaiBonus2(CompteurAnnee) = IIf(.DelaiBonus2(CompteurAnnee - 1) <> 0, .DelaiBonus2(CompteurAnnee - 1) - 1, 0)
    End With
Next NumLgn

End Sub


'*******************************************************************************
Sub CalculsCoeffPrimeAnneeEch()
'*******************************************************************************

Dim NumLgn As Double
Dim AuxDatePaiement(1 To 2) As Date
Dim CompteurSemestre As Integer

For NumLgn = 1 To NbContrats

    With BDD(NumLgn)
        If .DureeRestantePrime(CompteurAnnee) = 1 Then
        
            If .Periodicite = "Unique" Then 'OK
                .CoeffPrimeAnneeEch(CompteurAnnee) = 0
                
            ElseIf .Periodicite = "Annuel" Then 'OK
                .CoeffPrimeAnneeEch(CompteurAnnee) = 1
                
            ElseIf .Periodicite = "Semestriel" Then 'OK
                If Month(.DateEffet) <= 6 Then
                    .CoeffPrimeAnneeEch(CompteurAnnee) = 1
                Else
                    .CoeffPrimeAnneeEch(CompteurAnnee) = 1 / 2
                End If
            ElseIf .Periodicite = "Trimestriel" Then 'OK
                If Month(.DateEffet) <= 3 Then
                    .CoeffPrimeAnneeEch(CompteurAnnee) = 1
                ElseIf Month(.DateEffet) <= 6 Then
                    .CoeffPrimeAnneeEch(CompteurAnnee) = 1 / 4
                ElseIf Month(.DateEffet) <= 9 Then
                    .CoeffPrimeAnneeEch(CompteurAnnee) = 1 / 2
                Else
                    .CoeffPrimeAnneeEch(CompteurAnnee) = 3 / 4
                End If
            ElseIf .Periodicite = "Mensuel" Then 'OK
                .CoeffPrimeAnneeEch(CompteurAnnee) = (DateDiff("m", Auxdateproj(CompteurAnnee - 1), .DateAnniversaire(CompteurAnnee)) - 1) / 12
            End If
              
        End If
    End With
Next NumLgn

End Sub
                      
'*******************************************************************************
Sub CalculsCoeffPrimeAnneeBonus()
'*******************************************************************************

Dim NumLgn As Double
Dim AuxDatePaiementBonus(1 To 2) As Date
Dim CompteurSemestreBonus As Integer

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If (.DelaiBonus1(CompteurAnnee - 1) = 1 And .DelaiBonus2(CompteurAnnee - 1) = 0) Or (.DelaiBonus1(CompteurAnnee - 1) = 0 And _
                                            .DelaiBonus2(CompteurAnnee - 1) = 1) Then
            If .Periodicite = "Unique" Then ' OK
                .CoeffPrimeAnneeBonus(CompteurAnnee) = 0
            ElseIf .Periodicite = "Annuel" Then ' OK
                .CoeffPrimeAnneeBonus(CompteurAnnee) = 0
            ElseIf .Periodicite = "Semestriel" Then 'OK
                If Month(.DateEffet) <= 6 Then
                    .CoeffPrimeAnneeBonus(CompteurAnnee) = 0
                Else
                    .CoeffPrimeAnneeBonus(CompteurAnnee) = 1 / 2
                End If
            ElseIf .Periodicite = "Trimestriel" Then 'OK
                If Month(.DateEffet) <= 3 Then
                    .CoeffPrimeAnneeBonus(CompteurAnnee) = 0
                ElseIf Month(.DateEffet) <= 6 Then
                    .CoeffPrimeAnneeBonus(CompteurAnnee) = 1 / 4
                ElseIf Month(.DateEffet) <= 9 Then
                    .CoeffPrimeAnneeBonus(CompteurAnnee) = 1 / 2
                Else
                    .CoeffPrimeAnneeBonus(CompteurAnnee) = 3 / 4
                End If
            ElseIf .Periodicite = "Mensuel" Then ' OK
                .CoeffPrimeAnneeBonus(CompteurAnnee) = (DateDiff("m", Auxdateproj(CompteurAnnee - 1), .DateAnniversaire(CompteurAnnee)) - 1) / 12
            End If
        End If
    End With
Next NumLgn




End Sub

'*******************************************************************************
'*****************************Option prévoyance
Sub CalculsCoeffPrimePrev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
        With BDD(NumLgn)
             .CoeffPrimePrev = IIf(.Periodicite = "Unique", 0, 1) * IIf(.AncienneteContrat(CompteurAnnee) = 1, 12 - Month(DateValorisation) + Month(.DateEffet) - 1, 0) / 12 'La périodicité est mensuelle pour l'option prévoyance
        End With
Next NumLgn


End Sub
'*****************************


'*******************************************************************************
Sub CalculsTMG()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
'        If TxTechPré2011(.NomProd, .AncienneteContrat(CompteurAnnee)) = "MP" Then
'            .TMG(CompteurAnnee) = .TMG(0)
'        Else
'            If .AnneeEffet < 2011 Then
'                .TMG(CompteurAnnee) = TxTechPré2011(.NomProd, .AncienneteContrat(CompteurAnnee))
'            ElseIf .AnneeEffet >= 2011 Then
'                .TMG(CompteurAnnee) = TxTechPost2011(.NomProd, .AncienneteContrat(CompteurAnnee))
'            End If
'        End If
        .TMG(CompteurAnnee) = .TMG(0)
    End With
Next NumLgn

End Sub
           
'*******************************************************************************
Sub CalculsPrimeEuro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If PrimeOuSans = "Sans Primes" Then ' On projette que les primes lorsqu'il y a un bonus
            If .DelaiBonus1(CompteurAnnee - 1) > 0 Or .DelaiBonus2(CompteurAnnee - 1) > 0 Then 'Tant qu'un bonus doit être versé, des primes sont projetées
                If .DureeRestantePrime(CompteurAnnee) > 1 And (.DelaiBonus1(CompteurAnnee - 1) > 1 Or .DelaiBonus2(CompteurAnnee - 1) > 1) Then
                    .PrimeCommEuro(CompteurAnnee) = .PrimeCommEuro(0) * (.NbClotures(CompteurAnnee) / .NbTetes)
                ElseIf .DureeRestantePrime(CompteurAnnee) = 1 Or (.DelaiBonus1(CompteurAnnee - 1) = 1 Or .DelaiBonus2(CompteurAnnee - 1) = 1) Then ' lorsqu'on est dans la dernière année de paiement, on doit prendre uniquement la prime jusqu'à la date anniversaire
                    .PrimeCommEuro(CompteurAnnee) = .PrimeCommEuro(0) * (.NbClotures(CompteurAnnee) / .NbTetes) _
                                            * IIf(.DureeRestante(CompteurAnnee) = 1, _
                                            .CoeffPrimeAnneeEch(CompteurAnnee), IIf(.DelaiBonus1(CompteurAnnee - 1) = 1 Or .DelaiBonus2(CompteurAnnee - 1) = 1, _
                                            .CoeffPrimeAnneeBonus(CompteurAnnee), 0))
                End If
            Else 'Si pas/plus de bonus à la clé, pas/plus de versement de primes
                .PrimeCommEuro(CompteurAnnee) = 0
            End If
        ElseIf PrimeOuSans = "Primes" Then
            If .DureeRestantePrime(CompteurAnnee) > 1 Then
                .PrimeCommEuro(CompteurAnnee) = .PrimeCommEuro(0) * (.NbClotures(CompteurAnnee) / .NbTetes)
            ElseIf .DureeRestantePrime(CompteurAnnee) = 1 Then ' lorsqu'on est dans la dernière année de paiement, on doit prendre uniquement la prime jusqu'à la date anniversaire
                .PrimeCommEuro(CompteurAnnee) = .PrimeCommEuro(0) * (.NbClotures(CompteurAnnee) / .NbTetes) * .CoeffPrimeAnneeEch(CompteurAnnee)
            Else
                .PrimeCommEuro(CompteurAnnee) = 0
            End If
        End If
    End With
Next NumLgn

End Sub


'*******************************************************************************
Sub CalculsChargPrime_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
'        .ChargPrimeEuro(CompteurAnnee) = .PrimeCommEuro(CompteurAnnee) * ChPrime(.NomProd, .AncienneteContrat(CompteurAnnee))
        .ChargPrimeEuro(CompteurAnnee) = .PrimeCommEuro(CompteurAnnee) * .Taux_Chargement_Prime_Euro
    End With
Next NumLgn

End Sub
            
'*******************************************************************************
Sub CalculsCommissionPrime_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .Commissions_PrimesEuro(CompteurAnnee) = .PrimeCommEuro(CompteurAnnee) * .Taux_Com_Euro
    End With
Next NumLgn

End Sub
            
'*******************************************************************************
Sub CalculsPrimeNette_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .PrimeNetteEuro(CompteurAnnee) = .PrimeCommEuro(CompteurAnnee) - .ChargPrimeEuro(CompteurAnnee) '- .Commissions_PrimesEuro(CompteurAnnee)
        'La commission est comprise dans le taux de chargement   Estelle 08/17
        .InteretsPrimeEuro(CompteurAnnee) = .PrimeNetteEuro(CompteurAnnee) * ((1 + .TMG(CompteurAnnee)) ^ 0.5 - 1)
        .PBPrimeEuro(CompteurAnnee) = .PrimeNetteEuro(CompteurAnnee) * ((1 + VecteurPB(CompteurAnnee)) ^ 0.5 - 1)
    End With
Next NumLgn

End Sub


'*******************************************************************************
Sub CalculsPM_MiPeriode1_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .InteretsPM_MiPeriodeEuro(CompteurAnnee) = .PM_Euro(CompteurAnnee - 1) * ((1 + .TMG(CompteurAnnee)) ^ 0.5 - 1)
        .PB_MIPeriodeEuro(CompteurAnnee) = .PM_Euro(CompteurAnnee - 1) * ((1 + VecteurPB(CompteurAnnee)) ^ 0.5 - 1)
        .PM_MiPeriode1Euro(CompteurAnnee) = .PrimeNetteEuro(CompteurAnnee) + .InteretsPrimeEuro(CompteurAnnee) + .PBPrimeEuro(CompteurAnnee) + .PM_Euro(CompteurAnnee - 1) + _
                                            .InteretsPM_MiPeriodeEuro(CompteurAnnee) + .PB_MIPeriodeEuro(CompteurAnnee) + .BonusRetEuro(CompteurAnnee) '+ .BonusRetUC(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsTauxRachat_ChocMassif_Euro()
'*******************************************************************************

Dim NumLgn As Double

If LancerChoc(NumChoc) = True Then
    If CompteurAnnee = 1 Then
        For NumLgn = 1 To NbContrats
            BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot = BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot + ChocMassif
        Next NumLgn
    End If
End If
End Sub
'*******************************************************************************
Sub CalculsSinistres_Euro()
'*******************************************************************************

Dim NumLgn As Double

If NumChoc = 6 And LancerChoc(NumChoc) = True And CompteurAnnee = 1 Then ' Si le choc de rachat massif est activé
'Choc appliqué sur les rachats de l'année de proj 1
    For NumLgn = 1 To NbContrats
        With BDD(NumLgn)
        .SinDecesEuro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) * (1 - 0.5 * (TxTirage(.NomProd, CompteurAnnee) + _
                                         ChocMassif)) * _
                                        (qxChoques(NumLgn, CompteurAnnee, NumChoc) / (1 + ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee))))
        .SinTirageEuro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) * Min(TxTirage(.NomProd, CompteurAnnee), _
                                        Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc))) / (1 + ChTirage(.NomProd, CompteurAnnee))
        .SinRachatTotEuro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) * Min(ChocMassif, _
                                            Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - _
                                            TxTirage(.NomProd, CompteurAnnee))) / (1 + ChRachatTot(.NomProd, .AncienneteContrat(CompteurAnnee)))
        .SinRachatPartEuro(CompteurAnnee) = 0
        End With
    Next NumLgn
Else
    For NumLgn = 1 To NbContrats
        With BDD(NumLgn)
            .SinDecesEuro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) * (1 - 0.5 * (TxTirage(.NomProd, CompteurAnnee) + _
                                            BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot + _
                                            BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatPart)) * _
                                            (qxChoques(NumLgn, CompteurAnnee, NumChoc) / (1 + ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee))))
            .SinTirageEuro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) * Min(TxTirage(.NomProd, CompteurAnnee), _
                                            Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc))) / (1 + ChTirage(.NomProd, CompteurAnnee))
            .SinRachatTotEuro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot, _
                                                Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - _
                                                TxTirage(.NomProd, CompteurAnnee))) / (1 + ChRachatTot(.NomProd, .AncienneteContrat(CompteurAnnee)))
            .SinRachatPartEuro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatPart, _
                                                Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee) - _
                                                BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot)) / _
                                                (1 + ChRachatPart(.NomProd, .AncienneteContrat(CompteurAnnee)))
        End With
    Next NumLgn
End If

End Sub


'*******************************************************************************
Sub CalculsChargementsSinistres_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .ChargDecesEuro(CompteurAnnee) = .SinDecesEuro(CompteurAnnee) * ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee))
        .ChargTirageEuro(CompteurAnnee) = .SinTirageEuro(CompteurAnnee) * ChTirage(.NomProd, CompteurAnnee)
        .ChargRachatTotEuro(CompteurAnnee) = .SinRachatTotEuro(CompteurAnnee) * ChRachatTot(.NomProd, .AncienneteContrat(CompteurAnnee))
        .ChargRachatPartEuro(CompteurAnnee) = .SinRachatPartEuro(CompteurAnnee) * ChRachatPart(.NomProd, .AncienneteContrat(CompteurAnnee))

        .PM_MiPeriode2Euro(CompteurAnnee) = .PM_MiPeriode1Euro(CompteurAnnee) - .SinDecesEuro(CompteurAnnee) - .SinTirageEuro(CompteurAnnee) - _
                                            .SinRachatTotEuro(CompteurAnnee) - .SinRachatPartEuro(CompteurAnnee) - .ChargDecesEuro(CompteurAnnee) - _
                                            .ChargTirageEuro(CompteurAnnee) - .ChargRachatTotEuro(CompteurAnnee) - .ChargRachatPartEuro(CompteurAnnee)
    End With
Next NumLgn

End Sub


'*******************************************************************************
Sub CalculsSinistreTerme_ChargementTerme_Euro()
'*******************************************************************************

Dim NumLgn As Double
    
For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
                If NumLgn = 29264 Then
                    x = x
                End If
        .SinTermeEuro(CompteurAnnee) = .IndTermeContrat(CompteurAnnee) * (.PM_MiPeriode2Euro(CompteurAnnee) / _
                                        (1 + ChPrestTerme(.NomProd, .AncienneteContrat(CompteurAnnee))) - (.PM_Euro(CompteurAnnee - 1) + _
                                       .InteretsPM_MiPeriodeEuro(CompteurAnnee) + .InteretsPrimeEuro(CompteurAnnee) + (.PrimeNetteEuro(CompteurAnnee) - _
                                       .SinDecesEuro(CompteurAnnee) - .SinTirageEuro(CompteurAnnee) - .SinRachatTotEuro(CompteurAnnee) - _
                                       .SinRachatPartEuro(CompteurAnnee) - .PM_MiPeriode2Euro(CompteurAnnee)) / 2) * .Taux_Chargement_PM_Euro)
        .ChargTermeEuro(CompteurAnnee) = .SinTermeEuro(CompteurAnnee) * ChPrestTerme(.NomProd, .AncienneteContrat(CompteurAnnee))
        
        .PM_MiPeriode3Euro(CompteurAnnee) = .PM_MiPeriode2Euro(CompteurAnnee) - .SinTermeEuro(CompteurAnnee) - .ChargTermeEuro(CompteurAnnee)
    End With
Next NumLgn

End Sub
        

'*******************************************************************************
Sub CalculsTransfertsCapitaux_Chargements_Euro_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .Cap_Euro_UC(CompteurAnnee) = .PM_MiPeriode3Euro(CompteurAnnee) * TxPass_Euro_UC(.NomProd, .AncienneteContrat(CompteurAnnee)) / _
                                      (1 + ChArb_Euro_UC(.NomProd, .AncienneteContrat(CompteurAnnee)))
        .ChargTransf_Euro_UC(CompteurAnnee) = .Cap_Euro_UC(CompteurAnnee) * ChArb_Euro_UC(.NomProd, .AncienneteContrat(CompteurAnnee))
        .Cap_UC_Euro(CompteurAnnee) = .PM_MiPeriode3UC(CompteurAnnee) * TxPass_UC_Euro(.NomProd, .AncienneteContrat(CompteurAnnee)) / _
                                      (1 + ChArb_UC_Euro(.NomProd, .AncienneteContrat(CompteurAnnee)))
        
        .PM_MiPeriode4Euro(CompteurAnnee) = .PM_MiPeriode3Euro(CompteurAnnee) + .Cap_UC_Euro(CompteurAnnee) - .Cap_Euro_UC(CompteurAnnee) - _
                                            .ChargTransf_Euro_UC(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsPMCloture_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .InteretsFinPeriodeEuro(CompteurAnnee) = .PM_MiPeriode4Euro(CompteurAnnee) * ((1 + .TMG(CompteurAnnee)) ^ 0.5 - 1)
        .PBFinPeriodeEuro(CompteurAnnee) = .PM_MiPeriode4Euro(CompteurAnnee) * ((1 + VecteurPB(CompteurAnnee)) ^ 0.5 - 1)
        .PM_ClotureEuro(CompteurAnnee) = IIf(.DureeRestante(CompteurAnnee) = 1, 0, .PM_MiPeriode4Euro(CompteurAnnee) + .InteretsFinPeriodeEuro(CompteurAnnee) + .PBFinPeriodeEuro(CompteurAnnee))
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CaculsChargPM_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
'        .ChargPM_Euro(CompteurAnnee) = (.TEST_Contrat_STBG * ChPM_Euro(.NomProd, .AncienneteContrat(CompteurAnnee)) + _
'                                        .TEST_Contrat_LILLE * .Taux_Chargement_PM_Lille_Euro) * (.PM_Euro(CompteurAnnee - 1) + _
'                                        .InteretsPrimeEuro(CompteurAnnee) + .InteretsFinPeriodeEuro(CompteurAnnee) + .InteretsPM_MiPeriodeEuro(CompteurAnnee) + _
'                                        0.5 * (.PrimeNetteEuro(CompteurAnnee) + .Cap_UC_Euro(CompteurAnnee) - .Cap_Euro_UC(CompteurAnnee) - _
'                                        .ChargTransf_Euro_UC(CompteurAnnee) - .SinDecesEuro(CompteurAnnee) - .SinTirageEuro(CompteurAnnee) - _
'                                        .SinRachatTotEuro(CompteurAnnee) - .SinRachatPartEuro(CompteurAnnee) - .SinTermeEuro(CompteurAnnee)))
        If .IndTermeContrat(CompteurAnnee) = 1 Then
            .ChargPM_Euro(CompteurAnnee) = (.PM_Euro(CompteurAnnee - 1) + _
                                       .InteretsPM_MiPeriodeEuro(CompteurAnnee) + .PB_MIPeriodeEuro(CompteurAnnee) + .InteretsPrimeEuro(CompteurAnnee) + .PBPrimeEuro(CompteurAnnee) _
                                       + (.PrimeNetteEuro(CompteurAnnee) - _
                                       .SinDecesEuro(CompteurAnnee) - .SinTirageEuro(CompteurAnnee) - .SinRachatTotEuro(CompteurAnnee) - _
                                       .SinRachatPartEuro(CompteurAnnee) - .PM_MiPeriode2Euro(CompteurAnnee)) / 2) * .Taux_Chargement_PM_Euro
            
        Else
            .ChargPM_Euro(CompteurAnnee) = Min(.Taux_Chargement_PM_Euro * (.PM_Euro(CompteurAnnee - 1) + _
                                            .InteretsPrimeEuro(CompteurAnnee) + .InteretsFinPeriodeEuro(CompteurAnnee) + .InteretsPM_MiPeriodeEuro(CompteurAnnee) + _
                                            .PBPrimeEuro(CompteurAnnee) + .PBFinPeriodeEuro(CompteurAnnee) + .PB_MIPeriodeEuro(CompteurAnnee) + _
                                            0.5 * (.PrimeNetteEuro(CompteurAnnee) + .Cap_UC_Euro(CompteurAnnee) - .Cap_Euro_UC(CompteurAnnee) - _
                                            .ChargTransf_Euro_UC(CompteurAnnee) - .SinDecesEuro(CompteurAnnee) - .SinTirageEuro(CompteurAnnee) - _
                                            .SinRachatTotEuro(CompteurAnnee) - .SinRachatPartEuro(CompteurAnnee) - .SinTermeEuro(CompteurAnnee))), .PM_ClotureEuro(CompteurAnnee))
        End If
    End With
Next NumLgn



End Sub
            
'*******************************************************************************
Sub CalculsCommissionPM_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .Commissions_PMeuro(CompteurAnnee) = .TxComSurEncours_Euro * (.PM_Euro(CompteurAnnee - 1) + .InteretsPrimeEuro(CompteurAnnee) + _
                                             .InteretsFinPeriodeEuro(CompteurAnnee) + .InteretsPM_MiPeriodeEuro(CompteurAnnee) + _
                                             .PBPrimeEuro(CompteurAnnee) + .PBFinPeriodeEuro(CompteurAnnee) + .PB_MIPeriodeEuro(CompteurAnnee) + _
                                             0.5 * (.PrimeNetteEuro(CompteurAnnee) + .Cap_UC_Euro(CompteurAnnee) - .Cap_Euro_UC(CompteurAnnee) - _
                                             .ChargTransf_Euro_UC(CompteurAnnee) - .SinDecesEuro(CompteurAnnee) - .SinTirageEuro(CompteurAnnee) - _
                                             .SinRachatTotEuro(CompteurAnnee) - .SinRachatPartEuro(CompteurAnnee) - .SinTermeEuro(CompteurAnnee)))
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsPM_Euro()
'*******************************************************************************

Dim NumLgn As Double

'Dim FeuilAct As Worksheet
'Set FeuilAct = Workbooks("Classeur1.xlsx").Worksheets("Feuil2")

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .PM_Euro(CompteurAnnee) = Max(0, .PM_ClotureEuro(CompteurAnnee) - .ChargPM_Euro(CompteurAnnee))
    
    End With
    
Next NumLgn

End Sub


'*******************************************************************************
Sub CalculsMarg_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .MargEuro(CompteurAnnee) = .PrimeCommEuro(CompteurAnnee) + .PM_Euro(CompteurAnnee - 1) + .InteretsPM_MiPeriodeEuro(CompteurAnnee) + _
                                   .InteretsPrimeEuro(CompteurAnnee) - .SinDecesEuro(CompteurAnnee) - .SinTirageEuro(CompteurAnnee) - _
                                   .SinRachatTotEuro(CompteurAnnee) - .SinRachatPartEuro(CompteurAnnee) - .SinTermeEuro(CompteurAnnee) + _
                                   .InteretsFinPeriodeEuro(CompteurAnnee) - .PM_Euro(CompteurAnnee) - .Cap_Euro_UC(CompteurAnnee) + .Cap_UC_Euro(CompteurAnnee) - _
                                   .Commissions_PMeuro(CompteurAnnee) - .Commissions_PrimesEuro(CompteurAnnee) _
                                   + .PBPrimeEuro(CompteurAnnee) + .PBFinPeriodeEuro(CompteurAnnee) + .PB_MIPeriodeEuro(CompteurAnnee)
                                   
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsCharg_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .ChargEuro(CompteurAnnee) = .ChargPrimeEuro(CompteurAnnee) + .ChargDecesEuro(CompteurAnnee) + .ChargTirageEuro(CompteurAnnee) + _
                                    .ChargRachatTotEuro(CompteurAnnee) + .ChargRachatPartEuro(CompteurAnnee) + .ChargTermeEuro(CompteurAnnee) + _
                                    .ChargPM_Euro(CompteurAnnee) + .ChargTransf_Euro_UC(CompteurAnnee) - _
                                   .Commissions_PMeuro(CompteurAnnee) - .Commissions_PrimesEuro(CompteurAnnee) ' Estelle 08/17
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsPrime_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If PrimeOuSans = "Sans Primes" Then ' On projette que les primes lorsqu'il y a un bonus
            If .DelaiBonus1(CompteurAnnee - 1) > 0 Or .DelaiBonus2(CompteurAnnee - 1) > 0 Then 'Tant qu'un bonus doit être versé, des primes sont projetées
                If .DureeRestantePrime(CompteurAnnee) > 1 And (.DelaiBonus1(CompteurAnnee - 1) > 1 Or .DelaiBonus2(CompteurAnnee - 1) > 1) Then
                     .PrimeCommUC(CompteurAnnee) = .PrimeCommUC(0) * (.NbClotures(CompteurAnnee) / .NbTetes)
                ElseIf .DureeRestantePrime(CompteurAnnee) = 1 Or (.DelaiBonus1(CompteurAnnee - 1) = 1 Or .DelaiBonus2(CompteurAnnee - 1) = 1) Then ' lorsqu'on est dans la dernière année de paiement, on doit prendre uniquement la prime jusqu'à la date anniversaire
                    .PrimeCommUC(CompteurAnnee) = .PrimeCommUC(0) * (.NbClotures(CompteurAnnee) / .NbTetes) _
                                          * IIf(.DureeRestante(CompteurAnnee) = 1, .CoeffPrimeAnneeEch(CompteurAnnee), IIf(.DelaiBonus1(CompteurAnnee - 1) = 1 Or .DelaiBonus2(CompteurAnnee - 1) = 1, .CoeffPrimeAnneeBonus(CompteurAnnee), 0))
                End If
            Else 'Si pas/plus de bonus à la clé, pas/plus de versement de primes
                .PrimeCommUC(CompteurAnnee) = 0
            End If
        ElseIf PrimeOuSans = "Primes" Then
            If .DureeRestantePrime(CompteurAnnee) > 1 Then
                .PrimeCommUC(CompteurAnnee) = .PrimeCommUC(0) * (.NbClotures(CompteurAnnee) / .NbTetes)
            ElseIf .DureeRestantePrime(CompteurAnnee) = 1 Then ' lorsqu'on est dans la dernière année de paiement, on doit prendre uniquement la prime jusqu'à la date anniversaire
                .PrimeCommUC(CompteurAnnee) = .PrimeCommUC(0) * (.NbClotures(CompteurAnnee) / .NbTetes) * .CoeffPrimeAnneeEch(CompteurAnnee)
            Else
                .PrimeCommUC(CompteurAnnee) = 0
            End If
        End If
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsChargPrime_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
'       .ChargPrimeUC(CompteurAnnee) = .PrimeCommUC(CompteurAnnee) * ChPrime(.NomProd, .AncienneteContrat(CompteurAnnee))
       .ChargPrimeUC(CompteurAnnee) = .PrimeCommUC(CompteurAnnee) * .Taux_Chargement_Prime_UC
    End With
Next NumLgn

End Sub
            
'*******************************************************************************
Sub CalculsCommissionPrime_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    If NumLgn = 13586 Then
        NumLgn = NumLgn
    End If
    With BDD(NumLgn)
        .Commissions_PrimesUC(CompteurAnnee) = .PrimeCommUC(CompteurAnnee) * .Taux_Com_UC
    
    End With
Next NumLgn

End Sub
            
'*******************************************************************************
Sub CalculsPrimeNette_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .PrimeNetteUC(CompteurAnnee) = .PrimeCommUC(CompteurAnnee) - .ChargPrimeUC(CompteurAnnee) '- .Commissions_PrimesUC(CompteurAnnee)    Estelle 08/17
        .InteretsPrimeUC(CompteurAnnee) = 0
                        '.PrimeNetteUC(CompteurAnnee) * ((1 + .TMG(CompteurAnnee)) ^ 0.5 - 1)
        .RendUCPrime(CompteurAnnee) = .PrimeNetteUC(CompteurAnnee) * ((1 + VecteurRendUC(CompteurAnnee)) ^ 0.5 - 1)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsPM_MiPeriode1_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats

    With BDD(NumLgn)
        If .PM_UC(0) > 0 Then
        x = x
        End If
        .InteretsPM_MiPeriodeUC(CompteurAnnee) = 0
                            '.PM_UC(CompteurAnnee - 1) * ((1 + .TMG(CompteurAnnee)) ^ 0.5 - 1)
        .RendUC_MiPeriodeUC(CompteurAnnee) = .PM_UC(CompteurAnnee - 1) * ((1 + VecteurRendUC(CompteurAnnee)) ^ 0.5 - 1)
        .PM_MiPeriode1UC(CompteurAnnee) = .PrimeNetteUC(CompteurAnnee) + .InteretsPrimeUC(CompteurAnnee) + .PM_UC(CompteurAnnee - 1) + _
                                          .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsSinistres_UC()
'*******************************************************************************

Dim NumLgn As Double

If NumChoc = 6 And LancerChoc(NumChoc) = True And CompteurAnnee = 1 Then ' Si le choc de rachat massif est activé
'Choc appliqué sur les rachats de l'année de proj 1
    For NumLgn = 1 To NbContrats
        With BDD(NumLgn)
            .SinDecesUC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) * (1 - 0.5 * (TxTirage(.NomProd, CompteurAnnee) + _
                                          ChocMassif)) * _
                                         (qxChoques(NumLgn, CompteurAnnee, NumChoc) / (1 + ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee))))
            .SinTirageUC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) * Min(TxTirage(.NomProd, CompteurAnnee), _
                                              Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc))) / (1 + ChTirage(.NomProd, CompteurAnnee))
            .SinRachatTotUC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) * Min(ChocMassif, _
                                             Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee))) / _
                                             (1 + ChRachatTot(.NomProd, .AncienneteContrat(CompteurAnnee)))
            .SinRachatPartUC(CompteurAnnee) = 0
        End With
    Next NumLgn
Else
    For NumLgn = 1 To NbContrats
        With BDD(NumLgn)
            .SinDecesUC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) * (1 - 0.5 * (TxTirage(.NomProd, CompteurAnnee) + _
                                         BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot + _
                                         BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatPart)) * _
                                         (qxChoques(NumLgn, CompteurAnnee, NumChoc) / (1 + ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee))))
            .SinTirageUC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) * Min(TxTirage(.NomProd, CompteurAnnee), _
                                              Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc))) / (1 + ChTirage(.NomProd, CompteurAnnee))
            .SinRachatTotUC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot, _
                                             Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee))) / _
                                             (1 + ChRachatTot(.NomProd, .AncienneteContrat(CompteurAnnee)))
            .SinRachatPartUC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatPart, _
                                              Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee) - _
                                              BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot)) / _
                                              (1 + ChRachatPart(.NomProd, .AncienneteContrat(CompteurAnnee)))
        End With
    Next NumLgn
End If

End Sub
                                                                                                                                     
'*******************************************************************************
Sub CalculsChargementsSinistres_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .ChargDecesUC(CompteurAnnee) = .SinDecesUC(CompteurAnnee) * ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee))
        .ChargTirageUC(CompteurAnnee) = .SinTirageUC(CompteurAnnee) * ChTirage(.NomProd, CompteurAnnee)
        .ChargRachatTotUC(CompteurAnnee) = .SinRachatTotUC(CompteurAnnee) * ChRachatTot(.NomProd, .AncienneteContrat(CompteurAnnee))
        .ChargRachatPartUC(CompteurAnnee) = .SinRachatPartUC(CompteurAnnee) * ChRachatPart(.NomProd, .AncienneteContrat(CompteurAnnee))
            
        .PM_MiPeriode2UC(CompteurAnnee) = .PM_MiPeriode1UC(CompteurAnnee) - .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - _
                                          .SinRachatTotUC(CompteurAnnee) - .SinRachatPartUC(CompteurAnnee) - .ChargDecesUC(CompteurAnnee) - _
                                          .ChargTirageUC(CompteurAnnee) - .ChargRachatTotUC(CompteurAnnee) - .ChargRachatPartUC(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsSinitreTerme_ChargementTerme_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .SinTermeUC(CompteurAnnee) = .IndTermeContrat(CompteurAnnee) * (.PM_MiPeriode2UC(CompteurAnnee) / _
                                     (1 + ChPrestTerme(.NomProd, .AncienneteContrat(CompteurAnnee))) - (.PM_UC(CompteurAnnee - 1) + _
                                     .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + .InteretsPrimeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + (.PrimeNetteUC(CompteurAnnee) - _
                                     .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - .SinRachatPartUC(CompteurAnnee) - _
                                     .PM_MiPeriode2UC(CompteurAnnee)) / 2) * (.TxRétroGlobal + _
                                     .Taux_Chargement_PM_UC))
        .ChargTermeUC(CompteurAnnee) = .SinTermeUC(CompteurAnnee) * ChPrestTerme(.NomProd, .AncienneteContrat(CompteurAnnee))
            
        .PM_MiPeriode3UC(CompteurAnnee) = .PM_MiPeriode2UC(CompteurAnnee) - .SinTermeUC(CompteurAnnee) - .ChargTermeUC(CompteurAnnee)
    End With
Next NumLgn

End Sub
       
'*******************************************************************************
Sub CalculsTransfertsCapitaux_Chargements_UC_Euro()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .ChargTransf_UC_Euro(CompteurAnnee) = .Cap_UC_Euro(CompteurAnnee) * _
                                                  ChArb_UC_Euro(.NomProd, .AncienneteContrat(CompteurAnnee))
        .ChargTransf_UC_UC(CompteurAnnee) = .PM_MiPeriode3UC(CompteurAnnee) * _
                                                TxPass_UC_UC(.NomProd, .AncienneteContrat(CompteurAnnee)) * _
                                                ChArb_UC_UC(.NomProd, .AncienneteContrat(CompteurAnnee))
        
        .PM_MiPeriode4UC(CompteurAnnee) = .PM_MiPeriode3UC(CompteurAnnee) + .Cap_Euro_UC(CompteurAnnee) - .Cap_UC_Euro(CompteurAnnee) - _
                                            .ChargTransf_UC_Euro(CompteurAnnee) - .ChargTransf_UC_UC(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsPMCloture_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .InteretsFinPeriodeUC(CompteurAnnee) = 0
        .RendFinPeriodeUC(CompteurAnnee) = .PM_MiPeriode4UC(CompteurAnnee) * ((1 + VecteurRendUC(CompteurAnnee)) ^ 0.5 - 1)
        .PM_ClotureUC(CompteurAnnee) = .PM_MiPeriode4UC(CompteurAnnee) + .InteretsFinPeriodeUC(CompteurAnnee) + .RendFinPeriodeUC(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CaculsChargPM_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
'        .ChargPM_UC(CompteurAnnee) = (.TEST_Contrat_STBG * ChPM_UC(.NomProd, .AncienneteContrat(CompteurAnnee)) + _
'                                                              .TEST_Contrat_LILLE * .Taux_Chargement_PM_Lille_UC) * _
'                                         (.PM_UC(CompteurAnnee - 1) + .InteretsPrimeUC(CompteurAnnee) + .InteretsFinPeriodeUC(CompteurAnnee) + _
'                                          .InteretsPM_MiPeriodeUC(CompteurAnnee) + 0.5 * (.PrimeNetteUC(CompteurAnnee) + .Cap_Euro_UC(CompteurAnnee) - _
'                                          .Cap_UC_Euro(CompteurAnnee) - .ChargTransf_UC_Euro(CompteurAnnee) - .ChargTransf_UC_UC(CompteurAnnee) - _
'                                          .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - _
'                                          .SinRachatPartUC(CompteurAnnee) - .SinTermeUC(CompteurAnnee) - .RetroGlobalPM_UC(CompteurAnnee)))

        If .IndTermeContrat(CompteurAnnee) = 1 Then
            .ChargPM_UC(CompteurAnnee) = (.PM_UC(CompteurAnnee - 1) + _
                                     .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + .InteretsPrimeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + (.PrimeNetteUC(CompteurAnnee) - _
                                     .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - .SinRachatPartUC(CompteurAnnee) - _
                                     .PM_MiPeriode2UC(CompteurAnnee)) / 2) * .Taux_Chargement_PM_UC
        Else
            .ChargPM_UC(CompteurAnnee) = Min(.Taux_Chargement_PM_UC * _
                                         (.PM_UC(CompteurAnnee - 1) + .InteretsPrimeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + .InteretsFinPeriodeUC(CompteurAnnee) + .RendFinPeriodeUC(CompteurAnnee) + _
                                          .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + 0.5 * (.PrimeNetteUC(CompteurAnnee) + .Cap_Euro_UC(CompteurAnnee) - _
                                          .Cap_UC_Euro(CompteurAnnee) - .ChargTransf_UC_Euro(CompteurAnnee) - .ChargTransf_UC_UC(CompteurAnnee) - _
                                          .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - _
                                          .SinRachatPartUC(CompteurAnnee) - .SinTermeUC(CompteurAnnee) - .RetroGlobalPM_UC(CompteurAnnee))), .PM_ClotureUC(CompteurAnnee) - .RetroGlobalPM_UC(CompteurAnnee))
        End If
         .PM_UC(CompteurAnnee) = Max(0, .PM_ClotureUC(CompteurAnnee) - .ChargPM_UC(CompteurAnnee) - .RetroGlobalPM_UC(CompteurAnnee)) '- .Commissions_PMuc(CompteurAnnee)   Estelle 8 / 17

    End With
Next NumLgn

End Sub
            
'*******************************************************************************
Sub CalculsCommissionPM_UC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .Commissions_PMuc(CompteurAnnee) = .TxComSurEncours_UC * (.PM_UC(CompteurAnnee - 1) + .InteretsPrimeUC(CompteurAnnee) + _
                                          .InteretsFinPeriodeUC(CompteurAnnee) + .InteretsPM_MiPeriodeUC(CompteurAnnee) + _
                                          .RendUCPrime(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + .RendFinPeriodeUC(CompteurAnnee) + _
                                          0.5 * (.PrimeNetteUC(CompteurAnnee) + .Cap_Euro_UC(CompteurAnnee) - .Cap_UC_Euro(CompteurAnnee) - _
                                          .ChargTransf_UC_Euro(CompteurAnnee) - .ChargTransf_UC_UC(CompteurAnnee) - .SinDecesUC(CompteurAnnee) - _
                                          .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - .SinRachatPartUC(CompteurAnnee) - _
                                          .SinTermeUC(CompteurAnnee)))
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsPM_UC()
'*******************************************************************************

Dim NumLgn As Double

'Dim FeuilAct As Worksheet
'Set FeuilAct = Workbooks("Classeur1.xlsx").Worksheets("Feuil2")
    
For NumLgn = 1 To NbContrats

    With BDD(NumLgn)
        If .IndTermeContrat(CompteurAnnee) = 1 Then
            .RetroGlobalPM_UC(CompteurAnnee) = (.PM_UC(CompteurAnnee - 1) + _
                                     .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + .InteretsPrimeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + (.PrimeNetteUC(CompteurAnnee) - _
                                     .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - .SinRachatPartUC(CompteurAnnee) - _
                                     .PM_MiPeriode2UC(CompteurAnnee)) / 2) * .TxRétroGlobal
            .RetroAEPM_UC(CompteurAnnee) = (.PM_UC(CompteurAnnee - 1) + _
                                     .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + .InteretsPrimeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + (.PrimeNetteUC(CompteurAnnee) - _
                                     .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - .SinRachatPartUC(CompteurAnnee) - _
                                     .PM_MiPeriode2UC(CompteurAnnee)) / 2) * .TxRétroAE
        Else
    
            .RetroGlobalPM_UC(CompteurAnnee) = Min((.TxRétroGlobal * _
                                             (.PM_UC(CompteurAnnee - 1) + .InteretsPrimeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + .InteretsFinPeriodeUC(CompteurAnnee) + .RendFinPeriodeUC(CompteurAnnee) + _
                                              .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + 0.5 * (.PrimeNetteUC(CompteurAnnee) + .Cap_Euro_UC(CompteurAnnee) - _
                                              .Cap_UC_Euro(CompteurAnnee) - .ChargTransf_UC_Euro(CompteurAnnee) - .ChargTransf_UC_UC(CompteurAnnee) - _
                                              .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - _
                                              .SinRachatPartUC(CompteurAnnee) - .SinTermeUC(CompteurAnnee)))), .PM_ClotureUC(CompteurAnnee))
            .RetroAEPM_UC(CompteurAnnee) = Min((.TxRétroAE * _
                                             (.PM_UC(CompteurAnnee - 1) + .InteretsPrimeUC(CompteurAnnee) + .RendUCPrime(CompteurAnnee) + .InteretsFinPeriodeUC(CompteurAnnee) + .RendFinPeriodeUC(CompteurAnnee) + _
                                              .InteretsPM_MiPeriodeUC(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + 0.5 * (.PrimeNetteUC(CompteurAnnee) + .Cap_Euro_UC(CompteurAnnee) - _
                                              .Cap_UC_Euro(CompteurAnnee) - .ChargTransf_UC_Euro(CompteurAnnee) - .ChargTransf_UC_UC(CompteurAnnee) - _
                                              .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - .SinRachatTotUC(CompteurAnnee) - _
                                              .SinRachatPartUC(CompteurAnnee) - .SinTermeUC(CompteurAnnee)))), .PM_ClotureUC(CompteurAnnee))
        End If
'         .PM_UC(CompteurAnnee) = Max(0, .PM_ClotureUC(CompteurAnnee) - .ChargPM_UC(CompteurAnnee) - .RetroGlobalPM_UC(CompteurAnnee)) '- .Commissions_PMuc(CompteurAnnee)   Estelle 8 / 17
         
         
        'FeuilAct.Cells(NumLgn + 1, 5) = .PM_UC(CompteurAnnee)
        
    End With
Next NumLgn


End Sub

'*******************************************************************************
Sub CalculsMargUC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    If NumLgn = 13586 Then
        NumLgn = NumLgn
    End If
    With BDD(NumLgn)
        .MargUC(CompteurAnnee) = .PrimeCommUC(CompteurAnnee) + .PM_UC(CompteurAnnee - 1) + .InteretsPM_MiPeriodeUC(CompteurAnnee) + _
                                 .InteretsPrimeUC(CompteurAnnee) - .SinDecesUC(CompteurAnnee) - .SinTirageUC(CompteurAnnee) - _
                                 .SinRachatTotUC(CompteurAnnee) - .SinRachatPartUC(CompteurAnnee) - .SinTermeUC(CompteurAnnee) + _
                                 .InteretsFinPeriodeUC(CompteurAnnee) - .PM_UC(CompteurAnnee) - .Cap_UC_Euro(CompteurAnnee) + _
                                 .Cap_Euro_UC(CompteurAnnee) - .Commissions_PMuc(CompteurAnnee) - .Commissions_PrimesUC(CompteurAnnee) _
                                 + .RendUCPrime(CompteurAnnee) + .RendUC_MiPeriodeUC(CompteurAnnee) + .RendFinPeriodeUC(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsChargUC()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .ChargUC(CompteurAnnee) = .ChargPrimeUC(CompteurAnnee) + .ChargDecesUC(CompteurAnnee) + .ChargTirageUC(CompteurAnnee) + _
                                  .ChargRachatTotUC(CompteurAnnee) + .ChargRachatPartUC(CompteurAnnee) + .ChargTermeUC(CompteurAnnee) + _
                                  .ChargPM_UC(CompteurAnnee) + .RetroGlobalPM_UC(CompteurAnnee) + .ChargTransf_UC_Euro(CompteurAnnee) + .ChargTransf_UC_UC(CompteurAnnee) - _
                                  .Commissions_PMuc(CompteurAnnee) - .Commissions_PrimesUC(CompteurAnnee)   ' Estelle 08/17
    End With
Next NumLgn

End Sub

'*******************************************************************************
'******************Option Prévoyance (Décès temporaire)************
'******************************************************************
'******************************************************************
Sub CalculsPrime_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
         .PrimeCommPrev(CompteurAnnee) = .CoeffPrimePrev * .PrimeCommPrev(CompteurAnnee - 1) * .NbClotures_Prev(CompteurAnnee) / .NbTetes  'Il ny a de versement de prime qu'au cours de la première année du contrat
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsChargPrime_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        
        .ChargPrimePrev(CompteurAnnee) = .PrimeCommPrev(CompteurAnnee) * ChPrime_Prev(.NomProd, .AncienneteContrat(CompteurAnnee))
        
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsCommissionPrime_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        
        .Commissions_PrimesPrev(CompteurAnnee) = .PrimeCommPrev(CompteurAnnee) * .Taux_Com_Prev
    
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsPrimeNette_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        
        .PrimeNettePrev(CompteurAnnee) = .PrimeCommPrev(CompteurAnnee) - .ChargPrimePrev(CompteurAnnee) '- .Commissions_PrimesPrev(CompteurAnnee)
    
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsSinistres_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        
         .SinDecesPrev(CompteurAnnee) = .CapitalDeces(CompteurAnnee) * _
                                        (qxChoques(NumLgn, CompteurAnnee, NumChoc) / (1 + ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee)))) * .NbClotures_Prev(CompteurAnnee - 1)
         .SinRachatTotPrev(CompteurAnnee) = 0
    
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsChargementsSinistres_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
         .ChargDecesPrev(CompteurAnnee) = .SinDecesPrev(CompteurAnnee) * ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee))
         .ChargRachatTotPrev(CompteurAnnee) = .SinRachatTotPrev(CompteurAnnee) * ChRachatTot(.NomProd, .AncienneteContrat(CompteurAnnee))
         
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CaculsChargPM_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
         .ChargPM_Prev(CompteurAnnee) = ChPM_Prev(.NomProd, .AncienneteContrat(CompteurAnnee)) * (.PM_Prev(CompteurAnnee - 1) + _
                                        0.5 * (.PrimeNettePrev(CompteurAnnee) - .SinDecesPrev(CompteurAnnee) - .SinRachatTotPrev(CompteurAnnee)))
        
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsCommissionPM_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .Commissions_PMprev(CompteurAnnee) = .TxComSurEncours_Prev * (.PM_Prev(CompteurAnnee - 1) + _
                                             0.5 * (.PrimeNettePrev(CompteurAnnee) - .SinDecesPrev(CompteurAnnee) - _
                                             .SinRachatTotUC(CompteurAnnee)))
    End With
Next NumLgn
End Sub

'*******************************************************************************
Sub CalculsPM_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
         If .AgeAssure(CompteurAnnee) > .AgeMaxprev Then
           .PM_Prev(CompteurAnnee) = 0
         Else
           .PM_Prev(CompteurAnnee) = .CapitalDeces(CompteurAnnee) * ((TableCommut(TransformTMGprev(.TMGprev), .AgeAssure(CompteurAnnee)).Mx - _
                                     TableCommut(TransformTMGprev(.TMGprev), .AgeAssure(CompteurAnnee) + .DureeRestante_Prev(CompteurAnnee)).Mx) / _
                                     TableCommut(TransformTMGprev(.TMGprev), .AgeAssure(CompteurAnnee)).Dx)
        End If
    End With
       
Next NumLgn
End Sub

'*******************************************************************************
Sub CalculsMarg_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats

    With BDD(NumLgn)
        .MargPrev(CompteurAnnee) = .PrimeCommPrev(CompteurAnnee) + .PM_Prev(CompteurAnnee - 1) - .SinDecesPrev(CompteurAnnee) - _
                                  .SinRachatTotPrev(CompteurAnnee) - .PM_Prev(CompteurAnnee) - .Commissions_PMprev(CompteurAnnee) - _
                                  .Commissions_PrimesPrev(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsCharg_Prev()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        .ChargPrev(CompteurAnnee) = .ChargPrimePrev(CompteurAnnee) + .ChargDecesPrev(CompteurAnnee) + .ChargRachatTotPrev(CompteurAnnee) + .ChargPM_Prev(CompteurAnnee)
    End With
Next NumLgn

End Sub

'*******************************************************************************
Sub CalculsSommes()
'*******************************************************************************

Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    '###########################################################
    '#################### HORS PREVOYANCE ######################
    '###########################################################
    If Perimetre = "NB" And BDD(NumLgn).AnneeEffet >= AnneeValorisation Or Perimetre = "Stock" And BDD(NumLgn).AnneeEffet < AnneeValorisation Or Perimetre = "Stock+NB" Then
        If TypePrime = "PP+PU" Or TypePrime = "PP" And BDD(NumLgn).Periodicite <> "Unique" Or TypePrime = "PU" And BDD(NumLgn).Periodicite = "Unique" Then
           If Obsèque_Epargne = "Total" Or Obsèque_Epargne = "Obsèque" And BDD(NumLgn).IndicObsèque = True Or Obsèque_Epargne = "Epargne" And BDD(NumLgn).IndicObsèque = False Then
                If TypeTauxMin = "Total" Or TypeTauxMin = "3" And BDD(NumLgn).IndicTypeTauxMin = "3" Or TypeTauxMin = "Autres" And BDD(NumLgn).IndicTypeTauxMin <> "3" Then
                    If ProjParTMG = 0 And Donnees(NumLgn).TMG = 0 Or ProjParTMG = 3.5 / 100 And Donnees(NumLgn).TMG = 3.5 / 100 Or ProjParTMG = 4.5 / 100 And Donnees(NumLgn).TMG = 4.5 / 100 Or ProjParTMG = -1 Then
                        With Totaux_Par_MP(BDD(NumLgn).ModelPoint, CompteurAnnee)
                            .Somm_BonusRetEuro = .Somm_BonusRetEuro + BDD(NumLgn).BonusRetEuro(CompteurAnnee)
                            .Somm_BonusRetUC = .Somm_BonusRetUC + BDD(NumLgn).BonusRetUC(CompteurAnnee)
                            .Somm_PrimeCommEuro = .Somm_PrimeCommEuro + BDD(NumLgn).PrimeCommEuro(CompteurAnnee)
                            .Somm_ChargPrimeEuro = .Somm_ChargPrimeEuro + BDD(NumLgn).ChargPrimeEuro(CompteurAnnee)
                            .Somm_PM_OuvertureEuro = .Somm_PM_OuvertureEuro + BDD(NumLgn).PM_Euro(CompteurAnnee - 1)
                            .Somm_InteretsPM_MiPeriodeEuro = .Somm_InteretsPM_MiPeriodeEuro + BDD(NumLgn).InteretsPM_MiPeriodeEuro(CompteurAnnee)
                            .Somm_PM_MiPeriode1Euro = .Somm_PM_MiPeriode1Euro + BDD(NumLgn).PM_MiPeriode1Euro(CompteurAnnee)
                            .Somm_SinDecesEuro = .Somm_SinDecesEuro + BDD(NumLgn).SinDecesEuro(CompteurAnnee)
                            .Somm_SinTirageEuro = .Somm_SinTirageEuro + BDD(NumLgn).SinTirageEuro(CompteurAnnee)
                            .Somm_SinRachatTotEuro = .Somm_SinRachatTotEuro + BDD(NumLgn).SinRachatTotEuro(CompteurAnnee)
                            .Somm_SinRachatPartEuro = .Somm_SinRachatPartEuro + BDD(NumLgn).SinRachatPartEuro(CompteurAnnee)
                            .Somm_ChargDecesEuro = .Somm_ChargDecesEuro + BDD(NumLgn).ChargDecesEuro(CompteurAnnee)
                            .Somm_ChargTirageEuro = .Somm_ChargTirageEuro + BDD(NumLgn).ChargTirageEuro(CompteurAnnee)
                            .Somm_ChargRachatTotEuro = .Somm_ChargRachatTotEuro + BDD(NumLgn).ChargRachatTotEuro(CompteurAnnee)
                            .Somm_ChargRachatPartEuro = .Somm_ChargRachatPartEuro + BDD(NumLgn).ChargRachatPartEuro(CompteurAnnee)
                            .Somm_PM_MiPeriode2Euro = .Somm_PM_MiPeriode2Euro + BDD(NumLgn).PM_MiPeriode2Euro(CompteurAnnee)
                            .Somm_SinTermeEuro = .Somm_SinTermeEuro + BDD(NumLgn).SinTermeEuro(CompteurAnnee)
                            .Somm_ChargTermeEuro = .Somm_ChargTermeEuro + BDD(NumLgn).ChargTermeEuro(CompteurAnnee)
                            .Somm_PM_MiPeriode3Euro = .Somm_PM_MiPeriode3Euro + BDD(NumLgn).PM_MiPeriode3Euro(CompteurAnnee)
                            .Somm_PM_MiPeriode4Euro = .Somm_PM_MiPeriode4Euro + BDD(NumLgn).PM_MiPeriode4Euro(CompteurAnnee)
                            .Somm_InteretsFinPeriodeEuro = .Somm_InteretsFinPeriodeEuro + BDD(NumLgn).InteretsFinPeriodeEuro(CompteurAnnee)
                            .Somm_PM_ClotureEuro = .Somm_PM_ClotureEuro + BDD(NumLgn).PM_ClotureEuro(CompteurAnnee)
                            .Somm_ChargPM_Euro = .Somm_ChargPM_Euro + BDD(NumLgn).ChargPM_Euro(CompteurAnnee)
                            .Somm_PrimeCommUC = .Somm_PrimeCommUC + BDD(NumLgn).PrimeCommUC(CompteurAnnee)
                            .Somm_ChargPrimeUC = .Somm_ChargPrimeUC + BDD(NumLgn).ChargPrimeUC(CompteurAnnee)
                            .Somm_PM_OuvertureUC = .Somm_PM_OuvertureUC + BDD(NumLgn).PM_UC(CompteurAnnee - 1)
                            .Somm_InteretsPM_MiPeriodeUC = .Somm_InteretsPM_MiPeriodeUC + BDD(NumLgn).InteretsPM_MiPeriodeUC(CompteurAnnee)
                            .Somm_PM_MiPeriode1UC = .Somm_PM_MiPeriode1UC + BDD(NumLgn).PM_MiPeriode1UC(CompteurAnnee)
                            .Somm_SinDecesUC = .Somm_SinDecesUC + BDD(NumLgn).SinDecesUC(CompteurAnnee)
                            .Somm_SinTirageUC = .Somm_SinTirageUC + BDD(NumLgn).SinTirageUC(CompteurAnnee)
                            .Somm_SinRachatTotUC = .Somm_SinRachatTotUC + BDD(NumLgn).SinRachatTotUC(CompteurAnnee)
                            .Somm_SinRachatPartUC = .Somm_SinRachatPartUC + BDD(NumLgn).SinRachatPartUC(CompteurAnnee)
                            .Somm_ChargDecesUC = .Somm_ChargDecesUC + BDD(NumLgn).ChargDecesUC(CompteurAnnee)
                            .Somm_ChargTirageUC = .Somm_ChargTirageUC + BDD(NumLgn).ChargTirageUC(CompteurAnnee)
                            .Somm_ChargRachatTotUC = .Somm_ChargRachatTotUC + BDD(NumLgn).ChargRachatTotUC(CompteurAnnee)
                            .Somm_ChargRachatPartUC = .Somm_ChargRachatPartUC + BDD(NumLgn).ChargRachatPartUC(CompteurAnnee)
                            .Somm_PM_MiPeriode2UC = .Somm_PM_MiPeriode2UC + BDD(NumLgn).PM_MiPeriode2UC(CompteurAnnee)
                            .Somm_SinTermeUC = .Somm_SinTermeUC + BDD(NumLgn).SinTermeUC(CompteurAnnee)
                            .Somm_ChargTermeUC = .Somm_ChargTermeUC + BDD(NumLgn).ChargTermeUC(CompteurAnnee)
                            .Somm_PM_MiPeriode3UC = .Somm_PM_MiPeriode3UC + BDD(NumLgn).PM_MiPeriode3UC(CompteurAnnee)
                            .Somm_PM_MiPeriode4UC = .Somm_PM_MiPeriode4UC + BDD(NumLgn).PM_MiPeriode4UC(CompteurAnnee)
                            .Somm_InteretsFinPeriodeUC = .Somm_InteretsFinPeriodeUC + BDD(NumLgn).InteretsFinPeriodeUC(CompteurAnnee)
                            .Somm_PM_ClotureUC = .Somm_PM_ClotureUC + BDD(NumLgn).PM_ClotureUC(CompteurAnnee)
                            .Somm_ChargPM_UC = .Somm_ChargPM_UC + BDD(NumLgn).ChargPM_UC(CompteurAnnee)
                            .Somm_RetroGlobalPM_UC = .Somm_RetroGlobalPM_UC + BDD(NumLgn).RetroGlobalPM_UC(CompteurAnnee)
                            .Somm_RetroAEPM_UC = .Somm_RetroAEPM_UC + BDD(NumLgn).RetroAEPM_UC(CompteurAnnee)
                            .Somm_Cap_Euro_UC = .Somm_Cap_Euro_UC + BDD(NumLgn).Cap_Euro_UC(CompteurAnnee)
                            .Somm_ChargTransf_Euro_UC = .Somm_ChargTransf_Euro_UC + BDD(NumLgn).ChargTransf_Euro_UC(CompteurAnnee)
                            .Somm_Cap_UC_Euro = .Somm_Cap_UC_Euro + BDD(NumLgn).Cap_UC_Euro(CompteurAnnee)
                            .Somm_ChargTransf_UC_Euro = .Somm_ChargTransf_UC_Euro + BDD(NumLgn).ChargTransf_UC_Euro(CompteurAnnee)
                            .Somm_ChargTransf_UC_UC = .Somm_ChargTransf_UC_UC + BDD(NumLgn).ChargTransf_UC_UC(CompteurAnnee)
                            .Somm_NbOuvertures = .Somm_NbOuvertures + BDD(NumLgn).NbClotures(CompteurAnnee - 1)
                            .Somm_NbOuverturesEuro = .Somm_NbOuverturesEuro + BDD(NumLgn).NbClotures_Euro(CompteurAnnee - 1)
                            .Somm_NbOuverturesUC = .Somm_NbOuverturesUC + BDD(NumLgn).NbClotures_UC(CompteurAnnee - 1)
                            .Somm_NbDeces = .Somm_NbDeces + BDD(NumLgn).NbDeces(CompteurAnnee)
                            .Somm_NbTirages = .Somm_NbTirages + BDD(NumLgn).NbTirages(CompteurAnnee)
                            .Somm_NbRachatsTot = .Somm_NbRachatsTot + BDD(NumLgn).NbRachatsTot(CompteurAnnee)
                            .Somm_NbRachatsPart = .Somm_NbRachatsPart + BDD(NumLgn).NbRachatsPart(CompteurAnnee)
                            .Somm_NbTermes = .Somm_NbTermes + BDD(NumLgn).NbTermes(CompteurAnnee)
                            .Somm_Commissions_PrimesEuro = .Somm_Commissions_PrimesEuro + BDD(NumLgn).Commissions_PrimesEuro(CompteurAnnee)
                            .Somm_Commissions_PMeuro = .Somm_Commissions_PMeuro + BDD(NumLgn).Commissions_PMeuro(CompteurAnnee)
                            .Somm_Commissions_PrimesUC = .Somm_Commissions_PrimesUC + BDD(NumLgn).Commissions_PrimesUC(CompteurAnnee)
                            .Somm_Commissions_PMuc = .Somm_Commissions_PMuc + BDD(NumLgn).Commissions_PMuc(CompteurAnnee)
                            .Somm_FraisAdmEuro = .Somm_FraisAdmEuro + BDD(NumLgn).CoutUnitaireGestion_Euro(CompteurAnnee)
                            .Somm_FraisAdmUC = .Somm_FraisAdmUC + BDD(NumLgn).CoutUnitaireGestion_UC(CompteurAnnee)
                            .Somm_FraisPrestEuro = .Somm_FraisPrestEuro + BDD(NumLgn).CoutUnitairePresta_Euro(CompteurAnnee)
                            .Somm_FraisPrestUC = .Somm_FraisPrestUC + BDD(NumLgn).CoutUnitairePresta_UC(CompteurAnnee)
                        End With
                    
                    '###########################################################
                    '####################### PREVOYANCE ########################
                    '###########################################################
                        With Totaux_Par_MP(10, CompteurAnnee)                                                                  'Pour le MP "AINV option prévoyance" créé uniquement dans la macro
                                                                                                                               'Les calculs sont effectués en partie dans la catégorie euro
                            .Somm_PrimeCommPrev = .Somm_PrimeCommPrev + BDD(NumLgn).PrimeCommPrev(CompteurAnnee) '***
                            .Somm_ChargPrimePrev = .Somm_ChargPrimePrev + BDD(NumLgn).ChargPrimePrev(CompteurAnnee) '***
                            .Somm_PM_CloturePrev = .Somm_PM_CloturePrev + BDD(NumLgn).PM_Prev(CompteurAnnee)  '***
                            .Somm_SinDecesEuro = .Somm_SinDecesEuro + BDD(NumLgn).SinDecesPrev(CompteurAnnee)
                            .Somm_SinRachatTotEuro = .Somm_SinRachatTotEuro + BDD(NumLgn).CapitalDeces(CompteurAnnee) * BDD(NumLgn).NbClotures_Prev(CompteurAnnee - 1) * _
                                                      BDR(BDD(NumLgn).CatRachatTot, BDD(NumLgn).AncienneteContrat(CompteurAnnee)).ProbaRachatTot_Prev         'Sinistre Rachat Total
                                                                                                                                                          '= Taux rachat * Capital ouverture
                                                                                                                    
                            .Somm_SinTermeEuro = .Somm_SinTermeEuro + BDD(NumLgn).CapitalDeces(CompteurAnnee) * BDD(NumLgn).NbTermes_Prev(CompteurAnnee) 'Sinistre Terme Contrat
                                                                                                                                                          '= Indicatrice terme (= Taux Terme) * Capital ouverture *(1-qx-TxRachatTot)
                                                                                               
                            .Somm_ChargDecesEuro = .Somm_ChargDecesEuro + BDD(NumLgn).ChargDecesPrev(CompteurAnnee)
                            .Somm_ChargRachatTotEuro = .Somm_ChargRachatTotEuro + BDD(NumLgn).ChargRachatTotPrev(CompteurAnnee)
                            .Somm_ChargPM_Euro = .Somm_ChargPM_Euro + BDD(NumLgn).ChargPM_Prev(CompteurAnnee)
                            .Somm_NbOuverturesPrev = .Somm_NbOuverturesPrev + BDD(NumLgn).NbClotures_Prev(CompteurAnnee - 1) '***
                            .Somm_Commissions_PrimesEuro = .Somm_Commissions_PrimesEuro + BDD(NumLgn).Commissions_PrimesPrev(CompteurAnnee)
                            .Somm_Commissions_PMeuro = .Somm_Commissions_PMeuro + BDD(NumLgn).Commissions_PMprev(CompteurAnnee)
                            .Somm_Capital_Garanti = .Somm_Capital_Garanti + BDD(NumLgn).CapitalDeces(CompteurAnnee) * BDD(NumLgn).NbClotures_Prev(CompteurAnnee - 1)
                            .Somm_PM_MiPeriode1Euro = .Somm_PM_MiPeriode1Euro + BDD(NumLgn).CapitalDeces(CompteurAnnee) * BDD(NumLgn).NbClotures_Prev(CompteurAnnee - 1) 'Somme PMmipériode1= Somme des capitaux ouverture
                            .Somm_PM_MiPeriode2Euro = .Somm_PM_MiPeriode2Euro + BDD(NumLgn).CapitalDeces(CompteurAnnee) * BDD(NumLgn).NbClotures_Prev(CompteurAnnee - 1) 'Somme PMmipériode2= Somme des capitaux ouverture
                        End With
                    End If
                End If
            End If
        End If
    End If
Next NumLgn
        
End Sub

'*******************************************************************************
Sub CalculsSommesPourAffichage()
'*******************************************************************************

Dim NumLgn As Double


For NumLgn = 1 To NbContrats
    If Perimetre = "NB" And BDD(NumLgn).AnneeEffet >= AnneeValorisation Or Perimetre = "Stock" And BDD(NumLgn).AnneeEffet < AnneeValorisation Or Perimetre = "Stock+NB" Then
        If TypePrime = "PP+PU" Or TypePrime = "PP" And BDD(NumLgn).Periodicite <> "Unique" Or TypePrime = "PU" And BDD(NumLgn).Periodicite = "Unique" Then
            If Obsèque_Epargne = "Total" Or Obsèque_Epargne = "Obsèque" And BDD(NumLgn).IndicObsèque = True Or Obsèque_Epargne = "Epargne" And BDD(NumLgn).IndicObsèque = False Then
                If TypeTauxMin = "Total" Or TypeTauxMin = "3" And BDD(NumLgn).IndicTypeTauxMin = "3" Or TypeTauxMin = "Autres" And BDD(NumLgn).IndicTypeTauxMin <> "3" Then
                    If ProjParTMG = 0 And Donnees(NumLgn).TMG = 0 / 100 Or ProjParTMG = 3.5 / 100 And Donnees(NumLgn).TMG = 3.5 / 100 Or ProjParTMG = 4.5 / 100 And Donnees(NumLgn).TMG = 4.5 / 100 Or ProjParTMG = -1 Then
                        With Totaux_Pour_Affichage(CompteurAnnee)
                            If CompteurAnnee = 1 Then
                                Totaux_Pour_Affichage(0).Somm_PrimeCommEuro = Totaux_Pour_Affichage(0).Somm_PrimeCommEuro + BDD(NumLgn).PrimeCommEuroPourAffichageAnnee0
                                Totaux_Pour_Affichage(0).Somm_PrimeCommUC = Totaux_Pour_Affichage(0).Somm_PrimeCommUC + BDD(NumLgn).PrimeCommUCPourAffichageAnnee0
                                Totaux_Pour_Affichage(0).Somm_ChargPrimeEuro = Totaux_Pour_Affichage(0).Somm_ChargPrimeEuro + BDD(NumLgn).ChargPrimeEuro(0)
                                Totaux_Pour_Affichage(0).Somm_ChargPrimeUC = Totaux_Pour_Affichage(0).Somm_ChargPrimeUC + BDD(NumLgn).ChargPrimeUC(0)
                                Totaux_Pour_Affichage(0).Somm_Commissions_PrimesEuro = Totaux_Pour_Affichage(0).Somm_Commissions_PrimesEuro + BDD(NumLgn).Commissions_PrimesEuro(0)
                                Totaux_Pour_Affichage(0).Somm_Commissions_PrimesUC = Totaux_Pour_Affichage(0).Somm_Commissions_PrimesUC + BDD(NumLgn).Commissions_PrimesUC(0)
                            End If
                            .Somm_PrimeCommEuro = .Somm_PrimeCommEuro + BDD(NumLgn).PrimeCommEuro(CompteurAnnee)
                            .Somm_ChargPrimeEuro = .Somm_ChargPrimeEuro + BDD(NumLgn).ChargPrimeEuro(CompteurAnnee)
                            .Somm_PM_OuvertureEuro = .Somm_PM_OuvertureEuro + BDD(NumLgn).PM_Euro(CompteurAnnee - 1)
                            .Somm_SinDecesEuro = .Somm_SinDecesEuro + BDD(NumLgn).SinDecesEuro(CompteurAnnee)
                            .Somm_SinTirageEuro = .Somm_SinTirageEuro + BDD(NumLgn).SinTirageEuro(CompteurAnnee)
                            .Somm_SinRachatTotEuro = .Somm_SinRachatTotEuro + BDD(NumLgn).SinRachatTotEuro(CompteurAnnee)
                            .Somm_SinRachatPartEuro = .Somm_SinRachatPartEuro + BDD(NumLgn).SinRachatPartEuro(CompteurAnnee)
                            .Somm_ChargDecesEuro = .Somm_ChargDecesEuro + BDD(NumLgn).ChargDecesEuro(CompteurAnnee)
                            .Somm_ChargTirageEuro = .Somm_ChargTirageEuro + BDD(NumLgn).ChargTirageEuro(CompteurAnnee)
                            .Somm_ChargRachatTotEuro = .Somm_ChargRachatTotEuro + BDD(NumLgn).ChargRachatTotEuro(CompteurAnnee)
                            .Somm_ChargRachatPartEuro = .Somm_ChargRachatPartEuro + BDD(NumLgn).ChargRachatPartEuro(CompteurAnnee)
                            .Somm_SinTermeEuro = .Somm_SinTermeEuro + BDD(NumLgn).SinTermeEuro(CompteurAnnee)
                            .Somm_ChargTermeEuro = .Somm_ChargTermeEuro + BDD(NumLgn).ChargTermeEuro(CompteurAnnee)
                            .Somm_PM_ClotureEuro = .Somm_PM_ClotureEuro + BDD(NumLgn).PM_Euro(CompteurAnnee)
                            .Somm_ChargPM_Euro = .Somm_ChargPM_Euro + BDD(NumLgn).ChargPM_Euro(CompteurAnnee)
    
                            .Somm_PrimeCommUC = .Somm_PrimeCommUC + BDD(NumLgn).PrimeCommUC(CompteurAnnee)
                            .Somm_ChargPrimeUC = .Somm_ChargPrimeUC + BDD(NumLgn).ChargPrimeUC(CompteurAnnee)
                            .Somm_PM_OuvertureUC = .Somm_PM_OuvertureUC + BDD(NumLgn).PM_UC(CompteurAnnee - 1)
                            .Somm_SinDecesUC = .Somm_SinDecesUC + BDD(NumLgn).SinDecesUC(CompteurAnnee)
                            .Somm_SinTirageUC = .Somm_SinTirageUC + BDD(NumLgn).SinTirageUC(CompteurAnnee)
                            .Somm_SinRachatTotUC = .Somm_SinRachatTotUC + BDD(NumLgn).SinRachatTotUC(CompteurAnnee)
                            .Somm_SinRachatPartUC = .Somm_SinRachatPartUC + BDD(NumLgn).SinRachatPartUC(CompteurAnnee)
                            .Somm_ChargDecesUC = .Somm_ChargDecesUC + BDD(NumLgn).ChargDecesUC(CompteurAnnee)
                            .Somm_ChargTirageUC = .Somm_ChargTirageUC + BDD(NumLgn).ChargTirageUC(CompteurAnnee)
                            .Somm_ChargRachatTotUC = .Somm_ChargRachatTotUC + BDD(NumLgn).ChargRachatTotUC(CompteurAnnee)
                            .Somm_ChargRachatPartUC = .Somm_ChargRachatPartUC + BDD(NumLgn).ChargRachatPartUC(CompteurAnnee)
                            .Somm_SinTermeUC = .Somm_SinTermeUC + BDD(NumLgn).SinTermeUC(CompteurAnnee)
                            .Somm_ChargTermeUC = .Somm_ChargTermeUC + BDD(NumLgn).ChargTermeUC(CompteurAnnee)
                            .Somm_PM_ClotureUC = .Somm_PM_ClotureUC + BDD(NumLgn).PM_UC(CompteurAnnee)
                            .Somm_ChargPM_UC = .Somm_ChargPM_UC + BDD(NumLgn).ChargPM_UC(CompteurAnnee)
                            .Somm_RetroGlobalPM_UC = .Somm_RetroGlobalPM_UC + BDD(NumLgn).RetroGlobalPM_UC(CompteurAnnee)
                            .Somm_RetroAEPM_UC = .Somm_RetroAEPM_UC + BDD(NumLgn).RetroAEPM_UC(CompteurAnnee)
                            .Somm_Rend_MiPeriodeUC = .Somm_Rend_MiPeriodeUC + BDD(NumLgn).RendUC_MiPeriodeUC(CompteurAnnee) + BDD(NumLgn).RendUCPrime(CompteurAnnee)
                            .Somm_RendFinPeriodeUC = .Somm_RendFinPeriodeUC + BDD(NumLgn).RendFinPeriodeUC(CompteurAnnee)
                            
                            .Somm_Commissions_PrimesEuro = .Somm_Commissions_PrimesEuro + BDD(NumLgn).Commissions_PrimesEuro(CompteurAnnee)
                            .Somm_Commissions_PMeuro = .Somm_Commissions_PMeuro + BDD(NumLgn).Commissions_PMeuro(CompteurAnnee)
                            .Somm_Commissions_PrimesUC = .Somm_Commissions_PrimesUC + BDD(NumLgn).Commissions_PrimesUC(CompteurAnnee)
                            .Somm_Commissions_PMuc = .Somm_Commissions_PMuc + BDD(NumLgn).Commissions_PMuc(CompteurAnnee)
                            .Somm_BonusRetEuro = .Somm_BonusRetEuro + BDD(NumLgn).BonusRetEuro(CompteurAnnee)
                            .Somm_BonusRetUC = .Somm_BonusRetUC + BDD(NumLgn).BonusRetUC(CompteurAnnee)
                            
                            .Somm_InteretsPM_MiPeriodeEuro = .Somm_InteretsPM_MiPeriodeEuro + BDD(NumLgn).InteretsPM_MiPeriodeEuro(CompteurAnnee) + BDD(NumLgn).InteretsPrimeEuro(CompteurAnnee)
                            .Somm_PB_MiPeriodeEuro = .Somm_PB_MiPeriodeEuro + BDD(NumLgn).PB_MIPeriodeEuro(CompteurAnnee) + BDD(NumLgn).PBPrimeEuro(CompteurAnnee)
                            .Somm_PM_MiPeriode1Euro = .Somm_PM_MiPeriode1Euro + BDD(NumLgn).PM_MiPeriode1Euro(CompteurAnnee)
                            .Somm_InteretsFinPeriodeEuro = .Somm_InteretsFinPeriodeEuro + BDD(NumLgn).InteretsFinPeriodeEuro(CompteurAnnee)
                            .Somm_PBFinPeriodeEuro = .Somm_PBFinPeriodeEuro + BDD(NumLgn).PBFinPeriodeEuro(CompteurAnnee)
                            
                            .Somm_NbOuverturesEuro = .Somm_NbOuverturesEuro + BDD(NumLgn).NbClotures_Euro(CompteurAnnee - 1)
                            .Somm_NbOuverturesUC = .Somm_NbOuverturesUC + BDD(NumLgn).NbClotures_UC(CompteurAnnee - 1)
                            .Somm_FraisAdmEuro = .Somm_FraisAdmEuro + BDD(NumLgn).CoutUnitaireGestion_Euro(CompteurAnnee)
                            .Somm_FraisAdmUC = .Somm_FraisAdmUC + BDD(NumLgn).CoutUnitaireGestion_UC(CompteurAnnee)
                            .Somm_FraisPrestEuro = .Somm_FraisPrestEuro + BDD(NumLgn).CoutUnitairePresta_Euro(CompteurAnnee)
                            .Somm_FraisPrestUC = .Somm_FraisPrestUC + BDD(NumLgn).CoutUnitairePresta_UC(CompteurAnnee)
                            If BDD(NumLgn).PM_Euro(0) > 0 Then
                                .Somm_NbDecesEuro = .Somm_NbDecesEuro + BDD(NumLgn).NbDeces(CompteurAnnee)
                                .Somm_NbTiragesEuro = .Somm_NbTiragesEuro + BDD(NumLgn).NbTirages(CompteurAnnee)
                                .Somm_NbRachatsTotEuro = .Somm_NbRachatsTotEuro + BDD(NumLgn).NbRachatsTot(CompteurAnnee)
                                .Somm_NbRachatsPartEuro = .Somm_NbRachatsPartEuro + BDD(NumLgn).NbRachatsPart(CompteurAnnee)
                                .Somm_NbTermesEuro = .Somm_NbTermesEuro + BDD(NumLgn).NbTermes(CompteurAnnee)
                            End If
                            
                            If BDD(NumLgn).PM_UC(0) > 0 Then
                                .Somm_NbDecesUC = .Somm_NbDecesUC + BDD(NumLgn).NbDeces(CompteurAnnee)
                                .Somm_NbTiragesUC = .Somm_NbTiragesUC + BDD(NumLgn).NbTirages(CompteurAnnee)
                                .Somm_NbRachatsTotUC = .Somm_NbRachatsTotUC + BDD(NumLgn).NbRachatsTot(CompteurAnnee)
                                .Somm_NbRachatsPartUC = .Somm_NbRachatsPartUC + BDD(NumLgn).NbRachatsPart(CompteurAnnee)
                                .Somm_NbTermesUC = .Somm_NbTermesUC + BDD(NumLgn).NbTermes(CompteurAnnee)
                            End If
                        End With
                    End If
                End If
            End If
        End If
    End If
Next NumLgn

End Sub

'*******************************************************************************
Public Sub CalculsRatiosFlexing()
'*******************************************************************************

Dim NumModelPoint As Integer

For NumModelPoint = 1 To NbModelPoint
    With RDF(NumModelPoint, CompteurAnnee, NumChoc)
        .NbOuvertures = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_NbOuvertures
        .NbOuverturesEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_NbOuverturesEuro
        .NbOuverturesUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_NbOuverturesUC
        .NbOuverturesPrev = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_NbOuverturesPrev '***
        .PrimeCommEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PrimeCommEuro
        .PM_OuvertureEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_OuvertureEuro
        .PrimeCommUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PrimeCommUC
        .PM_OuvertureUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_OuvertureUC
        .PrimeCommPrev = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PrimeCommPrev '***
        .PM_CloturePrev = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_CloturePrev  '***
        
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesEuro = 0 Then
            .TxChargDecesEuro = 0
        Else: .TxChargDecesEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargDecesEuro / _
                                  Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesEuro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro = 0 Then
            .ProbaDecesEuro = 0
        Else: .ProbaDecesEuro = (1 + .TxChargDecesEuro) * _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesEuro / _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro   'Pour les options prévoyance :
                                                                                                      'PM_MiPeriode1Euro = Capital ouverture
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageEuro = 0 Then
            .TxChargTirageEuro = 0
        Else: .TxChargTirageEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargTirageEuro / _
                                   Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageEuro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro = 0 Then
            .TxTirageEuro = 0
        Else: .TxTirageEuro = (1 + .TxChargTirageEuro) * _
                              Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageEuro / _
                              Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotEuro = 0 Then
            .TxChargRachatTotEuro = 0
        Else: .TxChargRachatTotEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargRachatTotEuro / _
                                      Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotEuro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro = 0 Then
            .ProbaRachatTotEuro = 0
        Else: .ProbaRachatTotEuro = (1 + .TxChargRachatTotEuro) * _
                                    Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotEuro / _
                                    Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro     'Pour les options prévoyance :
                                                                                                            'PM_MiPeriode1Euro = Capital ouverture
                                                                                                            'SinRachatTotEuro = Kouv * TxRachat
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartEuro = 0 Then
            .TxChargRachatPartEuro = 0
        Else: .TxChargRachatPartEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargRachatPartEuro / _
                                       Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartEuro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro = 0 Then
            .ProbaRachatPartEuro = 0
        Else: .ProbaRachatPartEuro = (1 + .TxChargRachatPartEuro) * _
                                     Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartEuro / _
                                     Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1Euro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeEuro = 0 Then
            .TxChargContratsTermesEuro = 0
        Else: .TxChargContratsTermesEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargTermeEuro / _
                                           Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeEuro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode2Euro = 0 Then
            .ProbaContratsTermesEuro = 0
        Else: .ProbaContratsTermesEuro = (1 + .TxChargContratsTermesEuro) * _
                                         Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeEuro / _
                                         Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode2Euro  'Pour les options prévoyance :
                                                                                                              'PM_MiPeriode2Euro = Capital ouverture
                                                                                                              'SinTermeEuro = TxTerme*Kouv*(1-qx-TxRachatTot)
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureEuro = 0 Then
            .TxChargPM_Euro = 0
        Else: .TxChargPM_Euro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargPM_Euro / _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureEuro
        End If
        If .PM_OuvertureEuro = 0 Then
            .TxTechDebutPeriodeEuro = 0
        Else: .TxTechDebutPeriodeEuro = (1 + Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_InteretsPM_MiPeriodeEuro / _
                                             .PM_OuvertureEuro) ^ 2 - 1
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode4Euro = 0 Then
            .TxTechFinPeriodeEuro = 0
        Else: .TxTechFinPeriodeEuro = (1 + Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_InteretsFinPeriodeEuro / _
                                           Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode4Euro) ^ 2 - 1
        End If
        '######Ancienne formule fausse######
        If CompteurAnnee = 1 Then
            .EvolutionPrimesEuro = 0
        Else
            If RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).PrimeCommEuro = 0 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaDecesEuro = 1 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).TxTirageEuro = 1 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaRachatTotEuro = 1 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaContratsTermesEuro = 1 Then
                    .EvolutionPrimesEuro = 0
            Else: .EvolutionPrimesEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PrimeCommEuro / _
                                         (RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).PrimeCommEuro * _
                                          (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaDecesEuro) * _
                                          (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).TxTirageEuro) * _
                                          (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaRachatTotEuro) * _
                                          (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaContratsTermesEuro))
            End If
        End If
        '###################################
        If .PrimeCommEuro = 0 Then
            .TxChargPrimesEuro = 0
        Else: .TxChargPrimesEuro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargPrimeEuro / _
                                   .PrimeCommEuro
        End If
        '#######Prévoyance#######
        If .PrimeCommPrev = 0 Then
            .TxChargPrimesPrev = 0
        Else: .TxChargPrimesPrev = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargPrimePrev / _
                                   .PrimeCommPrev
        End If
        '#######################
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesUC = 0 Then
            .TxChargDecesUC = 0
        Else: .TxChargDecesUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargDecesUC / _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC = 0 Then
            .ProbaDecesUC = 0
        Else: .ProbaDecesUC = (1 + .TxChargDecesUC) * _
                              Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesUC / _
                              Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageUC = 0 Then
            .TxChargTirageUC = 0
        Else: .TxChargTirageUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargTirageUC / _
                                 Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC = 0 Then
            .TxTirageUC = 0
        Else: .TxTirageUC = (1 + .TxChargTirageUC) * _
                            Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageUC / _
                            Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotUC = 0 Then
            .TxChargRachatTotUC = 0
        Else: .TxChargRachatTotUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargRachatTotUC / _
                                    Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC = 0 Then
            .ProbaRachatTotUC = 0
        Else: .ProbaRachatTotUC = (1 + .TxChargRachatTotUC) * _
                                  Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotUC / _
                                  Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartUC = 0 Then
            .TxChargRachatPartUC = 0
        Else: .TxChargRachatPartUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargRachatPartUC / _
                                     Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC = 0 Then
            .ProbaRachatPartUC = 0
        Else: .ProbaRachatPartUC = (1 + .TxChargRachatPartUC) * _
                                   Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartUC / _
                                   Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode1UC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeUC = 0 Then
            .TxChargContratsTermesUC = 0
        Else: .TxChargContratsTermesUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargTermeUC / _
                                         Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode2UC = 0 Then
            .ProbaContratsTermesUC = 0
        Else: .ProbaContratsTermesUC = (1 + .TxChargContratsTermesUC) * _
                                       Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeUC / _
                                       Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode2UC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureUC = 0 Then
            .TxChargPM_UC = 0
        Else: .TxChargPM_UC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargPM_UC / _
                              Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureUC = 0 Then
            .TxRetroGlobalPM_UC = 0
        Else: .TxRetroGlobalPM_UC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_RetroGlobalPM_UC / _
                              Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureUC = 0 Then
            .TxRetroAEPM_UC = 0
        Else: .TxRetroAEPM_UC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_RetroAEPM_UC / _
                              Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_ClotureUC
        End If
        If .PM_OuvertureUC = 0 Then
            .TxTechDebutPeriodeUC = 0
        Else: .TxTechDebutPeriodeUC = (1 + Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_InteretsPM_MiPeriodeUC / _
                                           .PM_OuvertureUC) ^ 2 - 1
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode4UC = 0 Then
            .TxTechFinPeriodeUC = 0
        Else: .TxTechFinPeriodeUC = (1 + Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_InteretsFinPeriodeUC / _
                                         Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode4UC) ^ 2 - 1
        End If
        '######Ancienne formule fausse######
        If CompteurAnnee = 1 Then
            .EvolutionPrimesUC = 0
        Else
            If RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).PrimeCommUC = 0 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaDecesUC = 1 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).TxTirageUC = 1 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaRachatTotUC = 1 Or _
               RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaContratsTermesUC = 1 Then
                    .EvolutionPrimesUC = 0
            Else: .EvolutionPrimesUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PrimeCommUC / _
                                       (RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).PrimeCommUC * _
                                        (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaDecesUC) * _
                                        (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).TxTirageUC) * _
                                        (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaRachatTotUC) * _
                                        (1 - RDF(NumModelPoint, CompteurAnnee - 1, NumChoc).ProbaContratsTermesUC))
            End If
        End If
        '##################################
        If .PrimeCommUC = 0 Then
            .TxChargPrimesUC = 0
        Else: .TxChargPrimesUC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargPrimeUC / _
                                 .PrimeCommUC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_Euro_UC = 0 Then
            .TxChargPass_Euro_UC = 0
        Else: .TxChargPass_Euro_UC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargTransf_Euro_UC / _
                                     Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_Euro_UC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode3Euro = 0 Then
            .TxPass_Euro_UC = 0
        Else: .TxPass_Euro_UC = (1 + .TxChargPass_Euro_UC) * _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_Euro_UC / _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode3Euro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_UC_Euro = 0 Then
            .TxChargPass_UC_Euro = 0
        Else: .TxChargPass_UC_Euro = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargTransf_UC_Euro / _
                                     Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_UC_Euro
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode3UC = 0 Then
            .TxPass_UC_Euro = 0
        Else: .TxPass_UC_Euro = (1 + .TxChargPass_UC_Euro) * _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_UC_Euro / _
                                Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode3UC
        End If
        If Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode3UC = 0 Then
            .ChargProbabilise_UC_UC = 0
        Else: .ChargProbabilise_UC_UC = Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_ChargTransf_UC_UC / _
                                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_PM_MiPeriode3UC
        End If
    End With
Next NumModelPoint

End Sub

'*******************************************************************************
'*********************************************************Afin d'obtenir la PM prévoyance de l'année 0 pour chaque MP
Sub CalculsPMprev0()
'*********************************************************
'*******************************************************************************
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    PM_Prev0(BDD(NumLgn).ModelPoint) = PM_Prev0(BDD(NumLgn).ModelPoint) + BDD(NumLgn).PM_Prev(0)
Next NumLgn
PM_Prev0(TransformModelPoint("AINV Option prévoyance")) = PM_Prev0(TransformModelPoint("ESCA Epar"))
PM_Prev0(TransformModelPoint("ESCA Epar")) = 0
End Sub

'*******************************************************************************
Sub CalculsPourAffichageAnnée0()
'*******************************************************************************
Dim CoefPeriodicite As Double, NumLgn As Double


For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If .AnneeEffet = AnneeValorisation Then
            If .Periodicite <> "Unique" Then
                If .Periodicite = "Annuel" Then
                    CoefPeriodicite = 1
                ElseIf .Periodicite = "Semestriel" Then
                    If Int(Month(DateValorisation) - Month(.DateEffet) + 1) / 6 >= 1 Then
                        CoefPeriodicite = 1
                    Else
                        CoefPeriodicite = 1 / 2
                    End If
                ElseIf .Periodicite = "Trimestriel" Then
                    If Int(Month(DateValorisation) - Month(.DateEffet) + 1) / 3 > 3 Then
                        CoefPeriodicite = 1
                    ElseIf Int(Month(DateValorisation) - Month(.DateEffet) + 1) / 3 > 2 Then
                        CoefPeriodicite = 3 / 4
                    ElseIf Int(Month(DateValorisation) - Month(.DateEffet) + 1) / 3 > 1 Then
                        CoefPeriodicite = 1 / 2
                    Else
                        CoefPeriodicite = 1 / 4
                    End If
                ElseIf .Periodicite = "Mensuel" Then
                    CoefPeriodicite = (Month(DateValorisation) - Month(.DateEffet) + 1) / 12
                End If
                .PrimeCommEuroPourAffichageAnnee0 = CoefPeriodicite * .PrimeCommEuro(0)
                .ChargPrimeEuro(0) = .PrimeCommEuroPourAffichageAnnee0 * .Taux_Chargement_Prime_Euro
                .Commissions_PrimesEuro(0) = .PrimeCommEuroPourAffichageAnnee0 * .Taux_Com_Euro
                .PrimeCommUCPourAffichageAnnee0 = CoefPeriodicite * .PrimeCommUC(0)
                .ChargPrimeUC(0) = .PrimeCommUCPourAffichageAnnee0 * .Taux_Chargement_Prime_UC
                .Commissions_PrimesUC(0) = .PrimeCommUCPourAffichageAnnee0 * .Taux_Com_UC
            ElseIf .Periodicite = "Unique" Then
                .PrimeCommEuroPourAffichageAnnee0 = .PM_Euro(0) / (1 - .Taux_Chargement_Prime_Euro)
                .ChargPrimeEuro(0) = .PrimeCommEuroPourAffichageAnnee0 * .Taux_Chargement_Prime_Euro
                .Commissions_PrimesEuro(0) = .PrimeCommEuroPourAffichageAnnee0 * .Taux_Com_Euro
                .PrimeCommUCPourAffichageAnnee0 = .PM_UC(0) / (1 - .Taux_Chargement_Prime_UC)
                .ChargPrimeUC(0) = .PrimeCommUCPourAffichageAnnee0 * .Taux_Chargement_Prime_UC
                .Commissions_PrimesUC(0) = .PrimeCommUCPourAffichageAnnee0 * .Taux_Com_UC
            End If
        End If
    End With
Next NumLgn

End Sub


'*******************************************************************************
 Sub ComptNbTMG() ' Compte et stocke les TMG différents en année de proj N+1
'*******************************************************************************
'L'année N+1 car les TMG de la base n'intègrent pas directemment les paramétrages de taux techniques pour les années
' d'effet > ou < 2011, ces paramétres peuvent différer du TMG initial, de sorte que le TMG change en N+1.
'Par la suite le TMG est supposé constant (=TMG(N+1)) pour le reste de la projection

Dim NumLgn As Long
Dim compteur As Long
Dim different As Boolean

NbTMG = 1
ReDim TMG_Tri(1 To NbTMG)
TMG_Tri(NbTMG) = BDD(1).TMG(1)   'Stocke le 1er TMG calculé en N+1
    
For NumLgn = 2 To NbContrats
    With BDD(NumLgn)
        If .TMG(1) <> BDD(NumLgn - 1).TMG(1) Then
           different = True
           For compteur = 1 To NbTMG
               If TMG_Tri(compteur) = .TMG(1) Then
                  different = False
               End If
           Next compteur
           If different = True Then
              NbTMG = NbTMG + 1
              ReDim Preserve TMG_Tri(1 To NbTMG)
              TMG_Tri(NbTMG) = .TMG(1)
           End If
        End If
    End With
Next NumLgn

End Sub

'*******************************************************************************
 Sub StockFlexingModeling() ' Compte et stocke les TMG différents en année de proj N+1
'*******************************************************************************
'Pour le cas central, stockage des flexing pour le nouvel Output
Dim NumLgn As Long, NumAnnee As Integer, NumPB As Integer, NumTMG As Integer

Call ComptNbTMG

'Distinguer UC et Euro pour les LoB, pour l'UC la PB et le TMG sont constants et nuls, ils ne constituent pas une dimension supplémentaire
ReDim FlexingEuro(0 To Horizon, 1 To NbNiveauxPB, 1 To NbTMG)
ReDim FlexingUC(0 To Horizon)

'*******************************************************************************
'Initialisation (année 0)
'*******************************************************************************
NumPB = 1
For NumTMG = 1 To NbTMG
    With FlexingEuro(0, NumPB, NumTMG)
        .LoB = 30
        .RachDyn = 1
        .TMG = TMG_Tri(NumTMG)
        .PB = 1 'Temporaire
        .Code_support = "Euro"
        For NumLgn = 1 To NbContrats
            If Perimetre = "NB" And BDD(NumLgn).AnneeEffet >= AnneeValorisation Or Perimetre = "Stock" And BDD(NumLgn).AnneeEffet < AnneeValorisation Or Perimetre = "Stock+NB" Then
                If TypePrime = "PP+PU" Or TypePrime = "PP" And BDD(NumLgn).Periodicite <> "Unique" Or TypePrime = "PU" And BDD(NumLgn).Periodicite = "Unique" Then
                    If Obsèque_Epargne = "Total" Or Obsèque_Epargne = "Obsèque" And BDD(NumLgn).IndicObsèque = True Or Obsèque_Epargne = "Epargne" And BDD(NumLgn).IndicObsèque = False Then
                        If TypeTauxMin = "Total" Or TypeTauxMin = "3" And BDD(NumLgn).IndicTypeTauxMin = "3" Or TypeTauxMin = "Autres" And BDD(NumLgn).IndicTypeTauxMin <> "3" Then
                            If ProjParTMG = 0 And Donnees(NumLgn).TMG = 0 Or ProjParTMG = 3.5 / 100 And Donnees(NumLgn).TMG = 3.5 / 100 Or ProjParTMG = 4.5 / 100 And Donnees(NumLgn).TMG = 4.5 / 100 Or ProjParTMG = -1 Then
                                If BDD(NumLgn).TMG(1) = TMG_Tri(NumTMG) Then
                                'La sélection est effectuées sur le TMG(1) et pas l'initial, car l'initial est modifé avec d'autres paramètres par la suite
                                    .PM = .PM + BDD(NumLgn).PM_Euro(0)
                                    .Cotisations = .Cotisations + BDD(NumLgn).PrimeCommEuro(0)
                                    .Effectifs = .Effectifs + BDD(NumLgn).NbClotures_Euro(0)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next NumLgn
    End With
Next NumTMG
With FlexingUC(0)
    .LoB = 31
    .RachDyn = 1
    .TMG = 0
    .PB = 0
    .Code_support = "UC1"
    For NumLgn = 1 To NbContrats
        If Perimetre = "NB" And BDD(NumLgn).AnneeEffet >= AnneeValorisation Or Perimetre = "Stock" And BDD(NumLgn).AnneeEffet < AnneeValorisation Or Perimetre = "Stock+NB" Then
            If TypePrime = "PP+PU" Or TypePrime = "PP" And BDD(NumLgn).Periodicite <> "Unique" Or TypePrime = "PU" And BDD(NumLgn).Periodicite = "Unique" Then
                If Obsèque_Epargne = "Total" Or Obsèque_Epargne = "Obsèque" And BDD(NumLgn).IndicObsèque = True Or Obsèque_Epargne = "Epargne" And BDD(NumLgn).IndicObsèque = False Then
                    If TypeTauxMin = "Total" Or TypeTauxMin = "3" And BDD(NumLgn).IndicTypeTauxMin = "3" Or TypeTauxMin = "Autres" And BDD(NumLgn).IndicTypeTauxMin <> "3" Then
                        If ProjParTMG = 0 And Donnees(NumLgn).TMG = 0 Or ProjParTMG = 3.5 / 100 And Donnees(NumLgn).TMG = 3.5 / 100 Or ProjParTMG = 4.5 / 100 And Donnees(NumLgn).TMG = 4.5 / 100 Or ProjParTMG = -1 Then
                            .PM = .PM + BDD(NumLgn).PM_UC(0)
                            .Cotisations = .Cotisations + BDD(NumLgn).PrimeCommUC(0)
                            .Effectifs = .Effectifs + BDD(NumLgn).NbClotures_UC(0)
                        End If
                    End If
                End If
            End If
        End If
    Next NumLgn
End With
    
'*******************************************************************************
'Pour la projection de N+1 à N+Horizon
'*******************************************************************************

For NumAnnee = 1 To Horizon
'Pour l'instant les niveaux de PB ne sont pas définis et paramétrés dans la base, on utilise 0 et 1 pour avec ou sans PB
    NumPB = 1
    For NumTMG = 1 To NbTMG
        With FlexingEuro(NumAnnee, NumPB, NumTMG)
            .LoB = 30
            .RachDyn = 1
            .TMG = TMG_Tri(NumTMG)
            .PB = 1 'Temporaire
            .Code_support = "Euro"
            For NumLgn = 1 To NbContrats
                If Perimetre = "NB" And BDD(NumLgn).AnneeEffet >= AnneeValorisation Or Perimetre = "Stock" And BDD(NumLgn).AnneeEffet < AnneeValorisation Or Perimetre = "Stock+NB" Then
                    If TypePrime = "PP+PU" Or TypePrime = "PP" And BDD(NumLgn).Periodicite <> "Unique" Or TypePrime = "PU" And BDD(NumLgn).Periodicite = "Unique" Then
                        If Obsèque_Epargne = "Total" Or Obsèque_Epargne = "Obsèque" And BDD(NumLgn).IndicObsèque = True Or Obsèque_Epargne = "Epargne" And BDD(NumLgn).IndicObsèque = False Then
                            If TypeTauxMin = "Total" Or TypeTauxMin = "3" And BDD(NumLgn).IndicTypeTauxMin = "3" Or TypeTauxMin = "Autres" And BDD(NumLgn).IndicTypeTauxMin <> "3" Then
                                If ProjParTMG = 0 And Donnees(NumLgn).TMG = 0 Or ProjParTMG = 3.5 / 100 And Donnees(NumLgn).TMG = 3.5 / 100 Or ProjParTMG = 4.5 / 100 And Donnees(NumLgn).TMG = 4.5 / 100 Or ProjParTMG = -1 Then
                                    If BDD(NumLgn).TMG(1) = TMG_Tri(NumTMG) Then
                                        'Le modèle Addactis considère que les prestations sont brutes de chgmts, il faut donc les rajouter ici
                                        .Presta_Deces = .Presta_Deces + BDD(NumLgn).SinDecesEuro(NumAnnee) + BDD(NumLgn).ChargDecesEuro(NumAnnee)
                                        .Presta_Rachat = .Presta_Rachat + BDD(NumLgn).SinRachatTotEuro(NumAnnee) + BDD(NumLgn).ChargRachatTotEuro(NumAnnee)
                                        .Presta_Terme = .Presta_Terme + BDD(NumLgn).SinTermeEuro(NumAnnee) + BDD(NumLgn).ChargTermeEuro(NumAnnee)
                                        .Presta_Tirage = .Presta_Tirage + BDD(NumLgn).SinTirageEuro(NumAnnee) + BDD(NumLgn).ChargTirageEuro(NumAnnee)
                                        .Presta_Autres = .Presta_Terme + .Presta_Tirage
                                        .PM = .PM + BDD(NumLgn).PM_Euro(NumAnnee)
                                        .Cotisations = .Cotisations + BDD(NumLgn).PrimeCommEuro(NumAnnee)
                                        .Bonus = .Bonus + BDD(NumLgn).BonusRetEuro(NumAnnee)
                                        .IT = .IT + BDD(NumLgn).InteretsPM_MiPeriodeEuro(NumAnnee) + BDD(NumLgn).InteretsFinPeriodeEuro(NumAnnee)
                                        .PS = 0
                                        .Chgt_Cotis = .Chgt_Cotis + BDD(NumLgn).ChargPrimeEuro(NumAnnee)
                                        .Chgt_Encours = .Chgt_Encours + BDD(NumLgn).ChargPM_Euro(NumAnnee) - BDD(NumLgn).BonusRetEuro(NumAnnee)
                                        .Chgt_Presta = .Chgt_Presta + BDD(NumLgn).ChargDecesEuro(NumAnnee) + BDD(NumLgn).ChargRachatTotEuro(NumAnnee) + BDD(NumLgn).ChargTermeEuro(NumAnnee) + BDD(NumLgn).ChargTirageEuro(NumAnnee)
                                        .Chgt_Presta_deces = .Chgt_Presta_deces + BDD(NumLgn).ChargDecesEuro(NumAnnee)
                                        .Chgt_Presta_rachat = .Chgt_Presta_rachat + BDD(NumLgn).ChargRachatTotEuro(NumAnnee)
                                        .Chgt_Presta_terme = .Chgt_Presta_terme + BDD(NumLgn).ChargTermeEuro(NumAnnee)
                                        .Chgt_Presta_tirage = .Chgt_Presta_tirage + BDD(NumLgn).ChargTirageEuro(NumAnnee)
                                        .Frais_Cotis = 0
                                        .Frais_Encours = 0
                                        .Frais_Presta = 0
                                        .Indemnites_Cotis = .Indemnites_Cotis + BDD(NumLgn).Commissions_PrimesEuro(NumAnnee)
                                        .Indemnites_Presta = 0
                                        If BDD(NumLgn).ChargPM_Euro(NumAnnee) = 0 And BDD(NumLgn).Commissions_PMeuro(NumAnnee) <> 0 Then
                                            .Comm_Autres = .Comm_Autres + BDD(NumLgn).Commissions_PMeuro(NumAnnee) ' Les contrats pour lesquels une comm est payée est AE directement --> taux de chgt=0 mais taux comm >0
                                            .Indemnites_Encours = .Indemnites_Encours + 0
                                        Else
                                            .Comm_Autres = .Comm_Autres + 0
                                            .Indemnites_Encours = .Indemnites_Encours + BDD(NumLgn).Commissions_PMeuro(NumAnnee)
                                        End If
                                        .Effectifs = .Effectifs + BDD(NumLgn).NbClotures_Euro(NumAnnee)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next NumLgn
        End With
    Next NumTMG
    With FlexingUC(NumAnnee)
        .LoB = 31
        .RachDyn = 1
        .TMG = 0
        .PB = 0
        .Code_support = "UC1"
        For NumLgn = 1 To NbContrats
            If Perimetre = "NB" And BDD(NumLgn).AnneeEffet >= AnneeValorisation Or Perimetre = "Stock" And BDD(NumLgn).AnneeEffet < AnneeValorisation Or Perimetre = "Stock+NB" Then
                If TypePrime = "PP+PU" Or TypePrime = "PP" And BDD(NumLgn).Periodicite <> "Unique" Or TypePrime = "PU" And BDD(NumLgn).Periodicite = "Unique" Then
                    If Obsèque_Epargne = "Total" Or Obsèque_Epargne = "Obsèque" And BDD(NumLgn).IndicObsèque = True Or Obsèque_Epargne = "Epargne" And BDD(NumLgn).IndicObsèque = False Then
                        If TypeTauxMin = "Total" Or TypeTauxMin = "3" And BDD(NumLgn).IndicTypeTauxMin = "3" Or TypeTauxMin = "Autres" And BDD(NumLgn).IndicTypeTauxMin <> "3" Then
                            If ProjParTMG = 0 And Donnees(NumLgn).TMG = 0 Or ProjParTMG = 3.5 / 100 And Donnees(NumLgn).TMG = 3.5 / 100 Or ProjParTMG = 4.5 / 100 And Donnees(NumLgn).TMG = 4.5 / 100 Or ProjParTMG = -1 Then
                                .Presta_Deces = .Presta_Deces + BDD(NumLgn).SinDecesUC(NumAnnee) + BDD(NumLgn).ChargDecesUC(NumAnnee)
                                .Presta_Rachat = .Presta_Rachat + BDD(NumLgn).SinRachatTotUC(NumAnnee) + BDD(NumLgn).ChargRachatTotUC(NumAnnee)
                                .Presta_Terme = .Presta_Terme + BDD(NumLgn).SinTermeUC(NumAnnee) + BDD(NumLgn).ChargTermeUC(NumAnnee)
                                .Presta_Tirage = .Presta_Tirage + BDD(NumLgn).SinTirageUC(NumAnnee) + BDD(NumLgn).ChargTirageUC(NumAnnee)
                                .Presta_Autres = .Presta_Terme + .Presta_Tirage
                                .PM = .PM + BDD(NumLgn).PM_UC(NumAnnee)
                                .Cotisations = .Cotisations + BDD(NumLgn).PrimeCommUC(NumAnnee)
                                .Bonus = .Bonus + BDD(NumLgn).BonusRetUC(NumAnnee)
                                .IT = .IT + BDD(NumLgn).InteretsFinPeriodeUC(NumAnnee) + BDD(NumLgn).InteretsPM_MiPeriodeUC(NumAnnee)
                                .PS = 0
                                .Chgt_Cotis = .Chgt_Cotis + BDD(NumLgn).ChargPrimeUC(NumAnnee)
                                .Chgt_Encours = .Chgt_Encours + BDD(NumLgn).ChargPM_UC(NumAnnee) + BDD(NumLgn).RetroGlobalPM_UC(NumAnnee) - BDD(NumLgn).RetroAEPM_UC(NumAnnee) - BDD(NumLgn).BonusRetUC(NumAnnee) ' Les retro globales hors AE dans les chgt pour boucler avec déroulé
                                .Chgt_Presta = .Chgt_Presta + BDD(NumLgn).ChargDecesUC(NumAnnee) + BDD(NumLgn).ChargRachatTotUC(NumAnnee) + BDD(NumLgn).ChargTermeUC(NumAnnee) + BDD(NumLgn).ChargTirageUC(NumAnnee)
                                .Chgt_Presta_deces = .Chgt_Presta_deces + BDD(NumLgn).ChargDecesUC(NumAnnee)
                                .Chgt_Presta_rachat = .Chgt_Presta_rachat + BDD(NumLgn).ChargRachatTotUC(NumAnnee)
                                .Chgt_Presta_terme = .Chgt_Presta_terme + BDD(NumLgn).ChargTermeUC(NumAnnee)
                                .Chgt_Presta_tirage = .Chgt_Presta_tirage + BDD(NumLgn).ChargTirageUC(NumAnnee)
                                .Frais_Cotis = 0
                                .Frais_Encours = 0
                                .Frais_Presta = 0
                                .Indemnites_Cotis = .Indemnites_Cotis + BDD(NumLgn).Commissions_PrimesUC(NumAnnee)
                                .Indemnites_Presta = 0
                                .Retro = .Retro + BDD(NumLgn).RetroGlobalPM_UC(NumAnnee)
                                .RetroAE = .RetroAE + BDD(NumLgn).RetroAEPM_UC(NumAnnee) ' dans la colonne retro seulement la part AE
                                .Effectifs = .Effectifs + BDD(NumLgn).NbClotures_UC(NumAnnee)
                                If BDD(NumLgn).ChargPM_UC(NumAnnee) = 0 And BDD(NumLgn).Commissions_PMuc(NumAnnee) <> 0 Then
                                    .Comm_Autres = .Comm_Autres + BDD(NumLgn).Commissions_PMuc(NumAnnee) ' Les contrats pour lesquels une comm est payée est AE directement --> taux de chgt=0 mais taux comm >0
                                    .Indemnites_Encours = .Indemnites_Encours + BDD(NumLgn).RetroGlobalPM_UC(NumAnnee) - BDD(NumLgn).RetroAEPM_UC(NumAnnee) ' Les retro hors AE  sont à traiter comme les comm, les retro aef devront être déduites ensuite
                                Else
                                    .Comm_Autres = .Comm_Autres + 0
                                    .Indemnites_Encours = .Indemnites_Encours + BDD(NumLgn).Commissions_PMuc(NumAnnee) + BDD(NumLgn).RetroGlobalPM_UC(NumAnnee) - BDD(NumLgn).RetroAEPM_UC(NumAnnee)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next NumLgn
    End With
Next NumAnnee
End Sub

Sub CalculsCoutUnitaireGestion()

Dim NumLgn As Double

'Set FeuilAct = ThisWorkbook.Worksheets("Test")

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If CompteurAnnee = 1 Then
            If TypeCU <> "Global" Then
                .CoutUnitaireGestion_Euro(0) = CU_adm_Euro
                .CoutUnitaireGestion_UC(0) = CU_adm_UC
                .CoutUnitairePresta_Euro(0) = CU_Prest_Euro
                .CoutUnitairePresta_UC(0) = CU_prest_UC
            Else
                .CoutUnitaireGestion_Euro(0) = CU_Adm
                .CoutUnitaireGestion_UC(0) = CU_Adm
                .CoutUnitairePresta_Euro(0) = CU_Prest
                .CoutUnitairePresta_UC(0) = CU_Prest
            End If
        Else
            ' impact le coût de départ de l'inflation
            .CoutUnitaireGestion_Euro(0) = .CoutUnitaireGestion_Euro(0) * (1 + Inflation(CompteurAnnee))
            .CoutUnitaireGestion_UC(0) = .CoutUnitaireGestion_UC(0) * (1 + Inflation(CompteurAnnee))
            .CoutUnitairePresta_Euro(0) = .CoutUnitairePresta_Euro(0) * (1 + Inflation(CompteurAnnee))
            .CoutUnitairePresta_UC(0) = .CoutUnitairePresta_UC(0) * (1 + Inflation(CompteurAnnee))
        End If
        If .PM_Euro(CompteurAnnee) + .PM_UC(CompteurAnnee) = 0 Then
            .CoutUnitaireGestion_Euro(CompteurAnnee) = 0
            .CoutUnitaireGestion_UC(CompteurAnnee) = 0
            .CoutUnitairePresta_Euro(CompteurAnnee) = 0
            .CoutUnitairePresta_UC(CompteurAnnee) = 0
        Else
            .CoutUnitaireGestion_Euro(CompteurAnnee) = .CoutUnitaireGestion_Euro(0) * .NbClotures_Euro(CompteurAnnee) * .NbTetesCU
            .CoutUnitaireGestion_UC(CompteurAnnee) = .CoutUnitaireGestion_UC(0) * .NbClotures_UC(CompteurAnnee) * .NbTetesCU
            .CoutUnitairePresta_Euro(CompteurAnnee) = .CoutUnitairePresta_Euro(0) * .NbClotures_Euro(CompteurAnnee) * .NbTetesCU
            .CoutUnitairePresta_UC(CompteurAnnee) = .CoutUnitairePresta_UC(0) * .NbClotures_UC(CompteurAnnee) * .NbTetesCU
            If TypeCU = "Global" Then
            'Dans le cas où le coût unitaire est global on le réparti au prorata de la PM entre euro et UC
            .CoutUnitaireGestion_Euro(CompteurAnnee) = .CoutUnitaireGestion_Euro(CompteurAnnee) * ((.PM_Euro(CompteurAnnee - 1) + .PM_Euro(CompteurAnnee)) / 2) / ((.PM_Euro(CompteurAnnee - 1) + _
                                                    .PM_Euro(CompteurAnnee) + .PM_UC(CompteurAnnee - 1) + .PM_UC(CompteurAnnee)) / 2)
            .CoutUnitaireGestion_UC(CompteurAnnee) = .CoutUnitaireGestion_UC(CompteurAnnee) * ((.PM_UC(CompteurAnnee - 1) + .PM_UC(CompteurAnnee)) / 2) / ((.PM_Euro(CompteurAnnee - 1) + _
                                                    .PM_Euro(CompteurAnnee) + .PM_UC(CompteurAnnee - 1) + .PM_UC(CompteurAnnee)) / 2)
            .CoutUnitairePresta_Euro(CompteurAnnee) = .CoutUnitairePresta_Euro(CompteurAnnee) * ((.PM_Euro(CompteurAnnee - 1) + .PM_Euro(CompteurAnnee)) / 2) / ((.PM_Euro(CompteurAnnee - 1) + _
                                                    .PM_Euro(CompteurAnnee) + .PM_UC(CompteurAnnee - 1) + .PM_UC(CompteurAnnee)) / 2)
            .CoutUnitairePresta_UC(CompteurAnnee) = .CoutUnitairePresta_UC(CompteurAnnee) * ((.PM_UC(CompteurAnnee - 1) + .PM_UC(CompteurAnnee)) / 2) / ((.PM_Euro(CompteurAnnee - 1) + _
                                                    .PM_Euro(CompteurAnnee) + .PM_UC(CompteurAnnee - 1) + .PM_UC(CompteurAnnee)) / 2)
        End If
        End If
'        If CompteurAnnee = 1 Then
'            FeuilAct.Cells(1 + NumLgn, 2) = .CoutUnitaireGestion_Euro(CompteurAnnee)
'        End If
    End With
Next NumLgn

End Sub
