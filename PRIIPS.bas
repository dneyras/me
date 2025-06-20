'#############################################################################################
'Calcul en mode Central, cependant pour les produits Priips, les flux ne sont pas probabilisés.
'#############################################################################################

'*******************************************************************************
Public Function EstPriips(NumModelPoint As Integer) As Boolean
'*******************************************************************************
Dim CompteurPriips As Integer

EstPriips = False
For CompteurPriips = 1 To NbPriips
    If InverseModelPoint(NumModelPoint) = NomsPriips(CompteurPriips) Then
        EstPriips = True
        Exit Function
    End If
Next CompteurPriips

End Function
'*******************************************************************************
Public Sub MainPriips()
'*******************************************************************************
AgeMP = True
Erase Totaux_Par_MP
NbModelPoint = NbModelPointPRIIPS

Initialisation.LitParametres
Initialisation.AbsenceErreur
Initialisation.ComptNbContrats
Initialisation.LitData
Initialisation.ComptNbTMGprev
Initialisation.LitHypotheses
Initialisation.LitHypothesesMortalite
Initialisation.LitLx
PRIIPS.LitPRIIPS

If ThisWorkbook.Worksheets("PARAMETRES").Range("G22").Value = "Oui" Then
    ScenCent = True
Else
    ScenCent = False
End If


CalculsPreliminaires_parContrat
CalculsRedimQx
CalculsTableCommut
CalculsLancerChoc
NumChoc = 0
    If LancerChoc(NumChoc) = True Then
        CalculsChocs_Qx
        CalculsChocs_TauxRachat
        For CompteurAnnee = 1 To Horizon
'            CalculsNbDeces_NbTirages_NbRachatsTot_NbRachatsPart_NbTermes_NbClotures
            PRIIPS.CalculsNbDeces_NbTirages_NbRachatsTot_NbRachatsPart_NbTermes_NbCloturesPRIIPS
            CalculsBonus
            CalculsCoeffPrimeAnneeEch
            CalculsCoeffPrimeAnneeBonus
            CalculsCoeffPrimePrev
            CalculsTMG
            CalculsPrimeEuro
            CalculsChargPrime_Euro
            CalculsCommissionPrime_Euro
            CalculsPrimeNette_Euro
            CalculsPM_MiPeriode1_Euro
'            CalculsSinistres_Euro
            PRIIPS.CalculsSinistres_EuroPriips
            CalculsChargementsSinistres_Euro
            CalculsSinistreTerme_ChargementTerme_Euro
            CalculsPrime_UC
            CalculsChargPrime_UC
            CalculsCommissionPrime_UC
            CalculsPrimeNette_UC
            CalculsPM_MiPeriode1_UC
'            CalculsSinistres_UC
            PRIIPS.CalculsSinistres_UCPriips
            CalculsChargementsSinistres_UC
            CalculsSinitreTerme_ChargementTerme_UC
            'CalculsTransfertsCapitaux_Chargements_Euro_UC
            PRIIPS.CalculsTransfertsCapitaux_Chargements_Euro_UC_Priips
            CalculsPrime_Prev
            CalculsChargPrime_Prev
            CalculsCommissionPrime_Prev
            CalculsPrimeNette_Prev
'            CalculsSinistres_Prev
            PRIIPS.CalculsSinistres_PrevPriips
            CalculsChargementsSinistres_Prev
            CalculsPMCloture_Euro
            CaculsChargPM_Euro
            CalculsCommissionPM_Euro
            CalculsPM_Euro
            CalculsMarg_Euro
            CalculsCharg_Euro
            CalculsTransfertsCapitaux_Chargements_UC_Euro
            CalculsPMCloture_UC
            CaculsChargPM_UC
            CalculsCommissionPM_UC
            CalculsPM_UC
            CalculsMargUC
            CalculsChargUC
            CaculsChargPM_Prev
            CalculsCommissionPM_Prev
            CalculsPM_Prev
            CalculsMarg_Prev
            CalculsCharg_Prev
            CalculsSommes
            CalculsRatiosFlexing
        Next CompteurAnnee
        CalculsPMprev0
        If FichierSorties = True Then
            Call ExportResultats.ExportResultats("CENTRAL", DossierOutPut, FichierOutPut)
        End If
    End If
    Erase Totaux_Par_MP
    ReDim Totaux_Par_MP(0 To NbModelPoint, 1 To Horizon)
 

Workbooks(FichierOutPut).Save
Workbooks(FichierOutPut).Close

FinSub:

End Sub
'*******************************************************************************
Sub LitPRIIPS()
'*******************************************************************************

Dim CompteurPriips As Integer
 
CompteurPriips = 12
NbPriips = 0
Do While ThisWorkbook.Worksheets("PRIIPS").Range("P" & CompteurPriips).Value <> ""
    NbPriips = NbPriips + 1
    CompteurPriips = CompteurPriips + 1
Loop

ReDim NomsPriips(1 To NbPriips)
For CompteurPriips = 1 To NbPriips
    NomsPriips(CompteurPriips) = ThisWorkbook.Worksheets("PRIIPS").Cells(CompteurPriips + 11, 16)
Next CompteurPriips

End Sub
'*******************************************************************************
Sub CalculsNbDeces_NbTirages_NbRachatsTot_NbRachatsPart_NbTermes_NbCloturesPRIIPS()
'*******************************************************************************
Dim NumLgn As Double
        
For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If EstPriips(.ModelPoint) = True Then
            .NbDeces(CompteurAnnee) = 0
            .NbTirages(CompteurAnnee) = 0
            .NbRachatsTot(CompteurAnnee) = 0
            .NbRachatsPart(CompteurAnnee) = 0
            .NbTermes(CompteurAnnee) = .IndTermeContrat(CompteurAnnee - 1) * (.NbClotures(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - _
                                        .NbRachatsTot(CompteurAnnee))
            .NbClotures(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) - .NbTermes(CompteurAnnee)
            .NbDeces_Prev(CompteurAnnee) = 0
            .NbRachatsTot_Prev(CompteurAnnee) = 0
            .NbTermes_Prev(CompteurAnnee) = IIf(.DureeRestante_Prev(CompteurAnnee - 1) = 1, .NbClotures_Prev(CompteurAnnee - 1) - .NbDeces_Prev(CompteurAnnee) - _
                                        .NbRachatsTot_Prev(CompteurAnnee), 0)
            .NbClotures_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) - .NbTermes_Prev(CompteurAnnee)
            
        Else
            .NbDeces(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Max(0, Min(1, qxChoques(NumLgn, CompteurAnnee, NumChoc)))
            .NbTirages(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(TxTirage(.NomProd, CompteurAnnee), Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc)))
            .NbRachatsTot(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot, _
                                               Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee)))
            .NbRachatsPart(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatPart, _
                                                Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc) - TxTirage(.NomProd, CompteurAnnee) - _
                                                BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot))
            .NbTermes(CompteurAnnee) = .IndTermeContrat(CompteurAnnee - 1) * (.NbClotures(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - _
                                        .NbRachatsTot(CompteurAnnee))
            .NbClotures(CompteurAnnee) = .NbClotures(CompteurAnnee - 1) - .NbDeces(CompteurAnnee) - .NbTirages(CompteurAnnee) - .NbRachatsTot(CompteurAnnee) - _
                                         .NbTermes(CompteurAnnee)
            .NbDeces_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) * Max(0, Min(1, qxChoques(NumLgn, CompteurAnnee, NumChoc))) '**
            .NbRachatsTot_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) * Min(BDR(.CatRachatTot, .AncienneteContrat(CompteurAnnee)).ProbaRachatTot_Prev, _
                                               Max(0, 1 - qxChoques(NumLgn, CompteurAnnee, NumChoc)))                                      '**
            .NbTermes_Prev(CompteurAnnee) = IIf(.DureeRestante_Prev(CompteurAnnee - 1) = 1, .NbClotures_Prev(CompteurAnnee - 1) - .NbDeces_Prev(CompteurAnnee) - _
                                        .NbRachatsTot_Prev(CompteurAnnee), 0)                                                              '**
            .NbClotures_Prev(CompteurAnnee) = .NbClotures_Prev(CompteurAnnee - 1) - .NbDeces_Prev(CompteurAnnee) - .NbRachatsTot_Prev(CompteurAnnee) - _
                                         .NbTermes_Prev(CompteurAnnee) '**
        End If
    End With
Next NumLgn

End Sub
'******************************************************************
Sub CalculsSinistres_EuroPriips()
'******************************************************************
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If EstPriips(.ModelPoint) = True Then
            .SinDecesEuro(CompteurAnnee) = 0
            .SinTirageEuro(CompteurAnnee) = 0
            .SinRachatTotEuro(CompteurAnnee) = 0
            .SinRachatPartEuro(CompteurAnnee) = 0
        Else
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
        End If
    End With
Next NumLgn

End Sub
'******************************************************************
Sub CalculsSinistres_UCPriips()
'******************************************************************
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If EstPriips(.ModelPoint) = True Then
            .SinDecesUC(CompteurAnnee) = 0
            .SinTirageUC(CompteurAnnee) = 0
            .SinRachatTotUC(CompteurAnnee) = 0
            .SinRachatPartUC(CompteurAnnee) = 0
        Else
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
        End If
    End With
Next NumLgn

End Sub
'******************************************************************
Sub CalculsSinistres_PrevPriips()
'******************************************************************
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If EstPriips(.ModelPoint) = True Then
            .SinDecesPrev(CompteurAnnee) = 0
            .SinRachatTotPrev(CompteurAnnee) = 0
        Else
            .SinDecesPrev(CompteurAnnee) = .CapitalDeces(CompteurAnnee) * _
                                           (qxChoques(NumLgn, CompteurAnnee, NumChoc) / (1 + ChDeces(.NomProd, .AncienneteContrat(CompteurAnnee)))) * .NbClotures_Prev(CompteurAnnee - 1)
            .SinRachatTotPrev(CompteurAnnee) = 0
        End If
    End With
Next NumLgn

End Sub
'******************************************************************
Sub CalculsTransfertsCapitaux_Chargements_Euro_UC_Priips()
'******************************************************************
Dim NumLgn As Double

For NumLgn = 1 To NbContrats
    With BDD(NumLgn)
        If EstPriips(.ModelPoint) = True Then
            .Cap_Euro_UC(CompteurAnnee) = 0
            .ChargTransf_Euro_UC(CompteurAnnee) = 0
            .Cap_UC_Euro(CompteurAnnee) = 0
            .PM_MiPeriode4Euro(CompteurAnnee) = .PM_MiPeriode3Euro(CompteurAnnee) + .Cap_UC_Euro(CompteurAnnee) - .Cap_Euro_UC(CompteurAnnee) - _
                                                .ChargTransf_Euro_UC(CompteurAnnee)
        Else
            .Cap_Euro_UC(CompteurAnnee) = .PM_MiPeriode3Euro(CompteurAnnee) * TxPass_Euro_UC(.NomProd, .AncienneteContrat(CompteurAnnee)) / _
                                          (1 + ChArb_Euro_UC(.NomProd, .AncienneteContrat(CompteurAnnee)))
            .ChargTransf_Euro_UC(CompteurAnnee) = .Cap_Euro_UC(CompteurAnnee) * ChArb_Euro_UC(.NomProd, .AncienneteContrat(CompteurAnnee))
            .Cap_UC_Euro(CompteurAnnee) = .PM_MiPeriode3UC(CompteurAnnee) * TxPass_UC_Euro(.NomProd, .AncienneteContrat(CompteurAnnee)) / _
                                          (1 + ChArb_UC_Euro(.NomProd, .AncienneteContrat(CompteurAnnee)))
            
            .PM_MiPeriode4Euro(CompteurAnnee) = .PM_MiPeriode3Euro(CompteurAnnee) + .Cap_UC_Euro(CompteurAnnee) - .Cap_Euro_UC(CompteurAnnee) - _
                                                .ChargTransf_Euro_UC(CompteurAnnee)
        End If
    End With
Next NumLgn

End Sub

