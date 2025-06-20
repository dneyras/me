Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongLong
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongLong

Public Sub ExportResultats(Scenario As String, DossierOutPut As String, FichierOutPut As String)

Dim NumModelPoint As Integer, CompteurAnnee As Integer
Dim decalage As Integer

If Not ExisteDossier(DossierOutPut) Then
    MsgBox "Vous avez entré un nom de dossier inexistant!", vbCritical + vbOKOnly, "Erreur"
    GoTo FinSub
End If

If ExisteFichier(DossierOutPut & "/" & FichierOutPut) Then
    If Not EstOuvert(FichierOutPut) Then
        Workbooks.Open Filename:=DossierOutPut & "/" & FichierOutPut, ReadOnly:=True
    End If
    If Not ExisteFeuille(FichierOutPut, "RATIOS DE FLEXING - " & Scenario) Then
        Workbooks(FichierOutPut).Activate
        Call creer("RATIOS DE FLEXING - " & Scenario)
    End If
Else: Workbooks.Add
      ActiveWorkbook.SaveAs Filename:=DossierOutPut & "/" & FichierOutPut, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
      Call creer("RATIOS DE FLEXING - " & Scenario)
      On Error GoTo Echec
      Application.DisplayAlerts = False
      Workbooks(FichierOutPut).Worksheets("Feuil1").Delete
'''      Workbooks(FichierOutPut).Worksheets("Feuil2").Delete
'''      Workbooks(FichierOutPut).Worksheets("Feuil3").Delete
      Application.DisplayAlerts = True
End If

Echec:

ThisWorkbook.Sheets("RATIOS DE FLEXING").Activate
Cells.Select
Selection.Copy
Workbooks(FichierOutPut).Sheets("RATIOS DE FLEXING - " & Scenario).Activate
Cells.Select
ActiveSheet.Paste
Call VidePressePapier
With Workbooks(FichierOutPut).Worksheets("RATIOS DE FLEXING - " & Scenario)
    For NumModelPoint = 1 To NbModelPoint
        For decalage = 5 To 4820 Step 72
            .Range("A" & NumModelPoint + decalage).Value = InverseModelPoint(NumModelPoint)
        Next decalage
        For CompteurAnnee = 1 To Horizon
            .Cells(NumModelPoint + 5, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargDecesEuro
            .Cells(NumModelPoint + 77, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaDecesEuro
            .Cells(NumModelPoint + 149, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargTirageEuro
            .Cells(NumModelPoint + 221, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxTirageEuro
            .Cells(NumModelPoint + 293, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargRachatTotEuro
            .Cells(NumModelPoint + 365, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaRachatTotEuro
            .Cells(NumModelPoint + 437, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargRachatPartEuro
            .Cells(NumModelPoint + 509, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaRachatPartEuro
            .Cells(NumModelPoint + 581, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargContratsTermesEuro
            .Cells(NumModelPoint + 653, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaContratsTermesEuro
            .Cells(NumModelPoint + 725, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargPM_Euro
            .Cells(NumModelPoint + 797, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxTechDebutPeriodeEuro
            .Cells(NumModelPoint + 869, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxTechFinPeriodeEuro
            .Cells(NumModelPoint + 941, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).EvolutionPrimesEuro
            .Cells(NumModelPoint + 1013, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargPrimesEuro
            .Cells(NumModelPoint + 1085, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargDecesUC
            .Cells(NumModelPoint + 1157, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaDecesUC
            .Cells(NumModelPoint + 1229, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargTirageUC
            .Cells(NumModelPoint + 1301, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxTirageUC
            .Cells(NumModelPoint + 1373, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargRachatTotUC
            .Cells(NumModelPoint + 1445, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaRachatTotUC
            .Cells(NumModelPoint + 1517, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargRachatPartUC
            .Cells(NumModelPoint + 1589, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaRachatPartUC
            .Cells(NumModelPoint + 1661, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargContratsTermesUC
            .Cells(NumModelPoint + 1733, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ProbaContratsTermesUC
            .Cells(NumModelPoint + 1805, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargPM_UC
            .Cells(NumModelPoint + 1877, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxTechDebutPeriodeUC
            .Cells(NumModelPoint + 1949, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxTechFinPeriodeUC
            .Cells(NumModelPoint + 2021, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).EvolutionPrimesUC
            .Cells(NumModelPoint + 2093, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargPrimesUC
            .Cells(NumModelPoint + 2165, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargPass_Euro_UC
            .Cells(NumModelPoint + 2237, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxPass_Euro_UC
            .Cells(NumModelPoint + 2309, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargPass_UC_Euro
            .Cells(NumModelPoint + 2381, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).TxPass_UC_Euro
            .Cells(NumModelPoint + 2453, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).ChargProbabilise_UC_UC
            .Cells(NumModelPoint + 2525, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).NbOuvertures
            .Cells(NumModelPoint + 2597, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).PrimeCommEuro
            .Cells(NumModelPoint + 2669, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).PM_OuvertureEuro
            .Cells(NumModelPoint + 2741, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).PrimeCommUC
            .Cells(NumModelPoint + 2813, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).PM_OuvertureUC
            .Cells(NumModelPoint + 2885, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Commissions_PrimesEuro
            .Cells(NumModelPoint + 2957, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Commissions_PMeuro
            .Cells(NumModelPoint + 3029, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Commissions_PrimesUC
            .Cells(NumModelPoint + 3101, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Commissions_PMuc
            .Cells(NumModelPoint + 3173, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesEuro
            .Cells(NumModelPoint + 3245, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageEuro
            .Cells(NumModelPoint + 3317, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotEuro
            .Cells(NumModelPoint + 3389, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartEuro
            .Cells(NumModelPoint + 3461, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinDecesUC
            .Cells(NumModelPoint + 3533, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTirageUC
            .Cells(NumModelPoint + 3605, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatTotUC
            .Cells(NumModelPoint + 3677, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinRachatPartUC
            .Cells(NumModelPoint + 3749, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_Euro_UC
            .Cells(NumModelPoint + 3821, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Cap_UC_Euro
            .Cells(NumModelPoint + 3893, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeEuro
            .Cells(NumModelPoint + 3965, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_SinTermeUC
            .Cells(NumModelPoint + 4037, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_BonusRetEuro
            .Cells(NumModelPoint + 4109, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_BonusRetUC
            .Cells(NumModelPoint + 4181, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_RetroGlobalPM_UC
            .Cells(NumModelPoint + 4253, CompteurAnnee + 2).Value = _
                        RDF(NumModelPoint, CompteurAnnee, NumChoc).TxRetroGlobalPM_UC
            .Cells(NumModelPoint + 4325, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_RetroAEPM_UC
            .Cells(NumModelPoint + 4397, CompteurAnnee + 2).Value = _
                        RDF(NumModelPoint, CompteurAnnee, NumChoc).TxRetroAEPM_UC
            .Cells(NumModelPoint + 4469, CompteurAnnee + 2).Value = _
                        Totaux_Par_MP(NumModelPoint, CompteurAnnee).Somm_Capital_Garanti
            .Cells(NumModelPoint + 4541, CompteurAnnee + 2).Value = _
                        RDF(NumModelPoint, CompteurAnnee, NumChoc).PM_CloturePrev '***
            .Cells(NumModelPoint + 4613, CompteurAnnee + 2).Value = _
                        RDF(NumModelPoint, CompteurAnnee, NumChoc).NbOuverturesPrev '***
            .Cells(NumModelPoint + 4685, CompteurAnnee + 2).Value = _
                        RDF(NumModelPoint, CompteurAnnee, NumChoc).PrimeCommPrev '***
            .Cells(NumModelPoint + 4757, CompteurAnnee + 2).Value = _
                        RDF(NumModelPoint, CompteurAnnee, NumChoc).TxChargPrimesPrev '***
            .Cells(NumModelPoint + 4820, CompteurAnnee + 2).Value = RDF(NumModelPoint, CompteurAnnee, NumChoc).NbOuverturesUC
        Next CompteurAnnee
        .Cells(NumModelPoint + 4541, 2).Value = PM_Prev0(NumModelPoint) '***
    Next NumModelPoint
End With

FinSub:

End Sub
Public Sub ExportResultatsModeling(Scenario As String, DossierOutPut As String, FichierOutPut As String)

Dim NumModelPoint As Integer, CompteurAnnee As Integer, FichierOutputChoc As String, NomChocModeling As String
Dim decalage As Integer

NomChocModeling = TransformNumChocModeling(NumChoc)
FichierOutputChoc = FichierOutPut & "_" & NomChocModeling & ".csv"

If Not ExisteDossier(DossierOutPut) Then
    MsgBox "Vous avez entré un nom de dossier inexistant!", vbCritical + vbOKOnly, "Erreur"
    GoTo FinSub
End If

If ExisteFichier(DossierOutPut & "/" & FichierOutputChoc) Then
    MsgBox "Le fichier ne peut être créé car un fichier du même nom est déjà dans le dossier!", vbCritical + vbOKOnly, "Erreur"
    GoTo FinSub

Else
    ThisWorkbook.Sheets("ADDACTIS " & Scenario).Activate
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Application.DisplayAlerts = False
    Cells.Select
    ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats
    ActiveWorkbook.SaveAs Filename:=DossierOutPut & "/" & FichierOutputChoc, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    Workbooks(FichierOutputChoc).Close
End If



Call VidePressePapier
Application.DisplayAlerts = True

FinSub:

End Sub

Private Sub VidePressePapier()

OpenClipboard 0
EmptyClipboard
CloseClipboard

End Sub





