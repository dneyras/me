Public Sub AbsenceErreur()

    BDE.PresenceErreur = False

End Sub

Public Sub ComptNbContrats()

Dim compteur As Long
Dim FeuilBDD As String

FeuilBDD = "MODEL POINT"

NbContrats = 0
compteur = 2

Do While ThisWorkbook.Worksheets(FeuilBDD).Range("A" & compteur).Value <> ""
    NbContrats = NbContrats + 1
    compteur = compteur + 1
Loop

'Compte le nombre de cales à ajouter à la base de données
FeuilBDD = "CALES"

NbCales = 0
compteur = 2

Do While ThisWorkbook.Worksheets(FeuilBDD).Range("A" & compteur).Value <> ""
    NbCales = NbCales + 1
    compteur = compteur + 1
Loop


End Sub
'***************************************************
Public Sub ComptNbTMGprev() ' Compte et stocke les TMG différents pour l'option prévoyance

Dim NumLgn As Long
Dim compteur As Long
Dim different As Boolean

NbTMGprev = 1
ReDim TMGprev(1 To NbTMGprev)
TMGprev(NbTMGprev) = Donnees(1).TMGprev    'Stocke le 1er TMG prévoyance du Model Point

For NumLgn = 2 To NbContrats
    With Donnees(NumLgn)
        If .TMGprev <> Donnees(NumLgn - 1).TMGprev Then
           different = True
           For compteur = 1 To NbTMGprev
               If TMGprev(compteur) = .TMGprev Then
                  different = False
               End If
           Next compteur
           If different = True Then
              NbTMGprev = NbTMGprev + 1
              ReDim Preserve TMGprev(1 To NbTMGprev)
              TMGprev(NbTMGprev) = .TMGprev
           End If
        End If
    End With
Next NumLgn

End Sub


'*******************************************************************************
Sub LitParametres()
'*******************************************************************************
Dim NumChoc As Integer

DateValorisation = ThisWorkbook.Worksheets("PARAMETRES").Range("G7").Value
AnneeValorisation = Year(DateValorisation)
AnneeInitiale = AnneeValorisation

Perimetre = ThisWorkbook.Worksheets("PARAMETRES").Range("G13").Value
TypePrime = ThisWorkbook.Worksheets("PARAMETRES").Range("G15").Value
Obsèque_Epargne = ThisWorkbook.Worksheets("PARAMETRES").Range("G17").Value
TypeTauxMin = ThisWorkbook.Worksheets("PARAMETRES").Range("G19").Value
PrimeOuSans = ThisWorkbook.Worksheets("PARAMETRES").Range("G21").Value
ProjParTMG = IIf(ThisWorkbook.Worksheets("PARAMETRES").Range("G23").Value = "Tout", -1, ThisWorkbook.Worksheets("PARAMETRES").Range("G23").Value)

ChocMortalite = ThisWorkbook.Worksheets("PARAMETRES").Range("I30").Value
ChocLongevite = ThisWorkbook.Worksheets("PARAMETRES").Range("I32").Value
ChocHausse = ThisWorkbook.Worksheets("PARAMETRES").Range("I34").Value
ChocBaisse = ThisWorkbook.Worksheets("PARAMETRES").Range("I36").Value
ChocMassif = ThisWorkbook.Worksheets("PARAMETRES").Range("I38").Value
ChocCatastrophe = ThisWorkbook.Worksheets("PARAMETRES").Range("I40").Value
ChocFrais = ThisWorkbook.Worksheets("PARAMETRES").Range("I42").Value

If ThisWorkbook.Worksheets("PARAMETRES").Range("G28").Value = "Oui" Then
    ScenCent = True
Else
    ScenCent = False
End If

If ThisWorkbook.Worksheets("PARAMETRES").Range("G30").Value = "Oui" Then
    ScenMort = True
Else
    ScenMort = False
End If

If ThisWorkbook.Worksheets("PARAMETRES").Range("G32").Value = "Oui" Then
    ScenLong = True
Else
    ScenLong = False
End If
            
If ThisWorkbook.Worksheets("PARAMETRES").Range("G34").Value = "Oui" Then
    ScenHausse = True
Else
    ScenHausse = False
End If
            
If ThisWorkbook.Worksheets("PARAMETRES").Range("G36").Value = "Oui" Then
    ScenBaisse = True
Else
    ScenBaisse = False
End If

If ThisWorkbook.Worksheets("PARAMETRES").Range("G38").Value = "Oui" Then
    ScenMassif = True
Else
    ScenMassif = False
End If

If ThisWorkbook.Worksheets("PARAMETRES").Range("G40").Value = "Oui" Then
    ScenCata = True
Else
    ScenCata = False
End If

If ThisWorkbook.Worksheets("PARAMETRES").Range("G42").Value = "Oui" Then
    ScenFrais = True
Else
    ScenFrais = False
End If



If ThisWorkbook.Worksheets("PARAMETRES").Range("G49").Value = "Oui" Then
    FichierSorties = True
Else
    FichierSorties = False
End If

DossierOutPut = ThisWorkbook.Worksheets("PARAMETRES").Range("E53").Value
FichierOutPut = ThisWorkbook.Worksheets("PARAMETRES").Range("E57").Value

If ThisWorkbook.Worksheets("PARAMETRES").Range("G60").Value = "Oui" Then
    FichierSortiesModeling = True
Else
    FichierSortiesModeling = False
End If

DossierOutPutModeling = ThisWorkbook.Worksheets("PARAMETRES").Range("E64").Value
FichierOutPutModeling = ThisWorkbook.Worksheets("PARAMETRES").Range("E68").Value

 ChocAddactis(0) = True
For NumChoc = 1 To 6
    ChocAddactis(NumChoc) = IIf(ThisWorkbook.Worksheets("PARAMETRES").Cells(70 + NumChoc * 2, 7) = "Choc", True, False)
Next NumChoc

'Paramétrage des imports
With FichierCbTaux
    .Dossier = ThisWorkbook.Worksheets("PARAMETRES").Range("E90").Value
    .Fichier = ThisWorkbook.Worksheets("PARAMETRES").Range("E92").Value
    .Onglet = ThisWorkbook.Worksheets("PARAMETRES").Range("E94").Value
End With

With FichierInflation
    .Dossier = ThisWorkbook.Worksheets("PARAMETRES").Range("E97").Value
    .Fichier = ThisWorkbook.Worksheets("PARAMETRES").Range("E99").Value
    .Onglet = ThisWorkbook.Worksheets("PARAMETRES").Range("E101").Value
End With

End Sub

Sub LitHypotheses()

Dim FeuilAct As Worksheet

Dim CompteurProd, CompteurAnnee As Integer
ReDim ChPrime_Prev(0 To NbProd, 0 To Anciennete) '**
ReDim ChDeces(0 To NbProd, 0 To Anciennete)
ReDim ChTirage(0 To NbProd, 0 To Anciennete)
ReDim ChRachatTot(0 To NbProd, 0 To Anciennete)
ReDim ChRachatPart(0 To NbProd, 0 To Anciennete)
ReDim ChPrestTerme(0 To NbProd, 0 To Anciennete)
ReDim ChPM_Prev(0 To NbProd, 0 To Anciennete)   '**
ReDim ChArb_Euro_UC(0 To NbProd, 0 To Anciennete)
ReDim ChArb_UC_Euro(0 To NbProd, 0 To Anciennete)
ReDim ChArb_UC_UC(0 To NbProd, 0 To Anciennete)
ReDim TxTirage(0 To NbProd, 0 To Anciennete)
ReDim TxRachatTot(1 To NbCat, 0 To Anciennete) 'Modif
ReDim TxRachatTot_Prev(1 To NbCat, 0 To Anciennete) 'Modif
ReDim TxRachatPart(1 To NbCat, 0 To Anciennete)  'Modif
ReDim TxPass_Euro_UC(0 To NbProd, 0 To Anciennete)
ReDim TxPass_UC_Euro(0 To NbProd, 0 To Anciennete)
ReDim TxPass_UC_UC(0 To NbProd, 0 To Anciennete)
ReDim TxTechPré2011(0 To NbProd, 0 To Anciennete)
ReDim TxTechPost2011(0 To NbProd, 0 To Anciennete)
ReDim ChargPrimeUC(1 To Horizon)
ReDim ChargDecesUC(1 To Horizon)
ReDim ChargTirageUC(1 To Horizon)
ReDim ChargRachatTotUC(1 To Horizon)
ReDim ChargRachatPartUC(1 To Horizon)
ReDim ChargTermeUC(1 To Horizon)
ReDim ChargPM_UC(1 To Horizon)
ReDim ChargTransf_UC_Euro(1 To Horizon)
ReDim ChargTransf_UC_UC(1 To Horizon)
ReDim BDD(1 To NbContrats)
ReDim Totaux_Par_MP(0 To NbModelPoint, 1 To Horizon)
ReDim Totaux_Pour_Affichage(0 To Horizon)
'format
ReDim Affichage_Macro(0 To Horizon)
ReDim LancerChoc(0 To 6)
ReDim qxContrats(1 To NbContrats, 1 To Horizon)
ReDim qxChoques(1 To NbContrats, 1 To Horizon, 0 To NbChocs)

ReDim BDR(1 To NbCat, 0 To Anciennete) '''''MODIFIE''''
ReDim RDF(1 To NbModelPoint, 1 To Horizon, 0 To NbChocs)

ReDim PM_Prev0(0 To NbModelPoint) '***
ReDim TxProrog(0 To NbProd)
ReDim DuréeProrog(0 To NbProd)

Set FeuilAct = ThisWorkbook.Worksheets("HYPOTHESES")

For CompteurProd = 11 To 70
    For CompteurAnnee = 0 To Anciennete
        ChPrime_Prev(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd, 2 + CompteurAnnee).Value '**
        ChDeces(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 72).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 72, 2 + CompteurAnnee).Value
        ChTirage(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 144).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 144, 2 + CompteurAnnee).Value
        ChRachatTot(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 216).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 216, 2 + CompteurAnnee).Value
        ChRachatPart(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 288).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 288, 2 + CompteurAnnee).Value
        ChPrestTerme(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 360).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 360, 2 + CompteurAnnee).Value
        ChPM_Prev(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 432).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 432, 2 + CompteurAnnee).Value   '**
        ChArb_Euro_UC(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 504).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 504, 2 + CompteurAnnee).Value
        ChArb_UC_Euro(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 576).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 576, 2 + CompteurAnnee).Value
        ChArb_UC_UC(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 648).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 648, 2 + CompteurAnnee).Value
        TxTirage(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 720).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 720, 2 + CompteurAnnee).Value
        'TxRachatTot(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 792).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 792, 2 + CompteurAnnee).Value
        'TxRachatTot_Prev(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 864).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 864, 2 + CompteurAnnee).Value '**
        'TxRachatPart(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 936).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 936, 2 + CompteurAnnee).Value
        TxPass_Euro_UC(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 1008).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 1008, 2 + CompteurAnnee).Value
        TxPass_UC_Euro(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 1080).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 1080, 2 + CompteurAnnee).Value
        TxPass_UC_UC(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 1152).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 1152, 2 + CompteurAnnee).Value
        TxTechPré2011(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 1224).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 1224, 2 + CompteurAnnee).Value
        TxTechPost2011(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 1296).Value), CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 1296, 2 + CompteurAnnee).Value
    Next CompteurAnnee
    TxProrog(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 1423).Value)) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 1423, 2).Value
    DuréeProrog(TransformNomProd(ThisWorkbook.Worksheets("HYPOTHESES").Range("A" & CompteurProd + 1423).Value)) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurProd + 1423, 3).Value
Next CompteurProd

For CompteurAnnee = 0 To Horizon
    VecteurPB(CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(1498, 2 + CompteurAnnee).Value
    VecteurRendUC(CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(1502, 2 + CompteurAnnee).Value
Next CompteurAnnee

For CompteurCat = 1 To NbCat  '5 categories de Rachat Total
    For CompteurAnnee = 0 To Anciennete
        TxRachatTot(CompteurCat, CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurCat + 802, 2 + CompteurAnnee).Value
    Next CompteurAnnee
Next CompteurCat

For CompteurCat = 1 To NbCat '5 categories de Rachat Partiel (Tx nuls)
    For CompteurAnnee = 0 To Anciennete
        TxRachatPart(CompteurCat, CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurCat + 946, 2 + CompteurAnnee).Value
    Next CompteurAnnee
Next CompteurCat

For CompteurCat = 1 To NbCat  '5 categories de Rachat Total Prevoyance (Tx nuls)
    For CompteurAnnee = 0 To Anciennete
        TxRachatTot_Prev(CompteurCat, CompteurAnnee) = ThisWorkbook.Worksheets("HYPOTHESES").Cells(CompteurCat + 874, 2 + CompteurAnnee).Value
    Next CompteurAnnee
Next CompteurCat


End Sub

Public Sub LitHypothesesMortalite()

Dim CompteurSexe, CompteurAge As Integer
Erase BDC
ReDim BDC(1 To 2, 0 To Age_Maxi)


For CompteurSexe = 1 To 2
    TableUtilisee(CompteurSexe) = TransformTable(ThisWorkbook.Worksheets("HYPOTHESES").Range("B" & CompteurSexe + 3).Value)
    BDC(CompteurSexe, 0).TableMortalite = ThisWorkbook.Worksheets("HYPOTHESES MORTALITE").Cells(3, TableUtilisee(CompteurSexe)).Value
    For CompteurAge = 1 To Age_Maxi
        With BDC(CompteurSexe, CompteurAge)
            .TableMortalite = ThisWorkbook.Worksheets("HYPOTHESES MORTALITE").Cells(CompteurAge + 3, TableUtilisee(CompteurSexe)).Value
            .qx = 0.5 * (BDC(CompteurSexe, CompteurAge - 1).TableMortalite + .TableMortalite)
            'définition de qx pour l'année calendaire et NON celle de l'assuré
        End With
    Next CompteurAge
Next CompteurSexe

End Sub
'********************Lecture lx
Sub LitLx()

Dim NumTMG As Long
Dim age As Integer

Erase TableCommut
'ReDim TableCommut(1 To NbTMGprev, 0 To Age_Maxi)

For NumTMG = 1 To NbTMGprev
    For age = 0 To Age_Maxi
        TableCommut(NumTMG, age).lx = ThisWorkbook.Worksheets("TABLES MORTALITE TARIF").Range("AB" & age + 8).Value
    Next age
Next NumTMG

End Sub
'*****************************

Public Sub LitData()

Dim NumLgn As Long, NumFeuille As Integer, NbLignesLecture As Double, NumLigneBDD As Double
Dim FeuilBDD As String


ReDim Donnees(1 To NbContrats + NbCales)

NumLigneBDD = 1

For NumFeuille = 1 To 2
    If NumFeuille = 1 Then
        FeuilBDD = "MODEL POINT"
        NbLignesLecture = NbContrats
    Else
        FeuilBDD = "CALES"
        NbLignesLecture = NbCales
    End If
    For NumLgn = 1 To NbLignesLecture
    If NumLgn = 42891 Then
        x = x
    End If
        With Donnees(NumLigneBDD)
            .NumAdh = ThisWorkbook.Worksheets(FeuilBDD).Range("B" & NumLgn + 1).Value
            .TypeProd = ThisWorkbook.Worksheets(FeuilBDD).Range("C" & NumLgn + 1).Value
            .NomProd = TransformNomProd(ThisWorkbook.Worksheets(FeuilBDD).Range("D" & NumLgn + 1).Value)
            If ThisWorkbook.Worksheets(FeuilBDD).Range("D" & NumLgn + 1).Value = "AFI EPARGNE" Then
                x = x
            End If
            .DateNaissance = ThisWorkbook.Worksheets(FeuilBDD).Range("E" & NumLgn + 1).Value
            .AnneeNaissance = Year(.DateNaissance)
            .Sexe = TransformSexe(ThisWorkbook.Worksheets(FeuilBDD).Range("F" & NumLgn + 1).Value)
            .DateEffet = ThisWorkbook.Worksheets(FeuilBDD).Range("G" & NumLgn + 1).Value
            .AnneeEffet = Year(.DateEffet)
            .MoisEffet = Month(.DateEffet)
            .DateEcheance = ThisWorkbook.Worksheets(FeuilBDD).Range("H" & NumLgn + 1).Value
            .AnneeEcheance = Year(.DateEcheance)
    '        .ModelPoint = TransformModelPoint(ThisWorkbook.Worksheets(FeuilBDD).Range("B" & NumLgn + 1).Value)
            .NbTetes = 1
            .PrimeCommAnnualisee = ThisWorkbook.Worksheets(FeuilBDD).Range("I" & NumLgn + 1).Value
            'code précédent avec conditions pour s'assurer que les contrats "unique" ont une prime commerciale nulle : déjà effectué dans code R
            .PrimeCommEuro = ThisWorkbook.Worksheets(FeuilBDD).Range("J" & NumLgn + 1).Value
            .PrimeCommUC = ThisWorkbook.Worksheets(FeuilBDD).Range("K" & NumLgn + 1).Value
            '.PrimeCommPrev = ThisWorkbook.Worksheets(FeuilBDD).Range("L" & NumLgn + 1).Value '**
            .PM_Tot = ThisWorkbook.Worksheets(FeuilBDD).Range("L" & NumLgn + 1).Value
            .PM_Euro = ThisWorkbook.Worksheets(FeuilBDD).Range("M" & NumLgn + 1).Value
            .PM_UC = ThisWorkbook.Worksheets(FeuilBDD).Range("N" & NumLgn + 1).Value
            '.PM_Prev = ThisWorkbook.Worksheets(FeuilBDD).Range("P" & NumLgn + 1).Value '**
            .TMG = ThisWorkbook.Worksheets(FeuilBDD).Range("O" & NumLgn + 1).Value
            '.TMGprev = ThisWorkbook.Worksheets(FeuilBDD).Range("R" & NumLgn + 1).Value  '**
            .Taux_Com_Euro = ThisWorkbook.Worksheets(FeuilBDD).Range("P" & NumLgn + 1).Value
            .Taux_Com_UC = ThisWorkbook.Worksheets(FeuilBDD).Range("Q" & NumLgn + 1).Value
            '.Taux_Com_Prev = ThisWorkbook.Worksheets(FeuilBDD).Range("V" & NumLgn + 1).Value '**
            .TxComSurEncours_Euro = ThisWorkbook.Worksheets(FeuilBDD).Range("R" & NumLgn + 1).Value
            .TxComSurEncours_UC = ThisWorkbook.Worksheets(FeuilBDD).Range("S" & NumLgn + 1).Value
            '.TxComSurEncours_Prev = ThisWorkbook.Worksheets(FeuilBDD).Range("Y" & NumLgn + 1).Value '**
            .Periodicite = ThisWorkbook.Worksheets(FeuilBDD).Range("T" & NumLgn + 1).Value
            .Taux_Chargement_PM_Euro = ThisWorkbook.Worksheets(FeuilBDD).Range("U" & NumLgn + 1).Value
            .Taux_Chargement_PM_UC = ThisWorkbook.Worksheets(FeuilBDD).Range("V" & NumLgn + 1).Value
            .Taux_Chargement_Suspens_PP = ThisWorkbook.Worksheets(FeuilBDD).Range("W" & NumLgn + 1).Value
            .Taux_Chargement_PrimeEuro = ThisWorkbook.Worksheets(FeuilBDD).Range("X" & NumLgn + 1).Value
            .Taux_Chargement_PrimeUC = ThisWorkbook.Worksheets(FeuilBDD).Range("Y" & NumLgn + 1).Value
            .Position = ThisWorkbook.Worksheets(FeuilBDD).Range("AJ" & NumLgn + 1).Value
            If .Position = "En réduction" Then
                If .Taux_Chargement_PM_Euro > 0 Then
                    .Taux_Chargement_PM_Euro = .Taux_Chargement_PM_Euro + .Taux_Chargement_Suspens_PP
                End If
                If .Taux_Chargement_PM_UC Then
                    .Taux_Chargement_PM_UC = .Taux_Chargement_PM_UC + .Taux_Chargement_Suspens_PP
                End If
            End If
            .Bonus1 = ThisWorkbook.Worksheets(FeuilBDD).Range("Z" & NumLgn + 1).Value
            .DurVerBonus1 = ThisWorkbook.Worksheets(FeuilBDD).Range("AA" & NumLgn + 1).Value
            .DelaiBonus1 = .DurVerBonus1 - (AnneeValorisation - .AnneeEffet) - 1
            If .Periodicite = "Semetriel" And .MoisEffet > 6 Then
                .DelaiBonus1 = .DelaiBonus1 + 1
            ElseIf .Periodicite = "Trimestriel" And .MoisEffet > 3 Then
                .DelaiBonus1 = .DelaiBonus1 + 1
            End If
            .Bonus2 = ThisWorkbook.Worksheets(FeuilBDD).Range("AB" & NumLgn + 1).Value
            .DurVerBonus2 = ThisWorkbook.Worksheets(FeuilBDD).Range("AC" & NumLgn + 1).Value
            .DelaiBonus2 = .DurVerBonus2 - (AnneeValorisation - .AnneeEffet) - 1
            If .Periodicite = "Semetriel" And .MoisEffet > 6 Then
                .DelaiBonus2 = .DelaiBonus2 + 1
            ElseIf .Periodicite = "Trimestriel" And .MoisEffet > 3 Then
                .DelaiBonus2 = .DelaiBonus2 + 1
            End If
            .DurVer = ThisWorkbook.Worksheets(FeuilBDD).Range("AD" & NumLgn + 1).Value
            If ThisWorkbook.Worksheets(FeuilBDD).Range("AE" & NumLgn + 1).Value = 1 Then
                .Contrat_Proro = True
            Else
                .Contrat_Proro = False
            End If
            .IndicTypeTauxMin = ThisWorkbook.Worksheets(FeuilBDD).Range("AF" & NumLgn + 1).Value
            If ThisWorkbook.Worksheets(FeuilBDD).Range("AG" & NumLgn + 1).Value = 1 Then
                .IndicObsèque = True
            Else
                .IndicObsèque = False
            End If
            .DuréeRestantePrime = .DurVer - (AnneeValorisation - .AnneeEffet) - 1
            If .Periodicite = "Semetriel" And .MoisEffet > 6 Then
                .DuréeRestantePrime = .DuréeRestantePrime + 1
            ElseIf .Periodicite = "Trimestriel" And .MoisEffet > 3 Then
                .DuréeRestantePrime = .DuréeRestantePrime + 1
            End If
            
            If .PM_UC > 0 Then
                .NbTetesUC = .NbTetes
            Else
                .NbTetesUC = 0
            End If
            If .PM_Euro > 0 Then
                .NbTetesEuro = .NbTetes
            Else
                .NbTetesEuro = 0
            End If
            
            '.CapitalDeces = ThisWorkbook.Worksheets(FeuilBDD).Range("AN" & NumLgn + 1).Value '**
            '.DureeOptionPrev = ThisWorkbook.Worksheets(FeuilBDD).Range("AM" & NumLgn + 1).Value '**
            
            .FormP = ThisWorkbook.Worksheets(FeuilBDD).Range("AL" & NumLgn + 1).Value
            .TxRetroGlobal = ThisWorkbook.Worksheets(FeuilBDD).Range("AM" & NumLgn + 1).Value
            .TxRetroAE = ThisWorkbook.Worksheets(FeuilBDD).Range("AN" & NumLgn + 1).Value
            .NbTetesCU = ThisWorkbook.Worksheets(FeuilBDD).Range("AO" & NumLgn + 1).Value
            
            'Creation des cinq categories de rachat
            If ThisWorkbook.Worksheets(FeuilBDD).Range("AG" & NumLgn + 1).Value = 1 Then
                .CatRachatTot = 1
            ElseIf ThisWorkbook.Worksheets(FeuilBDD).Range("AG" & NumLgn + 1).Value = 0 Then
                If (.Periodicite = "Libre" Or .Periodicite = "Unique") And .FormP = 1 Then
                    .CatRachatTot = 2
                End If
                If (.Periodicite = "Libre" Or .Periodicite = "Unique") And .FormP = 2 Then
                    .CatRachatTot = 3
                End If
                If (.Periodicite <> "Libre" And .Periodicite <> "Unique") And .FormP = 1 Then
                    .CatRachatTot = 4
                End If
                If (.Periodicite <> "Libre" And .Periodicite <> "Unique") And .FormP = 2 Then
                    .CatRachatTot = 5
                End If
            End If
             
            
            ''' ****** PADDOP ****** '''
                .TEST_Contrat_Sorti_N = ThisWorkbook.Worksheets(FeuilBDD).Range("AF" & NumLgn + 1).Value
            ''' ****** PADDOP ****** '''
        End With
        NumLigneBDD = NumLigneBDD + 1
    Next NumLgn
Next NumFeuille

NbContrats = NbContrats + NbCales

End Sub

'*******************************************************************************
Sub ImportCbTx()
'*******************************************************************************
Dim FeuilCbTaux As Worksheet, NumAn As Integer

With FichierCbTaux
    Workbooks.Open Filename:=.Dossier & .Fichier, ReadOnly:=True
    Set FeuilCbTaux = Workbooks(.Fichier).Worksheets(.Onglet)
End With

For NumAn = 1 To Horizon
    CbTaux(NumAn) = FeuilCbTaux.Cells(10 + NumAn, 3)
Next NumAn

Workbooks(FichierCbTaux.Fichier).Close

End Sub


'*******************************************************************************
Sub ImportInflation()
'*******************************************************************************
Dim FeuilInflation As Worksheet, NumAn As Integer

With FichierInflation
    Workbooks.Open Filename:=.Dossier & .Fichier, ReadOnly:=True
    Set FeuilInflation = Workbooks(.Fichier).Worksheets(.Onglet)
End With

For NumAn = 1 To Horizon
    Inflation(NumAn) = FeuilInflation.Cells(1003, 2 + NumAn)
Next NumAn

Workbooks(FichierInflation.Fichier).Close

End Sub

'*******************************************************************************
Sub LitCoûtsUnitaires()
'*******************************************************************************
Dim FeuilCoûtsUnitaires As Worksheet

Set FeuilCoûtsUnitaires = ThisWorkbook.Worksheets("COUTS UNITAIRES")

If FeuilCoûtsUnitaires.Cells(4, 7) <> 0 Then
    TypeCU = "Global"
    CU_Adm = FeuilCoûtsUnitaires.Cells(4, 7)
    CU_Prest = FeuilCoûtsUnitaires.Cells(6, 7)
Else
    TypeCU = "Détail"
    CU_adm_Euro = FeuilCoûtsUnitaires.Cells(4, 3)
    CU_adm_UC = FeuilCoûtsUnitaires.Cells(4, 5)
    CU_Prest_Euro = FeuilCoûtsUnitaires.Cells(6, 3)
    CU_prest_UC = FeuilCoûtsUnitaires.Cells(6, 5)
End If

End Sub

'*******************************************************************************
Sub LitCourbeTx()
'*******************************************************************************
Dim NumLgn As Integer, FeuilResultat As Worksheet, CompteurAnnee As Integer

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP EURO (NEW)")

For CompteurAnnee = 1 To Horizon
    CbTaux(CompteurAnnee) = FeuilResultat.Cells(14, 6 + CompteurAnnee)
Next CompteurAnnee


End Sub


'*******************************************************************************
Sub LitInflation()
'*******************************************************************************
Dim NumLgn As Integer, FeuilResultat As Worksheet, CompteurAnnee As Integer

Set FeuilResultat = ThisWorkbook.Worksheets("TOTAL - EP EURO (NEW)")

For CompteurAnnee = 1 To Horizon
    Inflation(CompteurAnnee) = FeuilResultat.Cells(524, 6 + CompteurAnnee)
Next CompteurAnnee



End Sub
