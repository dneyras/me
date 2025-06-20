'*******************************************************************************
Option Explicit


Global Const Horizon = 50
Global Auxdateproj(0 To Horizon) As Date
Global Const Anciennete = 100
Global Const Age_Maxi = 200
Global Const NbErreurs = 3
Global Const NbTables = 18
Global Const NbModelPointPRIIPS = 15 '* Rajout du MP "AINV Option prévoyance + ---> Les PRIIPS
Global Const NbModelPointHorsPRIIPS = 10
Global NbModelPoint As Integer
Global Const NbProd = 52  ' Estelle 08/17 rajout 2 noms de produits Lille
Global Const NbCat = 5
Global Const NbChocs = 6
Global Const NbLob = 2, NbNiveauxPB = 1

Global BDE As BaseErreurs
Global Donnees() As BaseData
Global DonneesProrogee() As BaseData
Global BDD() As BaseContrats, BDC() As BaseDeces, BDR() As BaseDeRachats
Global Totaux_Par_MP() As Sommes, RDF() As RatiosDeFlexing

Global AnneeInitiale As Integer

Global FlexingEuro() As Flexing_Modeling, FlexingUC() As Flexing_Modeling ' Nouveaux Outputs de flexing pour Modeling
Global NbTMG As Integer, TMG_Tri() As Double  ' Nouveaux Outputs flexing
Global ChocAddactis(0 To 6) As Boolean ' E.A. Addactis pour choisir d'appliquer le choc ou pas, problématique de compensation TRUE : Choc / FALSE : Central

Global ChPrime_Prev() As Double, ChPM_Prev() As Double, TxRachatTot_Prev() As Double '**
Global NbTMGprev As Long, TMGprev() As Double, PM_Prev0() As Double       '***
Global ChPrime() As Double, ChDeces() As Double, ChTirage() As Double
Global ChRachatTot() As Double, ChRachatPart() As Double, ChPrestTerme() As Double
Global ChPM_Euro() As Double, ChPM_UC() As Double
Global ChArb_Euro_UC() As Double, ChArb_UC_Euro() As Double, ChArb_UC_UC() As Double
Global TxTirage() As Double, TxRachatTot() As Double, TxRachatPart() As Double
Global TxPass_Euro_UC() As Double, TxPass_UC_Euro() As Double, TxPass_UC_UC() As Double
Global TxTechPré2011() As Variant, TxTechPost2011() As Variant
Global TxProrog() As Double, DuréeProrog() As Double
Global AnneeValorisation As Integer, DateValorisation As Date, NbContrats As Long, NbCales As Integer
Global ChocMortalite As Double, ChocLongevite, ChocCatastrophe As Double, ChocFrais As Double
Global ChocHausse As Double, ChocBaisse As Double, ChocMassif As Double
Global ScenCent As Boolean, ScenMort As Boolean, ScenLong As Boolean, ScenFrais As Boolean
Global ScenHausse As Boolean, ScenBaisse As Boolean, ScenCata As Boolean, ScenMassif As Boolean
Global TableUtilisee(1 To 2) As Integer
Global DossierOutPut As String, FichierOutPut As String
Global DossierOutPutModeling As String, FichierOutPutModeling As String
Global qxContrats() As Double
Global qxChoques() As Double
Global TableCommut() As TableauCommutations '**
Global NbPriips As Integer, NomsPriips() As String ' Estelle 08/17
Global Totaux_Pour_Affichage() As Sommes
Global AgeMP As Boolean
Global FichierCbTaux As FichierAOuvrir
Global FichierInflation As FichierAOuvrir
Global VecteurPB(0 To Horizon) As Double
Global VecteurRendUC(0 To Horizon) As Double
Global CU_adm_Euro As Double, CU_adm_UC As Double
Global CU_Prest_Euro As Double, CU_prest_UC As Double
Global TypeCU As String
Global CU_Adm As Double, CU_Prest As Double



'Option pour les runs
Global FichierSorties As Boolean
Global FichierSortiesModeling As Boolean
Global Perimetre As String
Global TypePrime As String
Global Obsèque_Epargne As String
Global TypeTauxMin As String
Global PrimeOuSans As String
Global ProjParTMG As Double

Global CbTaux(1 To Horizon) As Double
Global Inflation(1 To Horizon) As Double

Global NumChoc As Integer
Global LancerChoc() As Boolean
Global CompteurAnnee As Integer

Public Type TableauCommutations  '**
    lx As Double
    Dx As Double
    Cx  As Double
    Mx As Double
End Type
    
Public Type BaseErreurs
    PresenceErreur As Boolean
    ErreurPresente(1 To NbErreurs) As Boolean
    ExplicationErreur(1 To NbErreurs) As String
End Type

Public Type FichierAOuvrir
    Dossier As String
    Fichier As String
    Onglet As String
End Type

Public Type BaseData
    'ModelPoint As Integer
    NumAdh As Double
    TypeProd As String
    NomProd As Integer
    NbTetes As Double
    NbTetesEuro As Double
    NbTetesUC As Double
    NbTetesCU As Double
    'NbPartsUC As Double
    DateNaissance As Date
    AnneeNaissance As Integer
    Sexe As Integer
    DateEffet As Date
    AnneeEffet As Integer
    MoisEffet As Integer
    DateEcheance As Date
    AnneeEcheance As Integer
    PrimeCommAnnualisee As Double
    PrimeCommEuro As Double
    PrimeCommUC As Double
    'PrimeCommPrev As Double '**
    PM_Tot As Double
    PM_Euro As Double
    PM_UC As Double
    'PM_Prev As Double  '**
    TMG As Double
    'TMGprev As Double  '**
    Taux_Com_Euro As Double
    Taux_Com_UC As Double
    'Taux_Com_Prev As Double '**
    TxComSurEncours_Euro As Double
    TxComSurEncours_UC As Double
    'TxComSurEncours_Prev As Double '**
    Periodicite As String
    Taux_Chargement_PM_Euro As Double
    Taux_Chargement_PM_UC As Double
    'Taux_Chargement_PM_Lille_Euro As Double
    'Taux_Chargement_PM_Lille_UC As Double
    'Taux_Chargement_PM_Euro As Double
    'Taux_Chargement_PM_UC As Double
        'Nouveau
    Taux_Chargement_Suspens_PP As Double
    Taux_Chargement_PrimeEuro As Double
    Taux_Chargement_PrimeUC As Double
    Bonus1 As Double
    Bonus2 As Double
    DurVerBonus1 As Integer
    DurVerBonus2 As Integer
    DelaiBonus1 As Integer
    DelaiBonus2 As Integer
    DurVer As Integer
    Contrat_Proro As Boolean
    IndicTypeTauxMin As Integer
    IndicObsèque As Boolean
    DureeRestante As Double
    DuréeRestantePrime As Double
    
    FormP As Integer
    CatRachatTot As Integer
    TxRetroGlobal As Double
    TxRetroAE As Double
    Position As String
    
    'TEST_Contrat_STBG As Double
    'TEST_Contrat_LILLE As Double
    
    'CapitalDeces As Double  '**
    'DureeOptionPrev As Integer  '**
    
    ''' ****** PADDOP ****** '''
        TEST_Contrat_Sorti_N As Integer
    ''' ****** PADDOP ****** '''
End Type

Public Type BaseContrats
    IndicObsèque As Boolean
    IndicTypeTauxMin As Integer
    FormP As Integer
    CatRachatTot As Integer
    ModelPoint As Integer
    NomProd As Integer
    NbTetes As Double
    NbTetesCU As Double
    NbClotures(0 To Horizon) As Double
    NbClotures_Euro(0 To Horizon) As Double
    NbClotures_UC(0 To Horizon) As Double
    NbClotures_Prev(0 To Horizon) As Double  '**
    NbDeces(1 To Horizon) As Double
    NbDeces_Prev(0 To Horizon) As Double '**
    NbTirages(1 To Horizon) As Double
    NbRachatsTot(1 To Horizon) As Double
    NbRachatsTot_Prev(1 To Horizon) As Double '**
    NbRachatsPart(1 To Horizon) As Double
    NbTermes(1 To Horizon) As Double
    NbTermes_Prev(1 To Horizon) As Double '**
    Sexe As Integer
    AnneeNaissance As Integer
    AgeAssure(0 To Horizon) As Integer
    AnneeEffet As Integer
    DateEffet As Date
    AncienneteContrat(0 To Horizon) As Integer
    AnneeEcheance As Integer
    DateEcheance As Date
    DureeRestante(0 To Horizon) As Integer
    DureeRestantePrime(0 To Horizon) As Integer
    DureeRestante_Prev(0 To Horizon) As Integer '**
    ResteContrat(0 To Horizon) As Integer
    IndTermeContrat(0 To Horizon) As Integer
    PrimeCommEuro(0 To Horizon) As Double
    PrimeCommEuroPourAffichageAnnee0 As Double
    ChargPrimeEuro(0 To Horizon) As Double
    PrimeNetteEuro(1 To Horizon) As Double
    PrimeCommUC(0 To Horizon) As Double
    PrimeCommUCPourAffichageAnnee0 As Double
    PrimeBruteUC(1 To Horizon) As Double
    ChargPrimeUC(0 To Horizon) As Double
    PrimeNetteUC(1 To Horizon) As Double
    PrimeCommPrev(0 To Horizon) As Double '**
    ChargPrimePrev(1 To Horizon) As Double '**
    PrimeNettePrev(0 To Horizon) As Double '**
    TMG(0 To Horizon) As Variant
    TMGprev As Double         '**
    PM_Euro(0 To Horizon) As Double
    'PM clôture 2 dans la feuille Proj Tête par Tête (Projection Euro)
    PM_UC(0 To Horizon) As Double
    'PM clôture 2 dans la feuille Proj Tête par Tête (Projection UC)
    PM_Prev(0 To Horizon) As Double '**
    InteretsPM_MiPeriodeEuro(1 To Horizon) As Double
    InteretsPrimeEuro(1 To Horizon) As Double
    PB_MIPeriodeEuro(1 To Horizon) As Double
    PBPrimeEuro(1 To Horizon) As Double
    PM_MiPeriode1Euro(1 To Horizon) As Double
    SinDecesEuro(1 To Horizon) As Double
    SinTirageEuro(1 To Horizon) As Double
    SinRachatTotEuro(1 To Horizon) As Double
    SinRachatPartEuro(1 To Horizon) As Double
    ChargDecesEuro(1 To Horizon) As Double
    ChargTirageEuro(1 To Horizon) As Double
    ChargRachatTotEuro(1 To Horizon) As Double
    ChargRachatPartEuro(1 To Horizon) As Double
    PM_MiPeriode2Euro(1 To Horizon) As Double
    SinTermeEuro(1 To Horizon) As Double
    ChargTermeEuro(1 To Horizon) As Double
    PM_MiPeriode3Euro(1 To Horizon) As Double
    PM_MiPeriode4Euro(1 To Horizon) As Double
    InteretsFinPeriodeEuro(1 To Horizon) As Double
    PBFinPeriodeEuro(1 To Horizon) As Double
    PM_ClotureEuro(0 To Horizon) As Double
    ChargPM_Euro(1 To Horizon) As Double
    InteretsPM_MiPeriodeUC(1 To Horizon) As Double
    InteretsPrimeUC(1 To Horizon) As Double
    RendUCPrime(1 To Horizon) As Double
    RendUC_MiPeriodeUC(1 To Horizon) As Double
    PM_MiPeriode1UC(1 To Horizon) As Double
    SinDecesUC(1 To Horizon) As Double
    SinTirageUC(1 To Horizon) As Double
    SinRachatTotUC(1 To Horizon) As Double
    SinRachatPartUC(1 To Horizon) As Double
    ChargDecesUC(1 To Horizon) As Double
    ChargTirageUC(1 To Horizon) As Double
    ChargRachatTotUC(1 To Horizon) As Double
    ChargRachatPartUC(1 To Horizon) As Double
    PM_MiPeriode2UC(1 To Horizon) As Double
    SinTermeUC(1 To Horizon) As Double
    ChargTermeUC(1 To Horizon) As Double
    PM_MiPeriode3UC(1 To Horizon) As Double
    PM_MiPeriode4UC(1 To Horizon) As Double
    InteretsFinPeriodeUC(1 To Horizon) As Double
    RendFinPeriodeUC(1 To Horizon) As Double
    PM_ClotureUC(0 To Horizon) As Double
    ChargPM_UC(1 To Horizon) As Double
    RetroGlobalPM_UC(1 To Horizon) As Double
    RetroAEPM_UC(1 To Horizon) As Double
    NbPartsUC As Double
    Cap_Euro_UC(1 To Horizon) As Double
    ChargTransf_Euro_UC(1 To Horizon) As Double
    Cap_UC_Euro(1 To Horizon) As Double
    ChargTransf_UC_Euro(1 To Horizon) As Double
    ChargTransf_UC_UC(1 To Horizon) As Double
    SinDecesPrev(1 To Horizon) As Double  '**
    SinRachatTotPrev(1 To Horizon) As Double  '**
    ChargDecesPrev(1 To Horizon) As Double  '**
    ChargRachatTotPrev(1 To Horizon) As Double  '**
    ChargPM_Prev(1 To Horizon) As Double '**
    MargEuro(1 To Horizon) As Double
    ChargEuro(1 To Horizon) As Double
    MargUC(1 To Horizon) As Double
    ChargUC(1 To Horizon) As Double
    MargPrev(1 To Horizon) As Double '**
    ChargPrev(1 To Horizon) As Double '**
    Taux_Com_Euro As Double
    Taux_Com_UC As Double
    Taux_Com_Prev As Double '**
    TxComSurEncours_Euro As Double
    TxComSurEncours_UC As Double
    TxComSurEncours_Prev As Double '**
    TxRétroGlobal As Double
    TxRétroAE As Double
    Commissions_PrimesEuro(0 To Horizon) As Double
    Commissions_PMeuro(1 To Horizon) As Double
    Commissions_PrimesUC(0 To Horizon) As Double
    Commissions_PMuc(1 To Horizon) As Double
    Commissions_PrimesPrev(1 To Horizon) As Double '**
    Commissions_PMprev(1 To Horizon) As Double  '**
    BonusRetEuro(1 To Horizon) As Double
    BonusRetUC(1 To Horizon) As Double
    Bonus1(0 To Horizon) As Double
    Bonus2(0 To Horizon) As Double
    DelaiBonus1(0 To Horizon) As Integer
    DelaiBonus2(0 To Horizon) As Integer
    Periodicite As String
    CoeffPrimeAnneeEch(1 To Horizon) As Double
    DateAnniversaire(0 To Horizon) As Date
    CoeffPrimeAnneeBonus(1 To Horizon) As Double
    CapitalDeces(0 To Horizon) As Double '**
    CoeffPrimePrev As Double '**
    AgeMaxprev As Integer '**
    Taux_Chargement_PM_Euro As Double
    Taux_Chargement_PM_UC As Double
    Taux_Chargement_Prime_Euro As Double
    Taux_Chargement_Prime_UC As Double
    CoutUnitaireGestion_Euro(0 To Horizon) As Double
    CoutUnitairePresta_Euro(0 To Horizon) As Double
    CoutUnitaireGestion_UC(0 To Horizon) As Double
    CoutUnitairePresta_UC(0 To Horizon) As Double
'    ''' ************************************ Chargement PM Contrats Lillois ************************************ '''
'        TEST_Contrat_STBG As Double
'        TEST_Contrat_LILLE As Double
'        Taux_Chargement_PM_Lille_Euro As Double
'        Taux_Chargement_PM_Lille_UC As Double
'    ''' ************************************ Chargement PM Contrats Lillois ************************************ '''
    ''' ****** PADDOP ****** '''
        TEST_Contrat_Sorti_N As Integer
        TEST_NouveauContrat As Integer
    ''' ****** PADDOP ****** '''
End Type

Public Type BaseDeces
    TableMortalite As Double
    qx As Double
    'qx défini pour l'année calendaire et NON celle de l'assuré
End Type

Public Type BaseDeRachats
    ProbaRachatTot As Double
    ProbaRachatPart As Double
    ProbaRachatTot_Prev As Double '**
End Type

Public Type Sommes
    Somm_BonusRetEuro As Double
    Somm_BonusRetUC As Double
    Somm_Capital_Garanti As Double '**
    Somm_PrimeCommEuro As Double
    Somm_PrimeCommPrev As Double '***
    Somm_ChargPrimeEuro As Double
    Somm_ChargPrimePrev As Double '***
    Somm_PM_OuvertureEuro As Double
    Somm_InteretsPM_MiPeriodeEuro As Double
    Somm_PB_MiPeriodeEuro As Double
    Somm_PM_MiPeriode1Euro As Double
    Somm_SinDecesEuro As Double
    Somm_SinTirageEuro As Double
    Somm_SinRachatTotEuro As Double
    Somm_SinRachatPartEuro As Double
    Somm_ChargDecesEuro As Double
    Somm_ChargTirageEuro As Double
    Somm_ChargRachatTotEuro As Double
    Somm_ChargRachatPartEuro As Double
    Somm_PM_MiPeriode2Euro As Double
    Somm_SinTermeEuro As Double
    Somm_ChargTermeEuro As Double
    Somm_PM_MiPeriode3Euro As Double
    Somm_PM_MiPeriode4Euro As Double
    Somm_InteretsFinPeriodeEuro As Double
    Somm_PBFinPeriodeEuro As Double
    Somm_PM_ClotureEuro As Double
    Somm_ChargPM_Euro As Double
    Somm_PrimeCommUC As Double
    Somm_RendUC As Double
    Somm_ChargPrimeUC As Double
    Somm_PM_OuvertureUC As Double
    Somm_InteretsPM_MiPeriodeUC As Double
    Somm_Rend_MiPeriodeUC As Double
    Somm_PM_MiPeriode1UC As Double
    Somm_SinDecesUC As Double
    Somm_SinTirageUC As Double
    Somm_SinRachatTotUC As Double
    Somm_SinRachatPartUC As Double
    Somm_ChargDecesUC As Double
    Somm_ChargTirageUC As Double
    Somm_ChargRachatTotUC As Double
    Somm_ChargRachatPartUC As Double
    Somm_PM_MiPeriode2UC As Double
    Somm_SinTermeUC As Double
    Somm_ChargTermeUC As Double
    Somm_PM_MiPeriode3UC As Double
    Somm_PM_MiPeriode4UC As Double
    Somm_InteretsFinPeriodeUC As Double
    Somm_RendFinPeriodeUC As Double
    Somm_PM_ClotureUC As Double
    Somm_ChargPM_UC As Double
    Somm_RetroGlobalPM_UC As Double
    Somm_RetroAEPM_UC As Double
    Somm_Cap_Euro_UC As Double
    Somm_ChargTransf_Euro_UC As Double
    Somm_Cap_UC_Euro As Double
    Somm_ChargTransf_UC_Euro As Double
    Somm_ChargTransf_UC_UC As Double
    Somm_PM_CloturePrev As Double     '***
    Somm_NbOuvertures As Double
    Somm_NbDeces As Double
    Somm_NbDecesEuro As Double
    Somm_NbDecesUC As Double
    Somm_NbTirages As Double
    Somm_NbTiragesEuro As Double
    Somm_NbTiragesUC As Double
    Somm_NbRachatsTot As Double
    Somm_NbRachatsTotEuro As Double
    Somm_NbRachatsTotUC As Double
    Somm_NbRachatsPart As Double
    Somm_NbRachatsPartEuro As Double
    Somm_NbRachatsPartUC As Double
    Somm_NbTermes As Double
    Somm_NbTermesEuro As Double
    Somm_NbTermesUC As Double
    Somm_NbOuverturesEuro As Double '***
    Somm_NbOuverturesUC As Double '***
    Somm_NbOuverturesPrev As Double '***
    Somm_Commissions_PrimesEuro As Double
    Somm_Commissions_PMeuro As Double
    Somm_Commissions_PrimesUC As Double
    Somm_Commissions_PMuc As Double
    Somm_FraisAdmEuro As Double
    Somm_FraisPrestEuro As Double
    Somm_FraisAdmUC As Double
    Somm_FraisPrestUC As Double
End Type

Public Type RatiosDeFlexing
    TxChargDecesEuro As Double
    ProbaDecesEuro As Double
    TxChargTirageEuro As Double
    TxTirageEuro As Double
    TxChargRachatTotEuro As Double
    ProbaRachatTotEuro As Double
    TxChargRachatPartEuro As Double
    ProbaRachatPartEuro As Double
    TxChargContratsTermesEuro As Double
    ProbaContratsTermesEuro As Double
    TxChargPM_Euro As Double
    TxTechDebutPeriodeEuro As Double
    TxTechFinPeriodeEuro As Double
    EvolutionPrimesEuro As Double
    TxChargDecesUC As Double
    ProbaDecesUC As Double
    TxChargTirageUC As Double
    TxTirageUC As Double
    TxChargRachatTotUC As Double
    ProbaRachatTotUC As Double
    TxChargRachatPartUC As Double
    ProbaRachatPartUC As Double
    TxChargContratsTermesUC As Double
    ProbaContratsTermesUC As Double
    TxChargPM_UC As Double
    TxRetroGlobalPM_UC As Double
    TxRetroAEPM_UC As Double
    TxTechDebutPeriodeUC As Double
    TxTechFinPeriodeUC As Double
    EvolutionPrimesUC As Double
    TxChargPrimesEuro As Double
    TxChargPrimesUC As Double
    TxChargPrimesPrev As Double  '***
    TxChargPass_Euro_UC As Double
    TxPass_Euro_UC As Double
    TxChargPass_UC_Euro As Double
    TxPass_UC_Euro As Double
    ChargProbabilise_UC_UC As Double
    NbOuvertures As Double
    NbOuverturesEuro As Double
    NbOuverturesUC As Double
    NbOuverturesPrev As Double  '***
    PrimeCommEuro As Double
    PM_OuvertureEuro As Double
    PrimeCommUC As Double
    PM_OuvertureUC As Double
    PrimeCommPrev As Double '***
    PM_CloturePrev As Double  '***
End Type

Public Type Flexing_Modeling
    RachDyn As Integer
    LoB As Integer
    PB As Double
    TMG As Double
    Code_support As String
    Presta_Deces As Double
    Presta_Rachat As Double
    Presta_Terme As Double
    Presta_Tirage As Double
    Presta_Autres As Double
    Cotisations As Double
    Bonus As Double
    PM As Double
    IT As Double
    PS As Double
    Chgt_Cotis As Double
    Chgt_Encours As Double
    Chgt_Presta As Double
    Chgt_Autres As Double
    Chgt_Presta_deces As Double
    Chgt_Presta_rachat As Double
    Chgt_Presta_terme As Double
    Chgt_Presta_tirage As Double
    Retro As Double
    RetroAE As Double
    Frais_Cotis As Double
    Frais_Encours As Double
    Frais_Presta As Double
    Frais_Autres As Double
    Indemnites_Cotis As Double
    Indemnites_Encours As Double
    Indemnites_Presta As Double
    Comm_Autres As Double
    Effectifs As Double
End Type


Sub Principale()

'Application.ScreenUpdating = False

AgeMP = True
Erase Totaux_Par_MP
Erase Totaux_Pour_Affichage
NbModelPoint = NbModelPointHorsPRIIPS

Initialisation.AbsenceErreur
Initialisation.LitParametres        '**
Initialisation.LitCourbeTx
Initialisation.LitInflation
Initialisation.ComptNbContrats
Initialisation.LitData
'Ecriture.EcritRésultatsCatRachat
'Initialisation.ComptNbTMGprev '**
Initialisation.LitHypotheses '*
'Ecriture.EcritRésultatsTxRachat
'Ecriture.EcritRésultatsTxRachat2
Initialisation.LitHypothesesMortalite
Initialisation.LitLx
Initialisation.LitCoûtsUnitaires

If ScenCent = False And ScenMort = False And ScenLong = False And ScenMassif = False And _
            ScenHausse = False And ScenBaisse = False And ScenCata = False Then
    BDE.PresenceErreur = True
    BDE.ErreurPresente(1) = True
End If

If BDE.PresenceErreur = True Then
    Erreurs.DefinitionErreurs
    Erreurs.TraitErreur
    GoTo FinSub
End If


CalculsPreliminaires_parContrat  '*
CalculsRedimQx
CalculsTableCommut             '**
CalculsLancerChoc


For NumChoc = 0 To NbChocs
    If LancerChoc(NumChoc) = True Then
        CalculsChocs_Qx '*
        CalculsChocs_TauxRachat '*
        For CompteurAnnee = 1 To Horizon
            If Perimetre = "NB" Then
                Calculs.CalculsPourAffichageAnnée0
            End If
            CalculsNbDeces_NbTirages_NbRachatsTot_NbRachatsPart_NbTermes_NbClotures '*
            CalculsBonus
            CalculsCoeffPrimeAnneeEch
            CalculsCoeffPrimeAnneeBonus
'            CalculsCoeffPrimePrev '**
            CalculsTMG
            CalculsPrimeEuro
            CalculsChargPrime_Euro
            CalculsCommissionPrime_Euro
            CalculsPrimeNette_Euro
            CalculsPM_MiPeriode1_Euro
            CalculsSinistres_Euro
            CalculsChargementsSinistres_Euro
            CalculsSinistreTerme_ChargementTerme_Euro
            CalculsPrime_UC
            CalculsChargPrime_UC
            CalculsCommissionPrime_UC
            CalculsPrimeNette_UC
            CalculsPM_MiPeriode1_UC
            CalculsSinistres_UC
            CalculsChargementsSinistres_UC
            CalculsSinitreTerme_ChargementTerme_UC
            CalculsTransfertsCapitaux_Chargements_Euro_UC
'            CalculsPrime_Prev '**
'            CalculsChargPrime_Prev '**
'            CalculsCommissionPrime_Prev '**
'            CalculsPrimeNette_Prev '**
'            CalculsSinistres_Prev '**
'            CalculsChargementsSinistres_Prev '**
            CalculsPMCloture_Euro
            CaculsChargPM_Euro
            CalculsCommissionPM_Euro
            CalculsPM_Euro
            CalculsMarg_Euro
            CalculsCharg_Euro
            CalculsTransfertsCapitaux_Chargements_UC_Euro
            CalculsPMCloture_UC
            CalculsPM_UC
            CaculsChargPM_UC
            CalculsCommissionPM_UC
            CalculsMargUC
            CalculsChargUC
            Calculs.CalculsCoutUnitaireGestion
'            CaculsChargPM_Prev '**
'            CalculsCommissionPM_Prev '**
'            CalculsPM_Prev '**
'            CalculsMarg_Prev '**
'            CalculsCharg_Prev '**
            CalculsSommes '*
            CalculsSommesPourAffichage
            CalculsRatiosFlexing '*
'            If CompteurAnnee = 1 Then
'                Ecriture.EcritDPMContrat
'            End If
        Next CompteurAnnee
        Exit Sub
        '*** Nouvelle proc pour remplir les Outputs flexing Modeling ****
        Calculs.StockFlexingModeling
        If NumChoc = 0 Then         ' if ScenCent = True
            Call EcritRésultatsFlexingModeling("CENTRAL", DossierOutPutModeling, FichierOutPutModeling)
        ElseIf NumChoc = 1 Then     ' if ScenMort = True
            Call EcritRésultatsFlexingModeling("MORTALITE", DossierOutPutModeling, FichierOutPutModeling)
        ElseIf NumChoc = 2 Then     ' if ScenLong = True
            Call EcritRésultatsFlexingModeling("LONGEVITE", DossierOutPutModeling, FichierOutPutModeling)
        ElseIf NumChoc = 3 Then     ' if ScenCata = True
            Call EcritRésultatsFlexingModeling("CATASTROPHE", DossierOutPutModeling, FichierOutPutModeling)
        ElseIf NumChoc = 4 Then     ' if ScenHausse = True
            Call EcritRésultatsFlexingModeling("HAUSSE", DossierOutPutModeling, FichierOutPutModeling)
        ElseIf NumChoc = 5 Then     ' if ScenBaisse = True
            Call EcritRésultatsFlexingModeling("BAISSE", DossierOutPutModeling, FichierOutPutModeling)
        ElseIf NumChoc = 6 Then     ' if ScenBaisse = True
            Call EcritRésultatsFlexingModeling("MASSIF", DossierOutPutModeling, FichierOutPutModeling)
        End If
        '*****************************************************************
'        CalculsPMprev0 '***
        If FichierSorties = True Then
            If NumChoc = 0 Then         ' if ScenCent = True
                Call ExportResultats.ExportResultats("CENTRAL", DossierOutPut, FichierOutPut)
            ElseIf NumChoc = 1 Then     ' if ScenMort = True
                Call ExportResultats.ExportResultats("MORTALITE", DossierOutPut, FichierOutPut)
            ElseIf NumChoc = 2 Then     ' if ScenLong = True
                Call ExportResultats.ExportResultats("LONGEVITE", DossierOutPut, FichierOutPut)
            ElseIf NumChoc = 3 Then     ' if ScenCata = True
                Call ExportResultats.ExportResultats("CATASTROPHE", DossierOutPut, FichierOutPut)
            ElseIf NumChoc = 4 Then     ' if ScenHausse = True
                Call ExportResultats.ExportResultats("HAUSSE", DossierOutPut, FichierOutPut)
            ElseIf NumChoc = 5 Then     ' if ScenBaisse = True
                Call ExportResultats.ExportResultats("BAISSE", DossierOutPut, FichierOutPut)
            End If
        End If
        EcritRésultats
        EcritRésultatsNew
        EcritRésultatsNewCourbeTx
        EcritRésultatsInflation
        
    End If
    Erase Totaux_Par_MP
    Erase Totaux_Pour_Affichage
    ReDim Totaux_Par_MP(0 To NbModelPoint, 1 To Horizon)
    ReDim Totaux_Pour_Affichage(0 To Horizon)
Next NumChoc

If FichierSorties = True Then
    Workbooks(FichierOutPut).Save
    Workbooks(FichierOutPut).Close
End If

Ecriture.EcritInfosExécution
Application.ScreenUpdating = True

FinSub:

End Sub


Sub ImportCbTx_Inflation()
    Initialisation.LitParametres
    Initialisation.ImportCbTx
    Initialisation.ImportInflation
    Ecriture.EcritRésultatsNewCourbeTx
    Ecriture.EcritRésultatsInflation
End Sub

