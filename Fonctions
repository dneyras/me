Public Function TransformTable(Table As String) As Integer

Dim NumCol As Integer


For NumCol = 1 To NbTables
    
    If ThisWorkbook.Worksheets("HYPOTHESES MORTALITE").Cells(1, 4 + NumCol).Value = Table Then
        
        TransformTable = NumCol + 4
        NumCol = NbTables + 1
    
    End If

Next NumCol

End Function

Public Function TransformSexe(Sexe As String) As Integer

If Sexe = "H" Then
    TransformSexe = 1
ElseIf Sexe = "F" Then
    TransformSexe = 2
End If

End Function
'*********************************************************
Public Function TransformTMGprev(TMGp As Double) As Integer

Dim NumTMG As Long
NumTMG = 1

Do While TMGp <> TMGprev(NumTMG)
         NumTMG = NumTMG + 1
Loop
TransformTMGprev = NumTMG

End Function
'*********************************************************


Public Function TransformNomProd(NomProd As String) As Integer

If NomProd = "AEP" Then
    TransformNomProd = 1
ElseIf NomProd = "AEP2" Then
    TransformNomProd = 2
ElseIf NomProd = "APA" Then
    TransformNomProd = 3
ElseIf NomProd = "APA2" Then
    TransformNomProd = 4
ElseIf NomProd = "APIU" Then
    TransformNomProd = 5
ElseIf NomProd = "ARA" Then
    TransformNomProd = 6
ElseIf NomProd = "CEP1" Then
    TransformNomProd = 7
ElseIf NomProd = "CFC" Then
    TransformNomProd = 8
ElseIf NomProd = "CFCP" Then
    TransformNomProd = 9
ElseIf NomProd = "CFD" Then
    TransformNomProd = 10
ElseIf NomProd = "CFDP" Then
    TransformNomProd = 11
ElseIf NomProd = "CMDV" Then
    TransformNomProd = 12
ElseIf NomProd = "COFE" Then
    TransformNomProd = 13
ElseIf NomProd = "EPA" Then
    TransformNomProd = 14
ElseIf NomProd = "EPR" Then
    TransformNomProd = 15
ElseIf NomProd = "EPR2" Then
    TransformNomProd = 16
ElseIf NomProd = "ERS" Then
    TransformNomProd = 17
ElseIf NomProd = "FIDP" Then
    TransformNomProd = 18
ElseIf NomProd = "FOFE" Then
    TransformNomProd = 19
ElseIf NomProd = "AINV" Then
    TransformNomProd = 20
ElseIf NomProd = "MAD" Then
    TransformNomProd = 21
ElseIf NomProd = "OB10" Then
    TransformNomProd = 22
ElseIf NomProd = "ORP" Then
    TransformNomProd = 23
ElseIf NomProd = "PERL" Then
    TransformNomProd = 24
ElseIf NomProd = "SPT" Then
    TransformNomProd = 25
ElseIf NomProd = "SRP" Then
    TransformNomProd = 26
ElseIf NomProd = "TRIA" Then
    TransformNomProd = 27
ElseIf NomProd = "TRIP" Then
    TransformNomProd = 28
ElseIf NomProd = "COE" Then
    TransformNomProd = 29
ElseIf NomProd = "OBEP" Then
    TransformNomProd = 30
ElseIf NomProd = "PRF2" Then
    TransformNomProd = 31
ElseIf NomProd = "PRF5" Then
    TransformNomProd = 32
ElseIf NomProd = "AEUC" Then
    TransformNomProd = 33
ElseIf NomProd = "UCAF" Then
    TransformNomProd = 34
ElseIf NomProd = "PRF1" Then
    TransformNomProd = 35
ElseIf NomProd = "PERS" Then
    TransformNomProd = 36
ElseIf NomProd = "PSVV" Then
    TransformNomProd = 37
ElseIf NomProd = "PSVC" Then
    TransformNomProd = 38
ElseIf NomProd = "PSVU" Then
    TransformNomProd = 39
ElseIf NomProd = "PSUC" Then
    TransformNomProd = 40
ElseIf NomProd = "OBER" Then
    TransformNomProd = 41
ElseIf NomProd = "AEPA" Then
    TransformNomProd = 42
ElseIf NomProd = "ALIB" Then
    TransformNomProd = 43
ElseIf NomProd = "ASVI" Then
    TransformNomProd = 44
ElseIf NomProd = "ASCA" Then
    TransformNomProd = 45
ElseIf NomProd = "ASTR" Then
    TransformNomProd = 46
ElseIf NomProd = "0" Or NomProd = "-" Or NomProd = "" Then
    TransformNomProd = 0
Else
    BDE.PresenceErreur = True
    BDE.ErreurPresente(2) = True
End If

End Function

Public Function TransformModelPoint(ModelPoint As String) As Integer

If ModelPoint = "AFI Epar" Then
    TransformModelPoint = 1
ElseIf ModelPoint = "ESCA Bon capi" Then
    TransformModelPoint = 2
ElseIf ModelPoint = "ESCA Bon capi 3.5" Then
    TransformModelPoint = 3
ElseIf ModelPoint = "ESCA Epar" Then
    TransformModelPoint = 4
ElseIf ModelPoint = "ESCA Epar 3.5" Then
    TransformModelPoint = 5
ElseIf ModelPoint = "ESCA Epar mixte" Then
    TransformModelPoint = 6
ElseIf ModelPoint = "ESCA Epar UC" Then
    TransformModelPoint = 7
ElseIf ModelPoint = "ESCA Madelin" Then
    TransformModelPoint = 8
ElseIf ModelPoint = "ESCA Rente viag" Then
    TransformModelPoint = 9
ElseIf ModelPoint = "AINV Option prévoyance" Then  '** Ce MP est créé uniquement dans la macro,
    TransformModelPoint = 10                       '** Il permet de distinguer l'option prévoyance de l'épargne du contrat AINV
ElseIf ModelPoint = "AFI COE 2016" Then
    TransformModelPoint = 11
ElseIf ModelPoint = "AFI OBEP 2016" Then
    TransformModelPoint = 12
ElseIf ModelPoint = "ESCA AINV 2016" Then
    TransformModelPoint = 13
ElseIf ModelPoint = "ESCA CFD 2016" Then
    TransformModelPoint = 14
ElseIf ModelPoint = "ESCA TRIP 2016" Then
    TransformModelPoint = 15

ElseIf ModelPoint = "0" Or ModelPoint = "-" Or ModelPoint = "" Then
    TransformModelPoint = 0
Else
    BDE.PresenceErreur = True
    BDE.ErreurPresente(3) = True
End If

End Function

Public Function InverseModelPoint(NumModelPoint As Integer) As String

If NumModelPoint = 1 Then
    InverseModelPoint = "AFI Epar"
ElseIf NumModelPoint = 2 Then
    InverseModelPoint = "ESCA Bon capi"
ElseIf NumModelPoint = 3 Then
    InverseModelPoint = "ESCA Bon capi 3.5"
ElseIf NumModelPoint = 4 Then
    InverseModelPoint = "ESCA Epar"
ElseIf NumModelPoint = 5 Then
    InverseModelPoint = "ESCA Epar 3.5"
ElseIf NumModelPoint = 6 Then
    InverseModelPoint = "ESCA Epar mixte"
ElseIf NumModelPoint = 7 Then
    InverseModelPoint = "ESCA Epar UC"
ElseIf NumModelPoint = 8 Then
    InverseModelPoint = "ESCA Madelin"
ElseIf NumModelPoint = 9 Then
    InverseModelPoint = "ESCA Rente viag"
ElseIf NumModelPoint = 10 Then                   '** Ce MP est créé uniquement dans la macro,
    InverseModelPoint = "AINV Option prévoyance" '** Il permet de distinguer l'option prévoyance de l'épargne du contrat AINV
ElseIf NumModelPoint = 11 Then
    InverseModelPoint = "AFI COE 2016"
ElseIf NumModelPoint = 12 Then
    InverseModelPoint = "AFI OBEP 2016"
ElseIf NumModelPoint = 13 Then
    InverseModelPoint = "ESCA AINV 2016"
ElseIf NumModelPoint = 14 Then
    InverseModelPoint = "ESCA CFD 2016"
ElseIf NumModelPoint = 15 Then
    InverseModelPoint = "ESCA TRIP 2016"
End If

End Function

Public Function Min(Arg1 As Double, Arg2 As Double) As Double

If Arg1 < Arg2 Then
    Min = Arg1
Else
    Min = Arg2
End If

End Function

Public Function Max(Arg1 As Double, Arg2 As Double) As Double

If Arg1 > Arg2 Then
    Max = Arg1
Else
    Max = Arg2
End If

End Function

Public Function ExisteDossier(NomDossier As String) As Boolean

ExisteDossier = Dir(NomDossier, vbSystem + vbDirectory + vbHidden) <> ""

End Function

Public Function ExisteFichier(NomFichier As String) As Boolean

ExisteFichier = Dir(NomFichier) <> ""

End Function

Public Function EstOuvert(NomFichier As String) As Boolean

EstOuvert = False
On Error Resume Next
EstOuvert = Not Workbooks(NomFichier) Is Nothing

End Function

Public Function ExisteFeuille(Classeur As String, Feuille As String) As Boolean

ExisteFeuille = False
On Error Resume Next
ExisteFeuille = Not Workbooks(Classeur).Worksheets(Feuille) Is Nothing

End Function

Public Sub creer(NomFeuil As String)

Dim Newsheet

Set Newsheet = Sheets.Add
Newsheet.Name = NomFeuil

End Sub


Public Function InverseNomProd(NumProd As Integer) As String
  
If NumProd = 1 Then
    InverseNomProd = "AEP"
ElseIf NumProd = 2 Then
    InverseNomProd = "AEP2"
ElseIf NumProd = 3 Then
    InverseNomProd = "APA"
ElseIf NumProd = 4 Then
    InverseNomProd = "APA2"
ElseIf NumProd = 5 Then
    InverseNomProd = "APIU"
ElseIf NumProd = 6 Then
    InverseNomProd = "ARA"
ElseIf NumProd = 7 Then
    InverseNomProd = "CEP1"
ElseIf NumProd = 8 Then
    InverseNomProd = "CFC"
ElseIf NumProd = 9 Then
    InverseNomProd = "CFCP"
ElseIf NumProd = 10 Then
    InverseNomProd = "CFD"
ElseIf NumProd = 11 Then
    InverseNomProd = "CFDP"
ElseIf NumProd = 12 Then
    InverseNomProd = "CMDV"
ElseIf NumProd = 13 Then
    InverseNomProd = "COFE"
ElseIf NumProd = 14 Then
    InverseNomProd = "EPA"
ElseIf NumProd = 15 Then
    InverseNomProd = "EPR"
ElseIf NumProd = 16 Then
    InverseNomProd = "EPR2"
ElseIf NumProd = 17 Then
    InverseNomProd = "ERS"
ElseIf NumProd = 18 Then
    InverseNomProd = "FIDP"
ElseIf NumProd = 19 Then
    InverseNomProd = "FOFE"
ElseIf NumProd = 20 Then
    InverseNomProd = "AINV"
ElseIf NumProd = 21 Then
    InverseNomProd = "MAD"
ElseIf NumProd = 22 Then
    InverseNomProd = "OB10"
ElseIf NumProd = 23 Then
    InverseNomProd = "ORP"
ElseIf NumProd = 24 Then
    InverseNomProd = "PERL"
ElseIf NumProd = 25 Then
    InverseNomProd = "SPT"
ElseIf NumProd = 26 Then
    InverseNomProd = "SRP"
ElseIf NumProd = 27 Then
    InverseNomProd = "TRIA"
ElseIf NumProd = 28 Then
    InverseNomProd = "TRIP"
ElseIf NumProd = 29 Then
    InverseNomProd = "COE"
ElseIf NumProd = 30 Then
    InverseNomProd = "OBEP"
ElseIf NumProd = 31 Then
    InverseNomProd = "PRF2"
ElseIf NumProd = 32 Then
    InverseNomProd = "PRF5"
ElseIf NumProd = 33 Then
    InverseNomProd = "AEUC"
ElseIf NumProd = 34 Then
    InverseNomProd = "UCAF"
ElseIf NumProd = 35 Then
    InverseNomProd = "PRF1"
ElseIf NumProd = 36 Then
    InverseNomProd = "PERS"
ElseIf NumProd = 37 Then
    InverseNomProd = "PSVV"
ElseIf NumProd = 38 Then
    InverseNomProd = "PSVC"
ElseIf NumProd = 39 Then
    InverseNomProd = "PSVU"
ElseIf NumProd = 40 Then
    InverseNomProd = "PSUC"
ElseIf NumProd = 41 Then
    InverseNomProd = "OBER"
ElseIf NumProd = 42 Then
    InverseNomProd = "AEPA"
ElseIf NumProd = 43 Then
    InverseNomProd = "ALIB"
ElseIf NumProd = 44 Then
    InverseNomProd = "ASVI"
ElseIf NumProd = 45 Then
    InverseNomProd = "ASCA"
ElseIf NumProd = 0 Then
    InverseNomProd = 0
Else
    BDE.PresenceErreur = True
    BDE.ErreurPresente(2) = True
End If

End Function
'*******************************************************************************
Public Function TransformNumChocModeling(NumChoc As Integer)
'*******************************************************************************

If NumChoc = 0 Then
    TransformNumChocModeling = "Central"
ElseIf NumChoc = 1 Then
    TransformNumChocModeling = "Mortality"
ElseIf NumChoc = 2 Then
    TransformNumChocModeling = "Longevity"
ElseIf NumChoc = 3 Then
    TransformNumChocModeling = "CAT"
ElseIf NumChoc = 4 Then
    TransformNumChocModeling = "LapseShockUp"
ElseIf NumChoc = 5 Then
    TransformNumChocModeling = "LapseShockDown"
ElseIf NumChoc = 6 Then
    TransformNumChocModeling = "LapseShockMass"
End If

End Function

'/* -----------------------------------------------------------------------
'   Optimiser la vitesse de calcul - *** F.C
'   ----------------------------------------------------------------------- */
Public Sub Optimiser_Calcul(Choix As Boolean)

With Application
    If Choix = True Then
      'Calcul manuel
      .Calculation = xlCalculationManual
      'Pas d'affichage écran
      .ScreenUpdating = False
      'Annuler message d'alerte
      .EnableEvents = False
    Else
      'Calcul automatique
      .Calculation = xlCalculationAutomatic
      'Affichage écran
      .ScreenUpdating = True
      'Annuler message d'alerte
      .EnableEvents = True
    End If
End With

End Sub
