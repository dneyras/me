'/* ---------------------------------------------------------------
'   Filtrer
'   --------------------------------------------------------------- */
Sub FiltrerAuto()
  Call Optimiser_Calcul(True)
  Range("_TABLEAU_SUIVI").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Range("_CHOIX_FILTRE")
  Call Optimiser_Calcul(False)
End Sub

'/* ---------------------------------------------------------------
'   Defiltrer
'   --------------------------------------------------------------- */
Sub DefiltrerAuto()
Dim MsgStr As Variant
  On Error GoTo Fin
  ActiveSheet.ShowAllData
  Exit Sub
Fin:
  MsgStr = MsgBox("Le filtre a déjà été supprimé  => Tableau inititial affiché !", , "Attention !")
End Sub

'/* ---------------------------------------------------------------
'   Grouper
'   --------------------------------------------------------------- */
Sub Grouper()
  ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
End Sub

'/* ---------------------------------------------------------------
'   Dissocier
'   --------------------------------------------------------------- */
Sub Dissocier()
  ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
End Sub


