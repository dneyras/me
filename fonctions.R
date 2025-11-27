

process_mp <- function(mp, annee_valorisation) {

  # Renommage rapide (sans copie si possible)
  if ("tx_charg_sur_pm_euro" %in% names(mp)) setnames(mp, "tx_charg_sur_pm_euro", "tx_charg_pm_euro")
  if ("annuites_avant_bonus1" %in% names(mp)) setnames(mp, "annuites_avant_bonus1", "annuites_avant_bonus_1")
  
  # Calculs vectorisés (plus rapide que mutate)
  # Utilisation de IDate si possible pour éviter la lourdeur POSIXct, sinon year() est ok
  mp[, `:=`(
    annee_naissance = year(date_de_naissance),
    annee_effet     = year(date_effet),
    mois_effet      = month(date_effet),
    annee_echeance  = year(date_echeance),
    nb_tetes        = 1
  )]
  
  # Logique conditionnelle rapide (fifelse est C++)
  mp[, `:=`(
    nb_tetes_uc   = fifelse(pm_uc > 0, 1, 0),
    nb_tetes_euro = fifelse(pm_euro > 0, 1, 0),
    
    # Chargements suspens
    tx_charg_pm_euro = fifelse(position_contrat == "En réduction" & tx_charg_pm_euro > 0, 
                               tx_charg_pm_euro + tx_charg_suspens_pp, tx_charg_pm_euro),
    tx_charg_pm_uc   = fifelse(position_contrat == "En réduction" & tx_charg_pm_uc > 0, 
                               tx_charg_pm_uc + tx_charg_suspens_pp, tx_charg_pm_uc),
    
    # Types
    contrat_prorogeable = as.logical(contrat_prorogeable),
    indic_obseque       = as.logical(indic_obseque)
  )]
  
  # Calcul des durées (vectorisé)
  offset_annee <- annee_valorisation - mp$annee_effet + 1
  mp[, `:=`(
    annuites_avant_bonus_1 = annuites_avant_bonus_1 - offset_annee,
    annuites_avant_bonus_2 = annuites_avant_bonus_2 - offset_annee,
    duree_versement        = duree_versement - offset_annee
  )]
  
  # Calcul ajustement et délais finaux
  # fcase est l'équivalent optimisé data.table de case_when
  mp[, ajustement_delai := fcase(
    periodicite == "Semestriel" & mois_effet > 6, 1,
    periodicite == "Trimestriel" & mois_effet > 3, 1,
    default = 0
  )]
  
  mp[, `:=`(
    delai_bonus_1 = annuites_avant_bonus_1 + ajustement_delai,
    delai_bonus_2 = annuites_avant_bonus_2 + ajustement_delai,
    duree_restante_prime = duree_versement + ajustement_delai
  )]
  
  # Suppression colonnes inutiles (par référence)
  cols_drop <- c("annuites_avant_bonus_1", "annuites_avant_bonus_2", "duree_versement", "ajustement_delai")
  # On vérifie qu'elles existent avant de supprimer pour éviter erreur
  cols_drop <- intersect(names(mp), cols_drop)
  if(length(cols_drop) > 0) mp[, (cols_drop) := NULL]
  
  return(mp)
}



proj_init_dt_optim <- function(mp, date_valorisation, annee_valorisation){
  
  # Créer la table des années (51 lignes)
  annees_proj <- data.table(annee_projection = 0:50)
  
  # !! OPTIMISATION !!
  # Calculer les 51 dates de projection ICI, sur la petite table
  annees_proj[, date_projection := date_valorisation + dyears(annee_projection)]
  
  # Ajouter la clé de jointure
  annees_proj[, temp_id := 1L]
  
  # --- Étape 1 : Cross Join ---
  # (Pensez à filtrer les colonnes de 'mp' en amont si possible)
  mp_proj <- mp[, temp_id := 1L]
  mp_proj <- mp_proj[annees_proj, on = "temp_id", allow.cartesian = TRUE]
  mp_proj[, temp_id := NULL]

  # Calculer les variables dépendantes en une seule passe
  mp_proj[, `:=`(
    age_assure = annee_valorisation - annee_naissance + annee_projection,
    anciennete_contrat = annee_valorisation - annee_effet + annee_projection,
    reste_contrat = annee_echeance - annee_valorisation - annee_projection,
    indic_terme_contrat = date_projection >= date_echeance & date_projection < date_echeance + dyears(1),
    date_anniversaire = calculer_date_anniversaire(date_effet, date_projection))] # VIGILANCE ICI
    
  mp_proj[, `:=`(
    # Optimisation : réutilisation de 'reste_contrat'
    # (Vérifiez si la logique '... + 1' est correcte vs 'year(date_echeance) - year(date_projection) + 1')
    duree_restante = pmax(reste_contrat + 1, 0), 
    
    duree_restante_prime = pmax(duree_restante_prime - annee_projection + 1, 0),
    delai_bonus_1 = delai_bonus_1 - annee_projection,
    delai_bonus_2 = delai_bonus_2 - annee_projection
  )]
  
  # --- Étape 3 : Second Mutate (fortement optimisé) ---
  
  # Identifier les groupes (clés uniques) qui ont un delai_bonus_1 == 0
  groupes_bonus_1 <- unique(mp_proj[delai_bonus_1 == 0, .(provenance, numero)])
  
  # Appliquer pmax(..., 0) uniquement à ces groupes via une "update-join"
  mp_proj[groupes_bonus_1, on = .(provenance, numero), delai_bonus_1 := pmax(delai_bonus_1, 0)]
  
  # Idem pour delai_bonus_2
  groupes_bonus_2 <- unique(mp_proj[delai_bonus_2 == 0, .(provenance, numero)])
  mp_proj[groupes_bonus_2, on = .(provenance, numero), delai_bonus_2 := pmax(delai_bonus_2, 0)]
  
  return(mp_proj)
}


join_assumptions_dt_fast <- function(mp, 
                                     hypotheses_morta, 
                                     hypotheses_produits, 
                                     hypotheses_rendement, 
                                     choc_mortalite,
                                     choc_longevite,
                                     choc_catastrophe, 
                                     choc
                                     ){
  
  # Assurer que les tables d'hypothèses sont des data.tables
  setDT(hypotheses_morta)
  setDT(hypotheses_produits)
  setDT(hypotheses_rendement)
  
  hypotheses_morta <- copy(hypotheses_morta)
  hypotheses_produits <- copy(hypotheses_produits)
  hypotheses_rendement <- copy(hypotheses_rendement)
  
  # choc mortalité 
  if (choc_mortalite){
    hypotheses_morta[, qx_approx := fifelse(qx_approx == 1, 1, pmin(qx_approx*(1 + choc$mortalite), 1))]
  }
  
  #choc longevite 
  if (choc_longevite){
    hypotheses_morta[, qx_approx := fifelse(qx_approx == 1, 1, pmin(qx_approx*(1 + choc$longevite), 1))]
  }
  
  
  # --- 1. Préparation de la table de mortalité ---
  hypo_morta_subset <- hypotheses_morta[, .(sexe, age, qx_approx)]
  
  cols <- setdiff(names(hypo_morta_subset), c("age", "sexe"))
  mp[, (cols) := 
       hypo_morta_subset[.SD, on = .(age = age_assure, sexe = sexe_client), .SD, .SDcols = cols]]
  
  cols <- setdiff(names(hypotheses_produits), c("produit", "anciennete"))
  mp[, (cols) := 
       hypotheses_produits[.SD, on = .(produit = nom_produit, anciennete = anciennete_contrat), .SD, .SDcols = cols]]
  
  cols <- setdiff(names(hypotheses_rendement), "annee_projection")
  mp[, (cols) := 
       hypotheses_rendement[.SD, on = .(annee_projection = annee_projection), .SD, .SDcols = cols]]
  
  # --- Chocs (déjà optimisés) ---
  if (choc_catastrophe) {
    mp[annee_projection == 1, 
       qx_approx := pmin(qx_approx + choc$catastrophe, 1)]
  }
  
  return(mp)
}




join_assumptions_rachat <- function(mp, 
                                     hypotheses_rachat, 
                                     choc_rachat_hausse,
                                     choc_rachat_baisse,
                                     choc_rachat_massif,
                                     choc,
                                     plancher_baisse_rachat = 0.2
){
  
  
  #choc rachat hausse
  if (choc_rachat_hausse){
    hypotheses_rachat <- hypotheses_rachat |> 
      mutate(across(c(tx_rachat_part, tx_rachat_tot), \(x) pmin(x * (1 + choc$rachat_hausse), 1)))
  }
  
  # choc rachat baisse
  if (choc_rachat_baisse){
    hypotheses_rachat <- hypotheses_rachat |> 
      mutate(across(c(tx_rachat_part, tx_rachat_tot), \(x) pmax(x * (1 + choc$rachat_baisse), x - plancher_baisse_rachat)))
  }
  
  
  setDT(hypotheses_rachat)
  
  cols <- setdiff(names(hypotheses_rachat), c("anciennete", "categorie_rachat"))
  mp[, (cols) := 
       hypotheses_rachat[.SD, on = .(categorie_rachat = categorie_rachat, anciennete = anciennete_contrat), .SD, .SDcols = cols]]
  
  
  if (choc_rachat_massif) {
    mp[annee_projection == 1, 
       `:=`(
         tx_rachat_tot = choc$rachat_massif,
         tx_rachat_part = 0
       )]
  }
  
  return(mp)
}




calcul_prob_dt <- function(mp){
  
  # S'assurer que mp est un data.table
  setDT(mp)
  
  # --- Étape 1 : Calcul des probabilités indépendantes ---
  # Calcule toutes les probabilités qui ne dépendent que des données initiales.
  mp[, `:=`(
    p_deces = pmax(0, pmin(1, qx_approx)),
    p_tirage = pmin(tx_tirage, pmax(0, 1 - qx_approx)),
    p_rachat_tot = pmin(tx_rachat_tot, pmax(0, 1 - qx_approx - tx_tirage)),
    p_rachat_part = pmin(tx_rachat_part, pmax(0, 1 - qx_approx - tx_tirage - tx_rachat_tot))
  )]
  
  # --- Étape 2 : Calcul des probabilités dépendantes ---
  # p_terme dépend des calculs de l'étape 1
  mp[, p_terme := indic_terme_contrat * (1 - p_deces - p_tirage - p_rachat_tot)]
  
  # p_clot dépend de p_terme
  mp[, p_clot := 1 - p_deces - p_tirage - p_rachat_tot - p_terme]
  
  # --- Étape 3 : Calcul de la probabilité cumulée (fonction fenêtre) ---
  # Équivalent du deuxième mutate()
  
  # Remplacement des NA (plus rapide avec data.table)
  mp[is.na(p_clot), p_clot := 0]
  
  # Application de la fonction fenêtre 'cumprod' par groupe
  # C'est l'équivalent de .by = c(provenance, numero)
  mp[, p_clot := c(1, cumprod(p_clot[-1])), by = .(provenance, numero)]
  
  # --- Étape 4 : Calculs finaux ---
  # Équivalent du troisième mutate()
  mp[, `:=`(
    nb_cloture = p_clot * nb_tetes,
    nb_cloture_uc = p_clot * nb_tetes_uc,
    nb_cloture_euro = p_clot * nb_tetes_euro
  )]
  
  # La fonction dplyr retournait mp, on fait de même
  return(mp)
}


calcul_prime_dt <- function(mp){
  
  setorderv(mp, c("provenance", "numero", "annee_projection")) 
  
  
  # Pré-calculer les coefficients UNE SEULE FOIS au lieu de 2 fois
  mp[, coeff_temp := coeff_prime_bonus_math(mois_effet, periodicite)]
  
  # 1. Initialiser les colonnes à 0 (très rapide)
  mp[, `:=`(
    coeff_prime_annee_echeance = 0,
    coeff_prime_annee_bonus_temp = 0
  )]
  
  # 2. Calcul de l'échéance (Mise à jour par référence sur sous-ensemble)
  # On évite le fifelse en ne sélectionnant que les lignes où la condition est vraie
  mp[duree_restante_prime == 1, coeff_prime_annee_echeance := coeff_temp]
  
  # 3. Calcul du bonus avec Vectorisation Globale
  # On calcule les décalages (lags) pour tout le tableau d'un coup
  # On vérifie que nous sommes toujours sur le même contrat pour ne pas décaler d'un contrat A vers B
  mp[, `:=`(
    lag_d1 = shift(delai_bonus_1, type = "lag", fill = 0),
    lag_d2 = shift(delai_bonus_2, type = "lag", fill = 0),
    same_contract = (numero == shift(numero, type = "lag", fill = -1)) # Protection frontière
  )]
  
  # Application de la logique simplifiée : (A=1 & B=0) | (A=0 & B=1) équivaut à A + B == 1
  # On ajoute la condition 'same_contract' pour s'assurer qu'on ne prend pas la valeur du contrat précédent
  mp[same_contract == TRUE & (lag_d1 + lag_d2 == 1), 
     coeff_prime_annee_bonus_temp := coeff_temp]
  
  # Nettoyage des colonnes temporaires (optionnel)
  mp[, `:=`(lag_d1 = NULL, lag_d2 = NULL, same_contract = NULL)]
  
  # Pré-calculer les ratios communs
  mp[, ratio_cloture := nb_cloture / nb_tetes]
  
  # Tout en une seule passe avec chaînage
  mp[, coeff_sans_primes := fcase(
    duree_restante == 1, coeff_prime_annee_echeance,
    delai_bonus_1 == 0 | delai_bonus_2 == 0, coeff_prime_annee_bonus_temp,
    default = 0
  )][, `:=`(
    prime_comm_euro = fcase(
      (delai_bonus_1 < 0 & delai_bonus_2 < 0), 0,
      (duree_restante_prime > 1 & (delai_bonus_1 > 0 | delai_bonus_2 > 0)), 
      prime_commerciale_euro * ratio_cloture,
      default = prime_commerciale_euro * ratio_cloture * coeff_sans_primes
    ),
    prime_comm_uc = fcase(
      (delai_bonus_1 < 0 & delai_bonus_2 < 0), 0,
      (duree_restante_prime > 1 & (delai_bonus_1 > 0 | delai_bonus_2 > 0)), 
      prime_commerciale_uc * ratio_cloture,
      default = prime_commerciale_uc * ratio_cloture * coeff_sans_primes
    )
  )][, `:=`(
    chargement_prime_euro = prime_comm_euro * tx_charg_prime_euro,
    chargement_prime_uc = prime_comm_uc * tx_charg_prime_uc,
    commissions_primes_euro = prime_comm_euro * taux_com_prime_euro,
    commissions_primes_uc = prime_comm_uc * taux_com_prime_uc
  )][, `:=`(
    prime_nette_euro = prime_comm_euro * (1 - tx_charg_prime_euro), # Évite une soustraction
    prime_nette_uc = prime_comm_uc * (1 - tx_charg_prime_uc)
  )][, `:=`(
    interets_prime_euro = prime_nette_euro * ((1 + taux_annuel_garanti)^0.5 - 1), 
    pb_prime_euro = prime_nette_euro * ((1 + pb)^0.5 - 1),
    rendement_uc_prime = prime_nette_uc * ((1 + rendement_uc)^0.5 - 1)
  )]
  
  # Nettoyage
  mp[, c("coeff_prime_annee_bonus_temp", "ratio_cloture") := NULL]
  
  return(mp)
}


# --- Fonction utilitaire pour la date anniversaire ---
# Renommée pour plus de clarté
calculer_date_anniversaire <- function(date_effet, date_valorisation_annee_proj){
  # Décale l'année de la date d'effet à l'année de la date de projection
  year(date_effet) <- year(date_valorisation_annee_proj)
  return(date_effet)
}


# il faut utiliser cette fonction 


f_coeff <- function(mois, periode){
  
  nb_max <- 12 %/% periode 
  
  res <- ceiling((12 - mois + 1) / periode)
  
  return((nb_max - res)/nb_max)
  
}



coeff_prime_bonus <- function(mois, periode){
  
  case_match(periode,
             
             "Unique" ~ 0,
             "Annuel" ~ 1 - f_coeff(mois, 12),
             "Semestriel" ~ 1 - f_coeff(mois, 6),
             "Trimestriel" ~ 1 - f_coeff(mois, 3),
             "Mensuel" ~ mois/12
  )
  
}



# VERSION INTERMÉDIAIRE : simplification mathématique
coeff_prime_bonus_math <- function(mois, periode) {
  fcase(
    periode == "Unique", 0,
    periode == "Mensuel", mois / 12,
    periode == "Annuel", (mois - 1) / 12,  # Simplifié !
    periode == "Semestriel", {
      # nb_max = 2, periode = 6
      res <- ceiling((13 - mois) / 6)
      (2 - res) / 2
    },
    periode == "Trimestriel", {
      # nb_max = 4, periode = 3
      res <- ceiling((13 - mois) / 3)
      (4 - res) / 4
    },
    default = NA_real_
  )
}

calcul_prime_avec <- function(duree_restante_prime, coeff_prime_annee, prime_com, nb_cloture, nb_tetes){
  
  case_when(
    duree_restante_prime > 1 ~ prime_com * nb_cloture / nb_tetes,
    duree_restante_prime == 1 ~ prime_com * nb_cloture / nb_tetes * coeff_prime_annee,
    .default = 0
  )
  
}



calcul_prime_sans <- function(duree_restante_prime, delai_bonus, coeff_prime_bonus, coeff_prime_annee, prime_com, nb_cloture, nb_tetes){
  
  if_else(delai_bonus > -1, 
          
          case_when(
            duree_restante_prime > 1 & delai_bonus > 0 ~ prime_com * nb_cloture / nb_tetes,
            duree_restante_prime == 1 ~ prime_com * nb_cloture / nb_tetes * coeff_prime_annee,
            delai_bonus == 0 ~ prime_com * nb_cloture / nb_tetes * coeff_prime_bonus,
            .default = 0
          ),
          
          0)
}

generer_flexings <- function(resultats_proj, date_valorisation) {
  
  # Flexing Euro
  flexing_euro <- resultats_proj %>%
    # --- GROUPEMENT CORRECT ---
    group_by(annee_projection, taux_annuel_garanti) %>% # Utiliser la colonne tmg calculée
    summarise(
      # --- Calculs intermédiaires pour les sommes ---
      # Gardez ceux-ci si vous les utilisez, sinon calculez directement
      sum_nb_cloture_euro = sum(nb_cloture_euro, na.rm = TRUE),
      sum_prime_comm_euro = if_else(annee_projection[1] == 0, # Utiliser annee_projection[1] car summarise attend une seule valeur
                                    sum(prime_commerciale_euro, na.rm = TRUE),
                                    sum(prime_comm_euro, na.rm = TRUE)),
      sum_pm_euro = sum(pm_euro, na.rm = TRUE), # pm_euro doit être la PM FINALE
      sum_sin_deces_euro = sum(sin_deces_euro, na.rm = TRUE),
      sum_charg_deces_euro = sum(charg_deces_euro, na.rm = TRUE),
      sum_sin_rachat_tot_euro = sum(sin_rachat_tot_euro, na.rm = TRUE),
      sum_charg_rachat_tot_euro = sum(charg_rachat_tot_euro, na.rm = TRUE),
      sum_sin_rachat_part_euro = sum(sin_rachat_part_euro, na.rm = TRUE),
      sum_charg_rachat_part_euro = sum(charg_rachat_part_euro, na.rm = TRUE),
      sum_sin_tirage_euro = sum(sin_tirage_euro, na.rm = TRUE),
      sum_charg_tirage_euro = sum(charg_tirage_euro, na.rm = TRUE),
      sum_sin_terme_euro = sum(sin_terme_euro, na.rm = TRUE),
      sum_charg_terme_euro = sum(charg_terme_euro, na.rm = TRUE),
      sum_int_pm_mi_euro = sum(interets_pm_mi_periode_euro, na.rm = TRUE),
      sum_int_fin_euro = sum(interets_fin_periode_euro, na.rm = TRUE),
      # sum_pb_mi_euro = sum(pb_mi_periode_euro, na.rm = TRUE), # Non utilisé dans PS=0
      # sum_pb_fin_euro = sum(pb_fin_periode_euro, na.rm = TRUE), # Non utilisé dans PS=0
      sum_charg_prime_euro = if_else(annee_projection[1] == 0, 0, sum(chargement_prime_euro, na.rm = TRUE)),
      sum_charg_pm_euro = sum(charg_pm_euro, na.rm = TRUE),
      sum_comm_prime_euro = if_else(annee_projection[1] == 0, 0, sum(commissions_primes_euro, na.rm = TRUE)),
      sum_comm_pm_euro = sum(commissions_pm_euro, na.rm = TRUE),
      # Calcul conditionnel pour Comm_Autres et Indemnites_Encours
      sum_comm_pm_euro_cond = sum(if_else(charg_pm_euro == 0 & commissions_pm_euro != 0, commissions_pm_euro, 0), na.rm = TRUE),
      sum_comm_pm_euro_indemn = sum(if_else(charg_pm_euro != 0 | commissions_pm_euro == 0, commissions_pm_euro, 0), na.rm = TRUE),
      
      # --- Assignation aux colonnes finales (ordre VBA) ---
      # Col 3: LoB
      LoB = 30,
      # Col 4: TMG (déjà dans le group_by)
      # Col 5: RachDyn
      RachDyn = 1,
      # Col 6: Code_support
      Code_support = "Euro",
      # Col 7: Effectifs
      Effectifs = sum_nb_cloture_euro,
      # Col 8: Presta_Deces
      Presta_Deces = sum_sin_deces_euro + sum_charg_deces_euro,
      # Col 9: Presta_Rachat
      Presta_Rachat = sum_sin_rachat_tot_euro + sum_sin_rachat_part_euro + sum_charg_rachat_tot_euro + sum_charg_rachat_part_euro,
      # Col 10: Presta_Autres (Tirage + Terme)
      Presta_Autres = (sum_sin_tirage_euro + sum_charg_tirage_euro) + (sum_sin_terme_euro + sum_charg_terme_euro),
      # Col 11: Cotisations
      Cotisations = sum_prime_comm_euro,
      # Col 12: PM
      PM = sum_pm_euro,
      # Col 13: IT
      IT = sum_int_pm_mi_euro + sum_int_fin_euro,
      # Col 14: PS (Mis à 0 comme dans le flexing VBA Euro)
      PS = 0,
      # Col 15 à 26: Zéros
      Col15 = 0, Col16 = 0, Col17 = 0, Col18 = 0, Col19 = 0, Col20 = 0,
      Col21 = 0, Col22 = 0, Col23 = 0, Col24 = 0, Col25 = 0, Col26 = 0,
      # Col 27: Chgt_Cotis
      Chgt_Cotis = sum_charg_prime_euro,
      # Col 28: Chgt_Encours
      Chgt_Encours = sum_charg_pm_euro,
      # Col 29: Chgt_Presta
      Chgt_Presta = sum_charg_deces_euro + sum_charg_rachat_tot_euro + sum_charg_rachat_part_euro + sum_charg_terme_euro + sum_charg_tirage_euro,
      # Col 30: Chgt_Autres (Mis à 0 comme dans le flexing VBA Euro)
      Chgt_Autres = 0,
      # Col 31: Frais_Cotis (Mis à 0 comme dans le flexing VBA Euro)
      Frais_Cotis = 0,
      # Col 32: Frais_Encours (Mis à 0 comme dans le flexing VBA Euro)
      Frais_Encours = 0,
      # Col 33: Frais_Presta (Mis à 0 comme dans le flexing VBA Euro)
      Frais_Presta = 0,
      # Col 34: Frais_Autres (Mis à 0 comme dans le flexing VBA Euro)
      Frais_Autres = 0,
      # Col 35: Indemnites_Cotis
      Indemnites_Cotis = sum_comm_prime_euro,
      # Col 36: Indemnites_Encours (Ajusté pour Comm_Autres)
      Indemnites_Encours = sum_comm_pm_euro_indemn,
      # Col 37: Indemnites_Presta (Mis à 0 comme dans le flexing VBA Euro)
      Indemnites_Presta = 0,
      # Col 38: Comm_Autres (Conditionnel)
      Comm_Autres = sum_comm_pm_euro_cond,
      # Col 39: RetroAE (Mis à 0 comme dans le flexing VBA Euro)
      RetroAE = 0,
      
      .groups = 'drop' # Important de garder ça
    ) %>%
    # --- Ajout de la colonne Année (Col 2) ---
    mutate(
      Annee = annee_projection + year(date_valorisation) # Assurez-vous que AnneeInitiale existe
    ) %>%
    # --- Sélection et Réorganisation finale des colonnes ---
    select(
      Annee,                 # Col 2
      LoB,                   # Col 3
      TMG = taux_annuel_garanti,                   # Col 4 (Nom à ajuster si besoin)
      RachDyn,               # Col 5
      Code_support,          # Col 6
      Effectifs,             # Col 7
      Presta_Deces,          # Col 8
      Presta_Rachat,         # Col 9
      Presta_Autres,         # Col 10
      Cotisations,           # Col 11
      PM,                    # Col 12
      IT,                    # Col 13
      PS,                    # Col 14
      Col15, Col16, Col17, Col18, Col19, Col20, # Cols 15-20
      Col21, Col22, Col23, Col24, Col25, Col26, # Cols 21-26
      Chgt_Cotis,            # Col 27
      Chgt_Encours,          # Col 28
      Chgt_Presta,           # Col 29
      Chgt_Autres,           # Col 30
      Frais_Cotis,           # Col 31
      Frais_Encours,         # Col 32
      Frais_Presta,          # Col 33
      Frais_Autres,          # Col 34
      Indemnites_Cotis,      # Col 35
      Indemnites_Encours,    # Col 36
      Indemnites_Presta,     # Col 37
      Comm_Autres,           # Col 38
      RetroAE                # Col 39
    ) %>%
    # Arrondir les résultats comme en VBA si nécessaire (optionnel)
    mutate(across(all_of(6:38), ~ round(., 2))) |> 
    arrange(TMG, Annee)
  
  # Flexing UC
  flexing_uc <- resultats_proj %>%
    # --- GROUPEMENT PAR ANNÉE ---
    group_by(annee_projection) %>%
    summarise(
      # --- Calculs intermédiaires pour les sommes ---
      sum_nb_cloture_uc = sum(nb_cloture_uc, na.rm = TRUE),
      sum_prime_comm_uc = if_else(annee_projection[1] == 0,
                                  sum(prime_commerciale_uc, na.rm = TRUE), # Prime initiale
                                  sum(prime_comm_uc, na.rm = TRUE)), # Prime projetée
      sum_pm_uc = sum(pm_uc, na.rm = TRUE), # PM finale de l'année
      sum_sin_deces_uc = sum(sin_deces_uc, na.rm = TRUE),
      sum_charg_deces_uc = sum(charg_deces_uc, na.rm = TRUE),
      sum_sin_rachat_tot_uc = sum(sin_rachat_tot_uc, na.rm = TRUE),
      sum_charg_rachat_tot_uc = sum(charg_rachat_tot_uc, na.rm = TRUE),
      sum_sin_rachat_part_uc = sum(sin_rachat_part_uc, na.rm = TRUE),
      sum_charg_rachat_part_uc = sum(charg_rachat_part_uc, na.rm = TRUE),
      sum_sin_tirage_uc = sum(sin_tirage_uc, na.rm = TRUE),
      sum_charg_tirage_uc = sum(charg_tirage_uc, na.rm = TRUE),
      sum_sin_terme_uc = sum(sin_terme_uc, na.rm = TRUE),
      sum_charg_terme_uc = sum(charg_terme_uc, na.rm = TRUE),
      # ATTENTION: VBA somme les Interets (mis à 0), pas les rendements pour IT
      # Pour correspondre au VBA :
      sum_it_vba = sum(interets_pm_mi_periode_uc + interets_fin_periode_uc, na.rm = TRUE), # Devrait être 0
      # Pour la justesse financière (à utiliser si besoin mais ne correspondra pas au VBA):
      # sum_rendements = sum(rend_uc_mi_periode_uc + rend_fin_periode_uc, na.rm = TRUE),
      sum_charg_prime_uc = if_else(annee_projection[1] == 0, 0, sum(chargement_prime_uc, na.rm = TRUE)),
      sum_charg_pm_uc = sum(charg_pm_uc, na.rm = TRUE),
      sum_retro_global_pm_uc = sum(retro_global_pm_uc, na.rm = TRUE),
      sum_retro_aepm_uc = sum(retro_aepm_uc, na.rm = TRUE),
      sum_comm_prime_uc = if_else(annee_projection[1] == 0, 0, sum(commissions_primes_uc, na.rm = TRUE)),
      sum_comm_pm_uc = sum(commissions_pm_uc, na.rm = TRUE),
      # Calcul conditionnel pour Comm_Autres et Indemnites_Encours UC
      sum_comm_pm_uc_cond = sum(if_else(charg_pm_uc == 0 & commissions_pm_uc != 0, commissions_pm_uc, 0), na.rm = TRUE),
      sum_comm_pm_uc_indemn = sum(if_else(charg_pm_uc != 0 | commissions_pm_uc == 0, commissions_pm_uc, 0), na.rm = TRUE),
      
      # --- Assignation aux colonnes finales (ordre VBA) ---
      # Col 3: LoB
      LoB = 31,
      # Col 4: TMG
      TMG = 0,
      # Col 5: RachDyn
      RachDyn = 1,
      # Col 6: Code_support
      Code_support = "UC1",
      # Col 7: Effectifs
      Effectifs = sum_nb_cloture_uc,
      # Col 8: Presta_Deces
      Presta_Deces = sum_sin_deces_uc + sum_charg_deces_uc,
      # Col 9: Presta_Rachat
      Presta_Rachat = sum_sin_rachat_tot_uc + sum_sin_rachat_part_uc + sum_charg_rachat_tot_uc + sum_charg_rachat_part_uc,
      # Col 10: Presta_Autres (Tirage + Terme)
      Presta_Autres = (sum_sin_tirage_uc + sum_charg_tirage_uc) + (sum_sin_terme_uc + sum_charg_terme_uc),
      # Col 11: Cotisations
      Cotisations = sum_prime_comm_uc,
      # Col 12: PM
      PM = sum_pm_uc,
      # Col 13: IT (Correspondance VBA = 0)
      IT = sum_it_vba, # Utilisez sum_rendements si vous préférez la justesse financière
      # Col 14: PS
      PS = 0, # Explicitement 0 dans le VBA pour UC
      # Col 15 à 26: Zéros
      Col15 = 0, Col16 = 0, Col17 = 0, Col18 = 0, Col19 = 0, Col20 = 0,
      Col21 = 0, Col22 = 0, Col23 = 0, Col24 = 0, Col25 = 0, Col26 = 0,
      # Col 27: Chgt_Cotis
      Chgt_Cotis = sum_charg_prime_uc,
      # Col 28: Chgt_Encours
      Chgt_Encours = sum_charg_pm_uc + sum_retro_global_pm_uc - sum_retro_aepm_uc,
      # Col 29: Chgt_Presta
      Chgt_Presta = sum_charg_deces_uc + sum_charg_rachat_tot_uc + sum_charg_rachat_part_uc+ sum_charg_terme_uc + sum_charg_tirage_uc,
      # Col 30: Chgt_Autres (Mis à 0 comme dans le flexing VBA UC)
      Chgt_Autres = 0,
      # Col 31: Frais_Cotis (Mis à 0 comme dans le flexing VBA UC)
      Frais_Cotis = 0,
      # Col 32: Frais_Encours (Mis à 0 comme dans le flexing VBA UC)
      Frais_Encours = 0,
      # Col 33: Frais_Presta (Mis à 0 comme dans le flexing VBA UC)
      Frais_Presta = 0,
      # Col 34: Frais_Autres (Mis à 0 comme dans le flexing VBA UC)
      Frais_Autres = 0,
      # Col 35: Indemnites_Cotis
      Indemnites_Cotis = sum_comm_prime_uc,
      # Col 36: Indemnites_Encours (Ajusté pour Comm_Autres + Rétros non-AE)
      Indemnites_Encours = sum_comm_pm_uc_indemn + (sum_retro_global_pm_uc - sum_retro_aepm_uc),
      # Col 37: Indemnites_Presta (Mis à 0 comme dans le flexing VBA UC)
      Indemnites_Presta = 0,
      # Col 38: Comm_Autres (Conditionnel)
      Comm_Autres = sum_comm_pm_uc_cond,
      # Col 39: RetroAE
      RetroAE = sum_retro_aepm_uc,
      
      .groups = 'drop' # Important
    ) %>%
    # --- Ajout de la colonne Année (Col 2) ---
    mutate(
      Annee = annee_projection + year(date_valorisation)
    ) %>%
    # --- Sélection et Réorganisation finale des colonnes ---
    select(
      Annee,                 # Col 2
      LoB,                   # Col 3
      TMG,                   # Col 4
      RachDyn,               # Col 5
      Code_support,          # Col 6
      Effectifs,             # Col 7
      Presta_Deces,          # Col 8
      Presta_Rachat,         # Col 9
      Presta_Autres,         # Col 10
      Cotisations,           # Col 11
      PM,                    # Col 12
      IT,                    # Col 13
      PS,                    # Col 14
      Col15, Col16, Col17, Col18, Col19, Col20, # Cols 15-20
      Col21, Col22, Col23, Col24, Col25, Col26, # Cols 21-26
      Chgt_Cotis,            # Col 27
      Chgt_Encours,          # Col 28
      Chgt_Presta,           # Col 29
      Chgt_Autres,           # Col 30
      Frais_Cotis,           # Col 31
      Frais_Encours,         # Col 32
      Frais_Presta,          # Col 33
      Frais_Autres,          # Col 34
      Indemnites_Cotis,      # Col 35
      Indemnites_Encours,    # Col 36
      Indemnites_Presta,     # Col 37
      Comm_Autres,           # Col 38
      RetroAE                # Col 39
    ) %>%
    # Arrondir les résultats comme en VBA si nécessaire (optionnel)
    mutate(across(all_of(6:38), ~ round(., 2))) |> 
    arrange(Annee)
  
  res <- bind_rows(flexing_euro, flexing_uc) |> 
    ungroup() |> 
    mutate(LINE = 1:n(), .before = 1)
  
  colnames(res) <- c("LINE",	"Annee_proj",	"LoB",	"TMG",	"Indic_Rachat_Dyn",	"Code_Support",	"Effectifs",	"Prestations_deces",	"Prestations_rachat",	"Prestations_Autres",	"Cotisations",	"PM_Comptable_Euro",	"IT",	"Prestations_deces_C",	"Prestations_rachat_C",	"Prestations_Autres_C",	"Cotisations_C",	"PM_Comptable_C",	"IT_C",	"Arb_vers_Euro",	"Arb_vers_UC1",	"Arb_vers_UC2",	"Arb_vers_UC3",	"Arb_vers_UC4",	"Arb_vers_UC5",	"Arb_vers_UC6",	"Chgmt_sur_cotisations",	"Chgmt_sur_encours",	"Chgmt_sur_prestations",	"Chgmt_Autres",	"Frais_sur_cotisations",	"Frais_sur_encours",	"Frais_sur_prestations",	"Frais_Autres",	"Commissions_cotisations",	"Commissions_encours",	"Commissions_prestations",	"Commissions_Autres",	"Retrocommissions")
  
  return(res)
}



missing_column <- function(data, base){
  
  con <- DBI::dbConnect(odbc::odbc(),
                        Driver = "SQL Server",
                        Server = "sql-prod-bi.groupe.intra\\sql1",
                        Database = base, timeout = Inf)
  
  
  Versions <- tbl(con, "Versions")
  
  Personne <- tbl(con, "Personne")
  
  Contrats <- tbl(con, "Contrats")
  
  
  distinct_num_adh <- data |> 
    distinct(num_adh)
  
  inline_num_adh <- dbplyr::copy_inline(con, distinct_num_adh)
  
  contrats_epargne <- Contrats |>
    select(NumAdh, Contrat, CodAss) |>
    inner_join(Personne |> select(CodPers, QualCiv), by = join_by(CodAss == CodPers)) |>
    inner_join(Versions |> filter(ClasseV == 'EPARGN') |> select(Contrat, Typrod, BrancheAna, slogan), by = "Contrat") |>
    inner_join(inline_num_adh, by = join_by(NumAdh == num_adh)) |> 
    collect() |> 
    mutate(
      type_entite = case_match(QualCiv,
                               1 ~ "H",
                               c(2, 3) ~ "F",
                               4 ~ "personne morale",
                               5 ~ "couple"),
      support = if_else(Typrod == "UNITES", "multisupport", "monosupport"),
      type_contrat = if_else(str_detect(slogan, "assurance") & str_detect(slogan, "vie"), "assurance vie", "capitalisation")
    ) |> 
    select(
      num_adh = NumAdh,
      type_entite,
      support,
      type_contrat,
      mode_versement = BrancheAna
    )
  
  
  return(contrats_epargne)
  
  
}


hypotheses_rachat_new <- function(mp, final_multinom, final_beta, choc, choc_rachat_hausse, choc_rachat_baisse, plancher_baisse_rachat = 0.2){
  
  
  infos_stras <- missing_column(mp, base = "AssDev")
  
  
  infos_lille <- missing_column(mp, base = "Afi_Prod")
  
  
  new_col <- bind_rows(infos_stras, infos_lille)
  
  mp <- mp |> 
    left_join(new_col, by = "num_adh", multiple = "first")
  
  
  mp <- mp |> 
    mutate(
      indic_obseque = as.numeric(indic_obseque),
      type_entite = replace_na(type_entite, "H"),
      type_entite = if_else(type_entite %in% c("couple", "personne morale"), "H", type_entite),
      support = replace_na(support, "multisupport"),
      type_contrat = replace_na(type_contrat, "assurance vie"),
      mode_versement = replace_na(mode_versement, "PP "),
      rachat_annee_precedente = FALSE,
    ) |> 
    mutate(across(c(type_entite, support, type_contrat, mode_versement, indic_obseque, rachat_annee_precedente), as_factor)) |> 
    mutate(
      pm_cut = cut(pm_totale,
                   breaks = c(0, 2500, 5000, 10000, 100000, Inf),
                   labels = c("[0, 2500[", "[2500, 5000[", "[5000, 10000[", "[10000, 100000[", "[100000, Inf["),
                   right = F
      ),
      age_cut = cut(age_assure,
                    breaks = c(0, 30, 40, 50, 60, 70, 80, 90, Inf),
                    labels = c("[0, 30[", "[30, 40[", "[40, 50[", "[50, 60[", "[60, 70[", "[70, 80[", "[80, 90[", "[90, Inf["),
                    right = F
      ),
      anciennete_contrat_cut = cut(anciennete_contrat,
                                   breaks = c(0, 4, 8, Inf),
                                   labels = c("[0, 4[", "[4, 8[", "[8, Inf["),
                                   right = F
      )
    ) |> 
    mutate(
      id_adherent = NA,
      annee_observation = NA,
      montant_rachat = NA,
      numero_contrat = NA,
      age_souscription = NA,
      date_naissance = NA,
      date_comptable_entree = NA,
      code_produit = NA,
      pm = NA,
      age = NA
    ) 
  
  proba_rachat <- predict(final_multinom, mp, type = "prob") |> 
    select(tx_rachat_tot = .pred_rachat_total, tx_rachat_part = .pred_rachat_partiel)
  
  prediction_beta <- predict(final_beta, newdata = mp |> mutate(mode_versement = fct_collapse(
    mode_versement, autres = c("PU ", "PP ")
  )), type = "response")
  
  mp <- mp |> 
    bind_cols(proba_rachat) |> 
    mutate(fraction_rachat = prediction_beta)
  
  # choc rachat hausse
  if (choc_rachat_hausse){
    mp <- mp |> 
      mutate(across(c(tx_rachat_part, tx_rachat_tot), \(x) pmin(x * (1 + choc$rachat_hausse), 1)))
  }
  
  # choc rachat baisse
  if (choc_rachat_baisse){
    mp <- mp |> 
      mutate(across(c(tx_rachat_part, tx_rachat_tot), \(x) pmax(x * (1 + choc$rachat_baisse), x - plancher_baisse_rachat)))
  }
  
  return(mp)
}





