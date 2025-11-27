# ==============================================================================
# 1. FONCTIONS DE CALCUL (PARTIE EURO)
# ==============================================================================

calculer_flux_euro <- function(curr, prev) {
  # Regroupement 1 : De l'initialisation jusqu'à la PM Mi-Période 2
  curr[, `:=`(
    interets_pm_mi_periode_euro = prev$pm_euro * ((1 + taux_annuel_garanti) ^ 0.5 - 1),
    pb_mi_periode_euro          = prev$pm_euro * ((1 + pb) ^ 0.5 - 1)
  )]
  
  # Calcul PM1 + Sinistres + Chargements Sinistres en un bloc pour utiliser les vars temporaires
  curr[, `:=`(
    pm_mi_periode_1_euro = prime_nette_euro + interets_prime_euro + pb_prime_euro + prev$pm_euro + 
      interets_pm_mi_periode_euro + pb_mi_periode_euro
  )]
  
  curr[, `:=`(
    sin_deces_euro       = pm_mi_periode_1_euro * (1 - 0.5 * (tx_tirage + tx_rachat_tot + tx_rachat_part)) * (qx_approx / (1 + ch_deces)),
    sin_tirage_euro      = pm_mi_periode_1_euro * pmin(tx_tirage, pmax(0, 1 - qx_approx)) / (1 + ch_tirage),
    sin_rachat_tot_euro  = pm_mi_periode_1_euro * pmin(tx_rachat_tot, pmax(0, 1 - qx_approx - tx_tirage)) / (1 + ch_rachat_tot),
    sin_rachat_part_euro = pm_mi_periode_1_euro * pmin(tx_rachat_part, pmax(0, 1 - qx_approx - tx_tirage - tx_rachat_tot)) / (1 + ch_rachat_part)
  )]
  
  # Calcul des chargements sinistres et PM2
  curr[, `:=`(
    charg_deces_euro       = sin_deces_euro * ch_deces,
    charg_tirage_euro      = sin_tirage_euro * ch_tirage,
    charg_rachat_tot_euro  = sin_rachat_tot_euro * ch_rachat_tot,
    charg_rachat_part_euro = sin_rachat_part_euro * ch_rachat_part
  )]
  
  curr[, pm_mi_periode_2_euro := pm_mi_periode_1_euro - sin_deces_euro - sin_tirage_euro - 
         sin_rachat_tot_euro - sin_rachat_part_euro - charg_deces_euro - 
         charg_tirage_euro - charg_rachat_tot_euro - charg_rachat_part_euro]
  
  # Regroupement 2 : Terme, Transferts et PM4
  # Astuce : On calcule sin_terme avec une multiplication (indic * ...) au lieu de ifelse
  curr[, `:=`(
    sin_terme_euro = indic_terme_contrat * (
      pm_mi_periode_2_euro / (1 + ch_prest_terme) - (prev$pm_euro + 
                                                       interets_pm_mi_periode_euro + interets_prime_euro + (prime_nette_euro - sin_deces_euro - sin_tirage_euro - sin_rachat_tot_euro -
                                                                                                              sin_rachat_part_euro - pm_mi_periode_2_euro) / 2) * tx_charg_pm_euro
    )
  )]
  
  curr[, `:=`(
    charg_terme_euro = sin_terme_euro * ch_prest_terme,
    pm_mi_periode_3_euro = pm_mi_periode_2_euro - sin_terme_euro - (sin_terme_euro * ch_prest_terme)
  )]
  
  # Gestion Transferts (Si colonne existe pas, init à 0 via data.table in-place)
  if (!"pm_mi_periode_3_uc" %in% names(curr)) curr[, pm_mi_periode_3_uc := 0]
  
  curr[, `:=`(
    cap_euro_uc = pm_mi_periode_3_euro * tx_pass_euro_uc / (1 + ch_arb_euro_uc),
    cap_uc_euro = pm_mi_periode_3_uc * tx_pass_uc_euro / (1 + ch_arb_uc_euro)
  )]
  
  curr[, `:=`(
    charg_transf_euro_uc = cap_euro_uc * ch_arb_euro_uc,
    pm_mi_periode_4_euro = pm_mi_periode_3_euro + cap_uc_euro - cap_euro_uc - (cap_euro_uc * ch_arb_euro_uc)
  )]
}

calculer_cloture_euro <- function(curr, prev) {
  
  curr[, `:=`(
    interets_fin_periode_euro = pm_mi_periode_4_euro * ((1 + taux_annuel_garanti) ^ 0.5 - 1),
    pb_fin_periode_euro       = pm_mi_periode_4_euro * ((1 + pb) ^ 0.5 - 1)
  )]
  
  # Calculs finaux avec logique conditionnelle mathématique (plus rapide que ifelse sur gros vecteurs)
  # Pour pm_cloture : (duree != 1) est converti en 0 ou 1
  curr[, pm_cloture_euro := (duree_restante != 1) * (pm_mi_periode_4_euro + interets_fin_periode_euro + pb_fin_periode_euro)]
  
  # Bloc optimisé : Calcul de la base chargement, du chargement et de la commission
  curr[, `:=`(
    base_charg_pm_euro = prev$pm_euro + interets_prime_euro + interets_fin_periode_euro + interets_pm_mi_periode_euro + 
      pb_prime_euro + pb_fin_periode_euro + pb_mi_periode_euro + 
      0.5 * (prime_nette_euro + cap_uc_euro - cap_euro_uc - charg_transf_euro_uc - sin_deces_euro - sin_tirage_euro - 
               sin_rachat_tot_euro - sin_rachat_part_euro - sin_terme_euro)
  )]
  
  curr[, charg_pm_euro := {
    # Variable temporaire locale pour le cas "Terme"
    base_terme_tmp <- (prev$pm_euro + interets_pm_mi_periode_euro + pb_mi_periode_euro + interets_prime_euro + pb_prime_euro +
                         (prime_nette_euro - sin_deces_euro - sin_tirage_euro - sin_rachat_tot_euro - sin_rachat_part_euro - pm_mi_periode_2_euro) / 2)
    
    # Logique vectorielle : indic * CasTerme + (1-indic) * CasVie
    indic_terme_contrat * (base_terme_tmp * tx_charg_pm_euro) + 
      (1 - indic_terme_contrat) * pmin(tx_charg_pm_euro * base_charg_pm_euro, pm_cloture_euro)
  }]
  
  curr[, commissions_pm_euro := tx_com_sur_encours_euro * base_charg_pm_euro]
  curr[, pm_euro := pmax(0, pm_cloture_euro - charg_pm_euro)]
  
  # Marges & Total Chargements
  curr[, `:=`(
    marg_euro = prime_comm_euro + prev$pm_euro + interets_pm_mi_periode_euro + interets_prime_euro - sin_deces_euro - 
      sin_tirage_euro - sin_rachat_tot_euro - sin_rachat_part_euro - sin_terme_euro + interets_fin_periode_euro - 
      pm_euro - cap_euro_uc + cap_uc_euro - commissions_pm_euro - commissions_primes_euro + 
      pb_prime_euro + pb_fin_periode_euro + pb_mi_periode_euro,
    
    charg_euro = chargement_prime_euro + charg_deces_euro + charg_tirage_euro + charg_rachat_tot_euro + charg_rachat_part_euro + 
      charg_terme_euro + charg_pm_euro + charg_transf_euro_uc - commissions_pm_euro - commissions_primes_euro
  )]
}

# ==============================================================================
# 2. FONCTIONS DE CALCUL (PARTIE UC)
# ==============================================================================

calculer_flux_uc <- function(curr, prev) {
  
  # 1. Intérêts et PM1
  curr[, `:=`(
    interets_prime_uc = 0, # Explicite
    interets_pm_mi_periode_uc = 0,
    rend_uc_prime = prime_nette_uc*((1 + rendement_uc)^0.5 - 1),
    rend_uc_mi_periode_uc = prev$pm_uc * ((1 + rendement_uc) ^ 0.5 - 1)
  )]
  
  curr[, pm_mi_periode_1_uc := prime_nette_uc + prev$pm_uc + rend_uc_prime + rend_uc_mi_periode_uc] # + 0 + 0 implicites
  
  # 2. Sinistres (Calcul groupé)
  curr[, `:=`(
    sin_deces_uc = pm_mi_periode_1_uc * (1 - 0.5 * (tx_tirage + tx_rachat_tot + tx_rachat_part)) * (qx_approx / (1 + ch_deces)),
    sin_tirage_uc = pm_mi_periode_1_uc * pmin(tx_tirage, pmax(0, 1 - qx_approx)) / (1 + ch_tirage),
    sin_rachat_tot_uc = pm_mi_periode_1_uc * pmin(tx_rachat_tot, pmax(0, 1 - qx_approx - tx_tirage)) / (1 + ch_rachat_tot),
    sin_rachat_part_uc = pm_mi_periode_1_uc * pmin(tx_rachat_part, pmax(0, 1 - qx_approx - tx_tirage - tx_rachat_tot)) / (1 + ch_rachat_part)
  )]
  
  # 3. Chargements Sinistres & PM2
  curr[, `:=`(
    charg_deces_uc = sin_deces_uc * ch_deces,
    charg_tirage_uc = sin_tirage_uc * ch_tirage,
    charg_rachat_tot_uc = sin_rachat_tot_uc * ch_rachat_tot,
    charg_rachat_part_uc = sin_rachat_part_uc * ch_rachat_part
  )]
  
  curr[, pm_mi_periode_2_uc := pm_mi_periode_1_uc - sin_deces_uc - sin_tirage_uc - sin_rachat_tot_uc - sin_rachat_part_uc - 
         charg_deces_uc - charg_tirage_uc - charg_rachat_tot_uc - charg_rachat_part_uc]
  
  # 4. Terme & Transferts UC
  # Utilisation d'un bloc {} pour base_sin_terme_uc afin de ne pas créer de colonne persistante
  curr[, sin_terme_uc := {
    base_tmp <- (prev$pm_uc + interets_pm_mi_periode_uc + rend_uc_mi_periode_uc + interets_prime_uc + rend_uc_prime + 
                   (prime_nette_uc - sin_deces_uc - sin_tirage_uc - sin_rachat_tot_uc - sin_rachat_part_uc - pm_mi_periode_2_uc) / 2) *
      (taux_retros_uc_global + tx_charg_pm_uc)
    
    indic_terme_contrat * (pm_mi_periode_2_uc / (1 + ch_prest_terme) - base_tmp)
  }]
  
  curr[, `:=`(
    charg_terme_uc = sin_terme_uc * ch_prest_terme,
    pm_mi_periode_3_uc = pm_mi_periode_2_uc - sin_terme_uc - (sin_terme_uc * ch_prest_terme)
  )]
  
  # Transferts (chargements)
  curr[, `:=`(
    charg_transf_uc_euro = cap_uc_euro * ch_arb_uc_euro,
    charg_transf_uc_uc = pm_mi_periode_3_uc * tx_pass_uc_uc * ch_arb_uc_uc
  )]
  
  curr[, pm_mi_periode_4_uc := pm_mi_periode_3_uc + cap_euro_uc - cap_uc_euro - charg_transf_uc_euro - charg_transf_uc_uc]
}

calculer_cloture_uc <- function(curr, prev) {
  # 1. Clôture brute
  curr[, `:=`(
    interets_fin_periode_uc = 0,
    rend_fin_periode_uc = pm_mi_periode_4_uc * ((1 + rendement_uc) ^ 0.5 - 1)
  )]
  curr[, pm_cloture_uc := pm_mi_periode_4_uc + rend_fin_periode_uc] # +0
  
  # 2. Rétrocessions et Chargements PM (Grosse optimisation RAM ici)
  # On calcule tout dans un grand bloc pour garder les bases "temp"
  curr[, `:=`(
    retro_global_pm_uc = 0, # Init pour typage correct si besoin, ou on laisse le bloc créer
    retro_aepm_uc = 0,
    charg_pm_uc = 0,
    commissions_pm_uc = 0
  )]
  
  # Correction ici : Utilisation de c("col1", "col2"...) := { ... list() }
  curr[, c("retro_global_pm_uc", "retro_aepm_uc", "charg_pm_uc", "commissions_pm_uc") := {
    
    # --- Variables temporaires (vecteurs en mémoire volatile) ---
    base_terme_tmp <- (prev$pm_uc + interets_pm_mi_periode_uc + rend_uc_mi_periode_uc + 
                         interets_prime_uc + rend_uc_prime +
                         (prime_nette_uc - sin_deces_uc - sin_tirage_uc - sin_rachat_tot_uc - sin_rachat_part_uc - pm_mi_periode_2_uc) / 2)
    
    base_non_terme_tmp <- (prev$pm_uc + interets_prime_uc + rend_uc_prime + 
                             interets_fin_periode_uc + rend_fin_periode_uc + 
                             interets_pm_mi_periode_uc + rend_uc_mi_periode_uc + 
                             0.5 * (prime_nette_uc + cap_euro_uc - cap_uc_euro - 
                                      charg_transf_uc_euro - charg_transf_uc_uc - 
                                      sin_deces_uc - sin_tirage_uc - sin_rachat_tot_uc - 
                                      sin_rachat_part_uc - sin_terme_uc))
    
    # --- Calcul des rétrocessions ---
    # Logique : indic * Terme + (1-indic) * NonTerme
    r_global <- indic_terme_contrat * (base_terme_tmp * taux_retros_uc_global) + 
      (1 - indic_terme_contrat) * pmin(base_non_terme_tmp * taux_retros_uc_global, pm_cloture_uc)
    
    r_aepm   <- indic_terme_contrat * (base_terme_tmp * taux_retro_uc_part_afi_esca) + 
      (1 - indic_terme_contrat) * pmin(base_non_terme_tmp * taux_retro_uc_part_afi_esca, pm_cloture_uc)
    
    # --- Calcul des chargements (dépend de r_global) ---
    c_pm     <- indic_terme_contrat * (base_terme_tmp * tx_charg_pm_uc) + 
      (1 - indic_terme_contrat) * pmin((base_non_terme_tmp - 0.5 * r_global) * tx_charg_pm_uc, pm_cloture_uc - r_global)
    
    # --- Commissions ---
    comm     <- tx_com_sur_encours_uc * base_non_terme_tmp
    
    # On retourne une liste pour l'assignation multiple dans data.table
    list(r_global, r_aepm, c_pm, comm)
  }]
  
  # 3. PM Finale et Marges
  curr[, pm_uc := pmax(0, pm_cloture_uc - charg_pm_uc - retro_global_pm_uc)]
  
  curr[, `:=`(
    marg_uc = prime_comm_uc + prev$pm_uc + interets_pm_mi_periode_uc + interets_prime_uc - sin_deces_uc - sin_tirage_uc - 
      sin_rachat_tot_uc - sin_rachat_part_uc - sin_terme_uc + interets_fin_periode_uc - pm_uc - cap_uc_euro + cap_euro_uc - 
      commissions_pm_uc - commissions_primes_uc + rend_uc_prime + rend_uc_mi_periode_uc + rend_fin_periode_uc,
    
    charg_uc = chargement_prime_uc + charg_deces_uc + charg_tirage_uc + charg_rachat_tot_uc + charg_rachat_part_uc + 
      charg_terme_uc + charg_pm_uc + retro_global_pm_uc + charg_transf_uc_euro + charg_transf_uc_uc - commissions_pm_uc - commissions_primes_uc
  )]
}