# Migration VBA vers R - Calcul itératif des PM
# Ce script implémente la boucle principale de calcul des provisions mathématiques
# qui manquait dans l'implémentation R existante

library(tidyverse)
library(lubridate)

# Source des fonctions existantes
source("base_model_point.R")
source("pre_process.R")

#' Fonction principale de calcul itératif des PM
#' Équivalent de la boucle CompteurAnnee du VBA
calcul_pm_iteratif <- function(model_point_data, hypotheses_data, hypotheses2_data, 
                               params, pb_uc_data, thtf_data, horizon = 50) {
  
  # Initialisation des données
  message("Initialisation des calculs...")
  
  # Créer la structure longitudinale avec toutes les années de projection
  data_long <- model_point_data %>%
    cross_join(tibble(annee_projection = 0:horizon)) %>%
    # Ajouter les calculs préliminaires déjà existants
    mutate(
      date_projection = params$date_valorisation + dyears(annee_projection),
      age_assure = params$annee_valorisation - annee_naissance + annee_projection,
      anciennete_contrat = params$annee_valorisation - annee_effet + annee_projection,
      reste_contrat = annee_echeance - params$annee_valorisation - annee_projection,
      indic_terme_contrat = date_projection >= date_echeance & date_projection < date_echeance + dyears(1),
      duree_restante = pmax(year(date_echeance) - year(date_projection) + 1, 0), 
      duree_restante_prime = pmax(duree_restante_prime - annee_projection + 1, 0),
      nb_cloture = nb_tetes_bis
    ) %>%
    # Jointure avec les tables de mortalité
    left_join(thtf_data %>% select(sexe, age, qx_approx), 
              by = join_by(age_assure == age, sexe_client == sexe)) %>%
    # Jointure avec les hypothèses
    left_join(hypotheses_data, 
              by = join_by(nom_produit == produit, anciennete_contrat == anciennete)) %>%
    left_join(hypotheses2_data, 
              by = join_by(anciennete_contrat == anciennete, cat_rachat_tot == categorie_rachat)) %>%
    # Calculs des probabilités
    mutate(
      p_deces = pmax(0, pmin(1, qx_approx)),
      p_tirage = pmin(tx_tirage, pmax(0, 1 - qx_approx)),
      p_rachat_tot = pmin(tx_rachat_tot, pmax(0, 1 - qx_approx - tx_tirage)),
      p_rachat_part = pmin(tx_rachat_part, pmax(0, 1 - qx_approx - tx_tirage - tx_rachat_tot)),
      p_terme = as.numeric(indic_terme_contrat) * (1 - p_deces - p_tirage - p_rachat_tot),
      p_cloture = 1 - p_deces - p_tirage - p_rachat_tot - p_terme
    ) %>%
    # Calculs des coefficients de prime
    mutate(
      coeff_prime_annee_echeance = if_else(duree_restante_prime == 1, 
                                           coeff_prime_bonus(mois_effet, periodicite), 0),
      coeff_prime_annee_bonus = if_else(delai_bonus == 0, 
                                        coeff_prime_bonus(mois_effet, periodicite), 0)
    ) %>%
    # Trier par contrat et année de projection
    arrange(numero, compartiment, type_bonus, annee_projection)
  
  message("Structure longitudinale créée.")
  
  # Initialisation des colonnes pour le calcul itératif
  data_long <- data_long %>%
    mutate(
      # PM initiales (année 0)
      pm_euro = if_else(annee_projection == 0 & compartiment == "euro", pm, 0),
      pm_uc = if_else(annee_projection == 0 & compartiment == "uc", pm, 0),
      
      # Variables de calcul à initialiser
      prime_comm = 0,
      chargement_prime = 0,
      commissions_primes = 0,
      prime_nette = 0,
      interets_prime = 0,
      pb_prime = 0,
      rendement_uc_prime = 0,
      
      # Variables PM mi-période
      pm_mi_periode_1 = 0,
      pm_mi_periode_2 = 0,
      pm_mi_periode_3 = 0,
      pm_mi_periode_4 = 0,
      
      # Sinistres
      sin_deces = 0,
      sin_tirage = 0,
      sin_rachat_tot = 0,
      sin_rachat_part = 0,
      sin_terme = 0,
      
      # Chargements sinistres
      charg_deces = 0,
      charg_tirage = 0,
      charg_rachat_tot = 0,
      charg_rachat_part = 0,
      charg_terme = 0,
      
      # Autres variables
      interets_pm_mi_periode = 0,
      pb_mi_periode = 0,
      rendement_mi_periode = 0,
      charg_pm = 0,
      commissions_pm = 0,
      pm_cloture = 0
    )
  
  message("Début de la boucle itérative...")
  
  # BOUCLE PRINCIPALE ITÉRATIVE (équivalent CompteurAnnee VBA)
  for (annee in 1:horizon) {
    message(paste("Calcul année", annee, "/", horizon))
    
    # Filtrer les données pour l'année courante
    data_annee_courante <- data_long %>% 
      filter(annee_projection == annee)
    
    # Jointure avec les données PB/UC pour l'année
    pb_uc_annee <- pb_uc_data %>% 
      filter(annee_projection == annee) %>%
      select(pb, rendement_uc)
    
    if (nrow(pb_uc_annee) == 0) {
      pb_annee <- 0
      rendement_uc_annee <- 0
    } else {
      pb_annee <- pb_uc_annee$pb[1]
      rendement_uc_annee <- pb_uc_annee$rendement_uc[1]
    }
    
    # 1. CALCUL DES PRIMES
    data_annee_courante <- data_annee_courante %>%
      mutate(
        # Récupération de la probabilité de clôture cumulée
        p_clot_cum = get_p_cloture_cumulee(numero, compartiment, type_bonus, annee, data_long),
        
        # Calcul de la prime commerciale
        prime_comm = calcul_prime_avec(duree_restante_prime, coeff_prime_annee_echeance, 
                                       prime_commerciale, p_clot_cum, nb_tetes),
        
        # Chargements et commissions sur primes
        chargement_prime = prime_comm * tx_charg_prime,
        commissions_primes = prime_comm * taux_com_prime,
        
        # Prime nette
        prime_nette = prime_comm - chargement_prime,
        
        # Intérêts et PB sur primes (mi-période)
        interets_prime = if_else(compartiment == "euro", 
                                prime_nette * ((1 + taux_annuel_garanti)^0.5 - 1), 0),
        pb_prime = if_else(compartiment == "euro", 
                          prime_nette * ((1 + pb_annee)^0.5 - 1), 0),
        rendement_uc_prime = if_else(compartiment == "uc", 
                                    prime_nette * ((1 + rendement_uc_annee)^0.5 - 1), 0)
      )
    
    # 2. CALCUL PM MI-PÉRIODE 1
    data_annee_courante <- data_annee_courante %>%
      mutate(
        # Récupération PM année précédente
        pm_precedente = get_pm_precedente(numero, compartiment, type_bonus, annee, data_long),
        
        # Intérêts et PB sur PM mi-période
        interets_pm_mi_periode = if_else(compartiment == "euro",
                                        pm_precedente * ((1 + taux_annuel_garanti)^0.5 - 1), 0),
        pb_mi_periode = if_else(compartiment == "euro",
                               pm_precedente * ((1 + pb_annee)^0.5 - 1), 0),
        rendement_mi_periode = if_else(compartiment == "uc",
                                      pm_precedente * ((1 + rendement_uc_annee)^0.5 - 1), 0),
        
        # PM mi-période 1
        pm_mi_periode_1 = prime_nette + interets_prime + pb_prime + rendement_uc_prime +
                          pm_precedente + interets_pm_mi_periode + pb_mi_periode + rendement_mi_periode
      )
    
    # 3. CALCUL DES SINISTRES ET CHARGEMENTS
    data_annee_courante <- data_annee_courante %>%
      mutate(
        # Sinistres
        sin_deces = pm_mi_periode_1 * p_deces / (1 + ch_deces),
        sin_tirage = pm_mi_periode_1 * p_tirage / (1 + ch_tirage),
        sin_rachat_tot = pm_mi_periode_1 * p_rachat_tot / (1 + ch_rachat_tot),
        sin_rachat_part = pm_mi_periode_1 * p_rachat_part / (1 + ch_rachat_part),
        
        # Chargements sur sinistres
        charg_deces = sin_deces * ch_deces,
        charg_tirage = sin_tirage * ch_tirage,
        charg_rachat_tot = sin_rachat_tot * ch_rachat_tot,
        charg_rachat_part = sin_rachat_part * ch_rachat_part,
        
        # PM mi-période 2
        pm_mi_periode_2 = pm_mi_periode_1 - sin_deces - sin_tirage - sin_rachat_tot - 
                          sin_rachat_part - charg_deces - charg_tirage - charg_rachat_tot - charg_rachat_part
      )
    
    # 4. CALCUL SINISTRES TERME ET PM MI-PÉRIODE 3
    data_annee_courante <- data_annee_courante %>%
      mutate(
        sin_terme = if_else(compartiment == "euro",
          # Version Euro
          as.numeric(indic_terme_contrat) * (pm_mi_periode_2 / (1 + ch_prest_terme) - 
            (pm_precedente + interets_pm_mi_periode + interets_prime + 
             (prime_nette - sin_deces - sin_tirage - sin_rachat_tot - sin_rachat_part - pm_mi_periode_2) / 2) * 
            tx_charg_pm),
          # Version UC
          as.numeric(indic_terme_contrat) * (pm_mi_periode_2 / (1 + ch_prest_terme) - 
            (pm_precedente + interets_pm_mi_periode + rendement_mi_periode + interets_prime + rendement_uc_prime + 
             (prime_nette - sin_deces - sin_tirage - sin_rachat_tot - sin_rachat_part - pm_mi_periode_2) / 2) * 
            (tx_retro_global + tx_charg_pm))
        ),
        
        charg_terme = sin_terme * ch_prest_terme,
        
        pm_mi_periode_3 = pm_mi_periode_2 - sin_terme - charg_terme
      )
    
    # 5. CALCUL TRANSFERTS (à implémenter selon besoin)
    # Pour l'instant, on suppose pas de transferts
    data_annee_courante <- data_annee_courante %>%
      mutate(pm_mi_periode_4 = pm_mi_periode_3)
    
    # 6. CALCUL PM CLÔTURE ET FINAL
    data_annee_courante <- data_annee_courante %>%
      mutate(
        # Intérêts et PB fin période
        interets_fin_periode = if_else(compartiment == "euro",
                                      pm_mi_periode_4 * ((1 + taux_annuel_garanti)^0.5 - 1), 0),
        pb_fin_periode = if_else(compartiment == "euro",
                                pm_mi_periode_4 * ((1 + pb_annee)^0.5 - 1), 0),
        rendement_fin_periode = if_else(compartiment == "uc",
                                       pm_mi_periode_4 * ((1 + rendement_uc_annee)^0.5 - 1), 0),
        
        # PM clôture
        pm_cloture = if_else(duree_restante == 1, 0, 
                            pm_mi_periode_4 + interets_fin_periode + pb_fin_periode + rendement_fin_periode),
        
        # Chargements PM
        charg_pm = if_else(as.numeric(indic_terme_contrat) == 1,
          # Cas contrat à terme
          (pm_precedente + interets_pm_mi_periode + pb_mi_periode + interets_prime + pb_prime +
           (prime_nette - sin_deces - sin_tirage - sin_rachat_tot - sin_rachat_part - pm_mi_periode_2) / 2) * 
          tx_charg_pm,
          # Cas normal
          pmin(tx_charg_pm * (pm_precedente + interets_prime + interets_fin_periode + interets_pm_mi_periode +
                             pb_prime + pb_fin_periode + pb_mi_periode + rendement_uc_prime + rendement_fin_periode + 
                             rendement_mi_periode + 0.5 * (prime_nette - sin_deces - sin_tirage - 
                             sin_rachat_tot - sin_rachat_part - sin_terme)), pm_cloture)
        ),
        
        # Commissions PM
        commissions_pm = tx_com_sur_encours * (pm_precedente + interets_prime + interets_fin_periode + 
                                              interets_pm_mi_periode + pb_prime + pb_fin_periode + pb_mi_periode +
                                              rendement_uc_prime + rendement_fin_periode + rendement_mi_periode +
                                              0.5 * (prime_nette - sin_deces - sin_tirage - sin_rachat_tot - 
                                                    sin_rachat_part - sin_terme)),
        
        # PM finale
        pm_finale = pmax(0, pm_cloture - charg_pm)
      )
    
    # Mise à jour des données dans data_long
    data_long <- data_long %>%
      rows_update(
        data_annee_courante %>% 
          select(numero, compartiment, type_bonus, annee_projection, 
                 prime_comm, chargement_prime, commissions_primes, prime_nette,
                 interets_prime, pb_prime, rendement_uc_prime,
                 pm_mi_periode_1, pm_mi_periode_2, pm_mi_periode_3, pm_mi_periode_4,
                 sin_deces, sin_tirage, sin_rachat_tot, sin_rachat_part, sin_terme,
                 charg_deces, charg_tirage, charg_rachat_tot, charg_rachat_part, charg_terme,
                 interets_pm_mi_periode, pb_mi_periode, rendement_mi_periode,
                 charg_pm, commissions_pm, pm_cloture, pm_finale),
        by = c("numero", "compartiment", "type_bonus", "annee_projection")
      )
    
    # Mise à jour des PM pour l'année suivante
    if (annee < horizon) {
      data_long <- data_long %>%
        mutate(
          pm_euro = if_else(annee_projection == annee + 1 & compartiment == "euro",
                           get_pm_finale_precedente(numero, "euro", type_bonus, annee + 1, data_long), pm_euro),
          pm_uc = if_else(annee_projection == annee + 1 & compartiment == "uc",
                         get_pm_finale_precedente(numero, "uc", type_bonus, annee + 1, data_long), pm_uc)
        )
    }
  }
  
  message("Calculs terminés.")
  return(data_long)
}

# Fonctions auxiliaires pour récupérer les valeurs des années précédentes
get_p_cloture_cumulee <- function(numero, compartiment, type_bonus, annee, data_long) {
  if (annee == 1) return(1)
  
  p_clot_values <- data_long %>%
    filter(numero == !!numero, compartiment == !!compartiment, 
           type_bonus == !!type_bonus, annee_projection > 0, annee_projection < !!annee) %>%
    pull(p_cloture)
  
  if (length(p_clot_values) == 0) return(1)
  return(cumprod(p_clot_values)[length(p_clot_values)])
}

get_pm_precedente <- function(numero, compartiment, type_bonus, annee, data_long) {
  if (annee == 1) {
    pm_init <- data_long %>%
      filter(numero == !!numero, compartiment == !!compartiment, 
             type_bonus == !!type_bonus, annee_projection == 0) %>%
      pull(pm)
    return(ifelse(length(pm_init) > 0, pm_init[1], 0))
  }
  
  pm_prev <- data_long %>%
    filter(numero == !!numero, compartiment == !!compartiment, 
           type_bonus == !!type_bonus, annee_projection == !!annee - 1) %>%
    pull(pm_finale)
  
  return(ifelse(length(pm_prev) > 0, pm_prev[1], 0))
}

get_pm_finale_precedente <- function(numero, compartiment, type_bonus, annee, data_long) {
  pm_prev <- data_long %>%
    filter(numero == !!numero, compartiment == !!compartiment, 
           type_bonus == !!type_bonus, annee_projection == !!annee - 1) %>%
    pull(pm_finale)
  
  return(ifelse(length(pm_prev) > 0, pm_prev[1], 0))
}

# Fonction wrapper pour faciliter l'utilisation
run_calcul_pm_migration <- function() {
  # Cette fonction sera appelée une fois que les données seront disponibles
  message("Fonction de migration VBA vers R créée.")
  message("Pour utiliser: calcul_pm_iteratif(model_point_data, hypotheses_data, hypotheses2_data, params, pb_uc_data, thtf_data)")
}