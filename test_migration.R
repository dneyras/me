# Test et validation de la migration VBA vers R
# Ce script teste l'implémentation de la boucle itérative

library(tidyverse)
library(lubridate)

# Source des scripts
source("calcul_pm_iteratif.R")

# Création de données de test simples pour validation
create_test_data <- function() {
  
  # Paramètres de base
  params <- list(
    date_valorisation = ymd("2023-12-31"),
    annee_valorisation = 2023,
    prime_ou_sans = "Sans Primes"
  )
  
  # Données de model point simplifiées pour test
  model_point_test <- tibble(
    numero = 1:5,
    num_adh = c(101, 102, 103, 104, 105),
    nom_produit = c(1, 2, 1, 3, 2),
    sexe_client = c(1, 2, 1, 2, 1),
    annee_naissance = c(1970, 1980, 1975, 1985, 1965),
    date_de_naissance = ymd(c("1970-06-15", "1980-03-20", "1975-09-10", "1985-12-05", "1965-01-30")),
    annee_effet = c(2020, 2019, 2021, 2022, 2018),
    date_effet = ymd(c("2020-01-01", "2019-06-01", "2021-03-01", "2022-09-01", "2018-12-01")),
    mois_effet = c(1, 6, 3, 9, 12),
    annee_echeance = c(2035, 2040, 2038, 2045, 2033),
    date_echeance = ymd(c("2035-01-01", "2040-06-01", "2038-03-01", "2045-09-01", "2033-12-01")),
    nb_tetes = 1,
    compartiment = rep(c("euro", "uc"), length.out = 5),
    type_bonus = rep(c("1", "2"), length.out = 5),
    pm = c(10000, 15000, 8000, 12000, 20000),
    prime_commerciale = c(1000, 1500, 800, 1200, 2000),
    tx_charg_prime = 0.05,
    taux_com_prime = 0.02,
    tx_charg_pm = 0.01,
    tx_com_sur_encours = 0.005,
    taux_annuel_garanti = 0.02,
    tx_retro_global = 0.001,
    periodicite = "Annuel",
    duree_versement = 10,
    annuites_avant_bonus = 5,
    delai_bonus = 2,
    nb_tetes_bis = 1,
    cat_rachat_tot = 2,
    position_contrat = "En cours",
    tx_charg_suspens_pp = 0,
    contrat_prorogeable = FALSE,
    indic_obseque = FALSE,
    form_p = 1
  ) %>%
  mutate(
    duree_restante_prime = duree_versement - (params$annee_valorisation - annee_effet) - 1
  )
  
  # Données d'hypothèses simplifiées
  hypotheses_test <- expand_grid(
    produit = 1:3,
    anciennete = 0:20
  ) %>%
  mutate(
    ch_deces = 0.02,
    ch_tirage = 0.01,
    ch_rachat_tot = 0.015,
    ch_rachat_part = 0.01,
    ch_prest_terme = 0.005,
    ch_pm_prev = 0.001,
    ch_arb_euro_uc = 0.001,
    ch_arb_uc_euro = 0.001,
    ch_arb_uc_uc = 0.001,
    tx_tirage = 0.05,
    tx_pass_euro_uc = 0.02,
    tx_pass_uc_euro = 0.02,
    tx_pass_uc_uc = 0.02,
    tx_tech_pre2011 = 0.025,
    tx_tech_post2011 = 0.02
  )
  
  # Données d'hypothèses 2 (rachats par catégorie)
  hypotheses2_test <- expand_grid(
    categorie_rachat = 1:5,
    anciennete = 0:20
  ) %>%
  mutate(
    tx_rachat_tot = case_when(
      categorie_rachat == 1 ~ 0.03,  # Obsèques
      categorie_rachat == 2 ~ 0.05,  # Libre/Unique Form P 1
      categorie_rachat == 3 ~ 0.04,  # Libre/Unique Form P 2
      categorie_rachat == 4 ~ 0.035, # Périodique Form P 1
      categorie_rachat == 5 ~ 0.025, # Périodique Form P 2
      TRUE ~ 0.04
    ),
    tx_rachat_part = tx_rachat_tot * 0.3,
    tx_rachat_tot_prev = tx_rachat_tot * 0.8
  )
  
  # Données PB et rendement UC
  pb_uc_test <- tibble(
    annee_projection = 0:50,
    pb = 0.01,  # 1% de PB
    rendement_uc = 0.03  # 3% de rendement UC
  )
  
  # Table de mortalité simplifiée
  thtf_test <- expand_grid(
    sexe = 1:2,
    age = 0:120
  ) %>%
  mutate(
    qx = pmax(0.001, 0.001 * exp((age - 60) / 20)),  # Mortalité simplifiée
    qx_approx = (qx + lag(qx, default = qx)) / 2
  )
  
  return(list(
    model_point = model_point_test,
    hypotheses = hypotheses_test,
    hypotheses2 = hypotheses2_test,
    params = params,
    pb_uc = pb_uc_test,
    thtf = thtf_test
  ))
}

# Test de l'implémentation
test_calcul_pm <- function() {
  message("=== Test de la migration VBA vers R ===")
  
  # Création des données de test
  test_data <- create_test_data()
  
  message("Données de test créées.")
  message(paste("Nombre de contrats:", nrow(test_data$model_point)))
  
  # Test avec un horizon réduit pour validation
  horizon_test <- 5
  
  tryCatch({
    # Lancement du calcul
    resultat <- calcul_pm_iteratif(
      model_point_data = test_data$model_point,
      hypotheses_data = test_data$hypotheses,
      hypotheses2_data = test_data$hypotheses2,
      params = test_data$params,
      pb_uc_data = test_data$pb_uc,
      thtf_data = test_data$thtf,
      horizon = horizon_test
    )
    
    message("=== Calcul terminé avec succès ===")
    
    # Affichage des résultats de validation
    message("\n=== Résultats de validation ===")
    
    # Évolution des PM pour le premier contrat
    evolution_pm <- resultat %>%
      filter(numero == 1) %>%
      select(numero, compartiment, type_bonus, annee_projection, pm, pm_finale, prime_comm) %>%
      arrange(compartiment, type_bonus, annee_projection)
    
    print("Évolution PM contrat 1:")
    print(evolution_pm)
    
    # Synthèse par année
    synthese_annuelle <- resultat %>%
      group_by(annee_projection) %>%
      summarise(
        nb_contrats = n_distinct(numero),
        total_pm_euro = sum(pm_finale[compartiment == "euro"], na.rm = TRUE),
        total_pm_uc = sum(pm_finale[compartiment == "uc"], na.rm = TRUE),
        total_primes = sum(prime_comm, na.rm = TRUE),
        .groups = 'drop'
      )
    
    print("\nSynthèse annuelle:")
    print(synthese_annuelle)
    
    return(resultat)
    
  }, error = function(e) {
    message(paste("Erreur durant le calcul:", e$message))
    return(NULL)
  })
}

# Fonction de comparaison avec les résultats VBA (quand disponibles)
compare_with_vba <- function(resultat_r, resultat_vba = NULL) {
  if (is.null(resultat_vba)) {
    message("Pas de données VBA disponibles pour comparaison.")
    return(invisible())
  }
  
  # Comparaison des PM finales
  # À implémenter quand les données VBA seront disponibles
  message("Comparaison avec VBA à implémenter...")
}

# Fonction de validation de cohérence
validate_coherence <- function(resultat) {
  message("\n=== Validation de cohérence ===")
  
  # Vérification 1: PM ne deviennent pas négatives
  pm_negatives <- resultat %>%
    filter(pm_finale < 0) %>%
    nrow()
  
  message(paste("Nombre de PM négatives:", pm_negatives))
  
  # Vérification 2: Continuité des calculs
  gaps <- resultat %>%
    group_by(numero, compartiment, type_bonus) %>%
    arrange(annee_projection) %>%
    summarise(
      nb_annees = n(),
      annees_consecutives = all(diff(annee_projection) == 1),
      .groups = 'drop'
    ) %>%
    filter(!annees_consecutives) %>%
    nrow()
  
  message(paste("Nombre de séries avec des gaps:", gaps))
  
  # Vérification 3: Évolution réaliste
  croissance_extreme <- resultat %>%
    group_by(numero, compartiment, type_bonus) %>%
    arrange(annee_projection) %>%
    mutate(croissance = (pm_finale / lag(pm_finale, default = 1)) - 1) %>%
    filter(abs(croissance) > 2, !is.na(croissance)) %>%  # Plus de 200% de croissance
    nrow()
  
  message(paste("Nombre de croissances extrêmes (>200%):", croissance_extreme))
  
  if (pm_negatives == 0 && gaps == 0 && croissance_extreme == 0) {
    message("✓ Tous les tests de cohérence sont passés.")
  } else {
    message("⚠ Certains tests de cohérence ont échoué. Vérification nécessaire.")
  }
}

# Exécution des tests
main_test <- function() {
  resultat <- test_calcul_pm()
  
  if (!is.null(resultat)) {
    validate_coherence(resultat)
    
    # Sauvegarde des résultats pour inspection
    saveRDS(resultat, "test_results_migration_vba_r.rds")
    message("\nRésultats sauvegardés dans 'test_results_migration_vba_r.rds'")
    
    return(resultat)
  }
}

# Pour exécuter le test, décommenter la ligne suivante:
# main_test()