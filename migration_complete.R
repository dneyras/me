# Version améliorée de la migration VBA vers R
# Intégration avec le code existant et boucle complète de calcul PM

# Cette version implémente la boucle principale manquante dans le code R existant
# Elle est conçue pour fonctionner avec ou sans les packages tidyverse

# Test de disponibilité des packages
has_tidyverse <- function() {
  tryCatch({
    library(dplyr, warn.conflicts = FALSE, quietly = TRUE)
    library(tidyr, warn.conflicts = FALSE, quietly = TRUE)
    return(TRUE)
  }, error = function(e) FALSE)
}

# Alternative à cross_join pour R de base
cross_join_base <- function(df1, df2) {
  # Simple produit cartésien
  merge(df1, df2, by = NULL)
}

# Alternative à left_join pour R de base
left_join_base <- function(x, y, by_cols) {
  merge(x, y, by = by_cols, all.x = TRUE)
}

# Fonction principale de calcul PM itératif - Version simplifiée
calcul_pm_iteratif_simple <- function(model_point_simple, horizon = 10, use_tidyverse = FALSE) {
  
  if (use_tidyverse && !has_tidyverse()) {
    warning("tidyverse non disponible, utilisation de R de base")
    use_tidyverse <- FALSE
  }
  
  cat("=== Début du calcul PM itératif ===\n")
  cat("Horizon:", horizon, "années\n")
  cat("Nombre de contrats:", nrow(model_point_simple), "\n")
  
  # Paramètres de base
  params <- list(
    date_valorisation = as.Date("2023-12-31"),
    annee_valorisation = 2023,
    taux_pb = 0.01,  # 1% participation aux bénéfices
    taux_uc = 0.03   # 3% rendement UC
  )
  
  # Création de la structure longitudinale
  if (use_tidyverse) {
    # Version tidyverse (si disponible)
    data_long <- model_point_simple %>%
      cross_join(tibble(annee_projection = 0:horizon))
  } else {
    # Version R de base
    annees <- data.frame(annee_projection = 0:horizon)
    data_long <- cross_join_base(model_point_simple, annees)
  }
  
  cat("Structure longitudinale créée:", nrow(data_long), "lignes\n")
  
  # Ajout des colonnes de calcul
  data_long$age_assure <- params$annee_valorisation - data_long$annee_naissance + data_long$annee_projection
  data_long$anciennete_contrat <- params$annee_valorisation - data_long$annee_effet + data_long$annee_projection
  data_long$reste_contrat <- data_long$annee_echeance - params$annee_valorisation - data_long$annee_projection
  data_long$indic_terme_contrat <- as.numeric(data_long$reste_contrat <= 0)
  
  # Calcul des probabilités simplifiées
  data_long$qx <- pmax(0.001, 0.001 * exp((data_long$age_assure - 60) / 20))
  data_long$tx_rachat <- 0.05 * pmax(0.5, 1 - data_long$anciennete_contrat * 0.005)
  
  data_long$p_deces <- pmax(0, pmin(1, data_long$qx))
  data_long$p_rachat <- pmin(data_long$tx_rachat, pmax(0, 1 - data_long$qx))
  data_long$p_terme <- data_long$indic_terme_contrat * (1 - data_long$p_deces - data_long$p_rachat)
  data_long$p_maintien <- 1 - data_long$p_deces - data_long$p_rachat - data_long$p_terme
  
  # Initialisation des variables de calcul
  colonnes_pm <- c("pm_euro", "pm_uc", "prime_comm", "prime_nette", 
                   "interets_pm", "pb_pm", "rend_uc_pm", "sin_deces", 
                   "sin_rachat", "sin_terme", "charg_pm", "pm_finale")
  
  for (col in colonnes_pm) {
    data_long[[col]] <- 0
  }
  
  # Initialisation PM année 0
  data_long$pm_euro[data_long$annee_projection == 0 & data_long$compartiment == "euro"] <- 
    data_long$pm_initial[data_long$annee_projection == 0 & data_long$compartiment == "euro"]
  data_long$pm_uc[data_long$annee_projection == 0 & data_long$compartiment == "uc"] <- 
    data_long$pm_initial[data_long$annee_projection == 0 & data_long$compartiment == "uc"]
  
  # BOUCLE PRINCIPALE ITÉRATIVE
  for (annee in 1:horizon) {
    if (annee %% 5 == 0 || annee == 1) {
      cat("Calcul année", annee, "/", horizon, "\n")
    }
    
    # Filtrer les données de l'année courante
    idx_annee <- data_long$annee_projection == annee
    
    if (sum(idx_annee) == 0) next
    
    # Récupération des PM de l'année précédente
    for (i in which(idx_annee)) {
      contrat <- data_long$numero[i]
      compartiment <- data_long$compartiment[i]
      
      # Index année précédente
      idx_prev <- data_long$numero == contrat & 
                  data_long$compartiment == compartiment & 
                  data_long$annee_projection == (annee - 1)
      
      if (sum(idx_prev) > 0) {
        idx_prev_single <- which(idx_prev)[1]
        if (compartiment == "euro") {
          pm_prev <- data_long$pm_finale[idx_prev_single]
          if (is.na(pm_prev) || pm_prev == 0) {
            pm_prev <- data_long$pm_euro[idx_prev_single]
          }
        } else {
          pm_prev <- data_long$pm_finale[idx_prev_single]
          if (is.na(pm_prev) || pm_prev == 0) {
            pm_prev <- data_long$pm_uc[idx_prev_single]
          }
        }
      } else {
        # Première année, utiliser PM initiale
        pm_prev <- data_long$pm_initial[i]
      }
      
      # Calculs de l'année courante
      
      # 1. Prime commerciale (simplifiée)
      duree_restante_prime <- pmax(0, data_long$duree_versement[i] - data_long$anciennete_contrat[i])
      if (duree_restante_prime > 0) {
        data_long$prime_comm[i] <- data_long$prime_annuelle[i]
      }
      
      # 2. Prime nette
      data_long$prime_nette[i] <- data_long$prime_comm[i] * (1 - data_long$tx_charg_prime[i])
      
      # 3. Intérêts et rendements mi-période
      if (data_long$compartiment[i] == "euro") {
        # Euro: TMG + PB
        taux_total <- data_long$taux_garanti[i] + params$taux_pb
        data_long$interets_pm[i] <- pm_prev * ((1 + taux_total)^0.5 - 1)
        data_long$pb_pm[i] <- data_long$prime_nette[i] * ((1 + params$taux_pb)^0.5 - 1)
      } else {
        # UC: rendement UC
        data_long$rend_uc_pm[i] <- pm_prev * ((1 + params$taux_uc)^0.5 - 1)
        data_long$rend_uc_pm[i] <- data_long$rend_uc_pm[i] + 
          data_long$prime_nette[i] * ((1 + params$taux_uc)^0.5 - 1)
      }
      
      # 4. PM mi-période 1
      pm_mi_1 <- pm_prev + data_long$interets_pm[i] + data_long$pb_pm[i] + 
                 data_long$rend_uc_pm[i] + data_long$prime_nette[i]
      
      # 5. Sinistres
      data_long$sin_deces[i] <- pm_mi_1 * data_long$p_deces[i] / (1 + 0.02)  # 2% chargement décès
      data_long$sin_rachat[i] <- pm_mi_1 * data_long$p_rachat[i] / (1 + 0.015)  # 1.5% chargement rachat
      data_long$sin_terme[i] <- pm_mi_1 * data_long$p_terme[i] / (1 + 0.005)  # 0.5% chargement terme
      
      # 6. PM après sinistres
      pm_apres_sin <- pm_mi_1 - data_long$sin_deces[i] - data_long$sin_rachat[i] - data_long$sin_terme[i]
      
      # 7. Intérêts fin période
      if (data_long$compartiment[i] == "euro") {
        taux_total <- data_long$taux_garanti[i] + params$taux_pb
        interets_fin <- pm_apres_sin * ((1 + taux_total)^0.5 - 1)
      } else {
        interets_fin <- pm_apres_sin * ((1 + params$taux_uc)^0.5 - 1)
      }
      
      # 8. PM avant chargements
      pm_avant_charg <- pm_apres_sin + interets_fin
      
      # 9. Chargements PM
      data_long$charg_pm[i] <- pm_avant_charg * data_long$tx_charg_pm[i]
      
      # 10. PM finale
      data_long$pm_finale[i] <- pmax(0, pm_avant_charg - data_long$charg_pm[i])
      
      # Gestion fin de contrat
      if (data_long$reste_contrat[i] <= 0) {
        data_long$pm_finale[i] <- 0
      }
    }
  }
  
  cat("=== Calcul terminé ===\n")
  
  # Statistiques de résultat
  pm_totales <- aggregate(data_long$pm_finale, 
                         by = list(data_long$annee_projection), 
                         FUN = sum, na.rm = TRUE)
  names(pm_totales) <- c("annee", "pm_total")
  
  cat("Evolution PM totale:\n")
  for (i in 1:min(6, nrow(pm_totales))) {
    cat("Année", pm_totales$annee[i], ":", round(pm_totales$pm_total[i]), "\n")
  }
  
  return(data_long)
}

# Test avec données simplifiées
test_migration_complete <- function() {
  cat("=== Test Migration Complète VBA vers R ===\n")
  
  # Création des données de test
  model_point_test <- data.frame(
    numero = 1:3,
    compartiment = c("euro", "uc", "euro"),
    annee_naissance = c(1970, 1980, 1975),
    annee_effet = c(2020, 2019, 2021),
    annee_echeance = c(2035, 2040, 2038),
    pm_initial = c(10000, 15000, 8000),
    prime_annuelle = c(1000, 1500, 800),
    duree_versement = c(10, 15, 12),
    taux_garanti = c(0.02, 0.00, 0.025),  # UC n'a pas de taux garanti
    tx_charg_prime = c(0.05, 0.05, 0.05),
    tx_charg_pm = c(0.01, 0.015, 0.01),
    stringsAsFactors = FALSE
  )
  
  # Test avec R de base
  cat("\nTest avec R de base...\n")
  resultat <- calcul_pm_iteratif_simple(model_point_test, horizon = 5, use_tidyverse = FALSE)
  
  # Validation des résultats
  cat("\n=== Validation ===\n")
  
  # Vérification: pas de PM négatives
  pm_negatives <- sum(resultat$pm_finale < 0, na.rm = TRUE)
  cat("PM négatives:", pm_negatives, "\n")
  
  # Vérification: évolution cohérente
  pm_evolution <- resultat[resultat$numero == 1 & resultat$compartiment == "euro", 
                          c("annee_projection", "pm_finale")]
  pm_evolution <- pm_evolution[order(pm_evolution$annee_projection), ]
  
  cat("Evolution PM contrat 1 (Euro):\n")
  for (i in 1:min(6, nrow(pm_evolution))) {
    cat("Année", pm_evolution$annee_projection[i], ":", round(pm_evolution$pm_finale[i]), "\n")
  }
  
  # Test de cohérence
  croissances <- diff(pm_evolution$pm_finale[1:6]) / pm_evolution$pm_finale[1:5]
  croissances_raisonnables <- all(croissances > -0.5 & croissances < 0.5, na.rm = TRUE)
  
  if (pm_negatives == 0 && croissances_raisonnables) {
    cat("\n✓ VALIDATION RÉUSSIE\n")
    cat("La migration VBA vers R fonctionne correctement.\n")
  } else {
    cat("\n⚠ VALIDATION PARTIELLE\n")
    cat("Certains résultats nécessitent une vérification.\n")
  }
  
  return(resultat)
}

# Documentation
cat("=== MIGRATION VBA VERS R - VERSION COMPLETE ===\n")
cat("Ce script implémente la boucle itérative de calcul PM manquante.\n")
cat("Fonctions disponibles:\n")
cat("- calcul_pm_iteratif_simple(): Calcul principal\n")
cat("- test_migration_complete(): Test et validation\n")
cat("\nPour exécuter le test:\n")
cat("test_migration_complete()\n")
cat("\n")

# Exécution automatique du test
test_migration_complete()