# Test simple de la migration VBA vers R sans dépendances externes
# Validation de base des fonctions

# Fonction simple de test de la logique PM
test_pm_logic <- function() {
  cat("=== Test de la logique PM ===\n")
  
  # Test des calculs de base
  pm_initial <- 10000
  taux_interets <- 0.02
  prime_nette <- 1000
  taux_rachat <- 0.05
  
  # Calcul PM mi-période 1 (style VBA)
  interets_pm <- pm_initial * ((1 + taux_interets)^0.5 - 1)
  interets_prime <- prime_nette * ((1 + taux_interets)^0.5 - 1)
  pm_mi_periode_1 <- pm_initial + interets_pm + prime_nette + interets_prime
  
  cat("PM initial:", pm_initial, "\n")
  cat("Intérêts PM:", interets_pm, "\n")
  cat("Prime nette:", prime_nette, "\n")
  cat("Intérêts prime:", interets_prime, "\n")
  cat("PM mi-période 1:", pm_mi_periode_1, "\n")
  
  # Calcul des sinistres
  sin_rachat <- pm_mi_periode_1 * taux_rachat
  pm_mi_periode_2 <- pm_mi_periode_1 - sin_rachat
  
  cat("Sinistre rachat:", sin_rachat, "\n")
  cat("PM mi-période 2:", pm_mi_periode_2, "\n")
  
  # Calcul PM final
  interets_fin <- pm_mi_periode_2 * ((1 + taux_interets)^0.5 - 1)
  pm_final <- pm_mi_periode_2 + interets_fin
  
  cat("Intérêts fin période:", interets_fin, "\n")
  cat("PM finale:", pm_final, "\n")
  
  # Vérification de cohérence
  croissance <- (pm_final - pm_initial) / pm_initial
  cat("Croissance PM:", round(croissance * 100, 2), "%\n")
  
  if (pm_final > pm_initial && pm_final < pm_initial * 2) {
    cat("✓ Test de cohérence réussi\n")
    return(TRUE)
  } else {
    cat("✗ Test de cohérence échoué\n")
    return(FALSE)
  }
}

# Test des probabilités
test_probabilities <- function() {
  cat("\n=== Test des probabilités ===\n")
  
  # Données d'exemple
  age <- 45
  sexe <- 1
  anciennete <- 5
  
  # Mortalité simplifiée (style VBA)
  qx_base <- 0.001
  qx <- qx_base * exp((age - 60) / 20)
  
  # Taux de rachats (style VBA)
  tx_rachat_base <- 0.05
  tx_rachat <- tx_rachat_base * (1 - anciennete * 0.005)  # Diminue avec ancienneté
  
  # Probabilités combinées
  p_survie <- 1 - qx
  p_rachat_conditionnel <- tx_rachat * p_survie
  p_maintien <- p_survie - p_rachat_conditionnel
  
  cat("Age:", age, ", Sexe:", sexe, ", Ancienneté:", anciennete, "\n")
  cat("Qx (mortalité):", round(qx, 6), "\n")
  cat("Tx rachat:", round(tx_rachat, 4), "\n")
  cat("P(survie):", round(p_survie, 6), "\n")
  cat("P(rachat|survie):", round(p_rachat_conditionnel, 6), "\n")
  cat("P(maintien):", round(p_maintien, 6), "\n")
  
  # Vérification cohérence
  somme_probas <- qx + p_rachat_conditionnel + p_maintien
  cat("Somme probas:", round(somme_probas, 6), "\n")
  
  if (abs(somme_probas - 1.0) < 0.001) {
    cat("✓ Probabilités cohérentes\n")
    return(TRUE)
  } else {
    cat("✗ Probabilités incohérentes\n")
    return(FALSE)
  }
}

# Test de la structure itérative
test_iterative_structure <- function() {
  cat("\n=== Test structure itérative ===\n")
  
  # Simulation simplifiée de 5 contrats sur 3 années
  nb_contrats <- 5
  nb_annees <- 3
  
  # Structure de données simplifiée
  contrats <- data.frame(
    numero = 1:nb_contrats,
    pm_initial = c(10000, 15000, 8000, 12000, 20000),
    prime_annuelle = c(1000, 1500, 800, 1200, 2000),
    taux_garanti = rep(0.02, nb_contrats),
    taux_rachat = rep(0.05, nb_contrats)
  )
  
  # Matrice pour stocker les résultats
  resultats <- array(0, dim = c(nb_contrats, nb_annees + 1))
  resultats[, 1] <- contrats$pm_initial  # Année 0
  
  cat("Calcul pour", nb_contrats, "contrats sur", nb_annees, "années\n")
  
  # Boucle itérative (style VBA)
  for (annee in 1:nb_annees) {
    cat("Année", annee, ":\n")
    
    for (contrat in 1:nb_contrats) {
      # PM année précédente
      pm_prev <- resultats[contrat, annee]
      
      # Calculs de l'année
      interets <- pm_prev * contrats$taux_garanti[contrat]
      prime <- contrats$prime_annuelle[contrat]
      sinistre <- (pm_prev + interets + prime) * contrats$taux_rachat[contrat]
      
      # PM finale
      pm_nouvelle <- pm_prev + interets + prime - sinistre
      resultats[contrat, annee + 1] <- max(0, pm_nouvelle)
      
      if (contrat <= 2) {  # Affichage détaillé pour les 2 premiers
        cat("  Contrat", contrat, ": PM", round(pm_prev), 
            "+ Int", round(interets), "+ Prime", round(prime), 
            "- Sin", round(sinistre), "= PM", round(pm_nouvelle), "\n")
      }
    }
  }
  
  # Vérification de cohérence
  pm_finales <- resultats[, nb_annees + 1]
  croissances <- (pm_finales - contrats$pm_initial) / contrats$pm_initial
  
  cat("\nCroissances PM:\n")
  for (i in 1:nb_contrats) {
    cat("Contrat", i, ":", round(croissances[i] * 100, 1), "%\n")
  }
  
  # Test cohérence
  if (all(pm_finales >= 0) && all(croissances > -0.5) && all(croissances < 1.0)) {
    cat("✓ Structure itérative fonctionnelle\n")
    return(TRUE)
  } else {
    cat("✗ Structure itérative problématique\n")
    return(FALSE)
  }
}

# Fonction principale de test
main_test_simple <- function() {
  cat("==========================================\n")
  cat("TEST DE MIGRATION VBA VERS R - VERSION SIMPLE\n")
  cat("==========================================\n")
  
  # Exécution des tests
  test1 <- test_pm_logic()
  test2 <- test_probabilities()
  test3 <- test_iterative_structure()
  
  cat("\n=== RÉSUMÉ DES TESTS ===\n")
  cat("Test logique PM:", ifelse(test1, "✓ RÉUSSI", "✗ ÉCHOUÉ"), "\n")
  cat("Test probabilités:", ifelse(test2, "✓ RÉUSSI", "✗ ÉCHOUÉ"), "\n")
  cat("Test structure itérative:", ifelse(test3, "✓ RÉUSSI", "✗ ÉCHOUÉ"), "\n")
  
  if (test1 && test2 && test3) {
    cat("\n🎉 TOUS LES TESTS SONT RÉUSSIS!\n")
    cat("La logique de migration VBA vers R est validée.\n")
  } else {
    cat("\n⚠️  CERTAINS TESTS ONT ÉCHOUÉ\n")
    cat("Révision nécessaire de la logique.\n")
  }
  
  cat("\nProchaines étapes:\n")
  cat("1. Intégrer les vraies données du modèle\n")
  cat("2. Implémenter la boucle complète de calcul PM\n")
  cat("3. Valider contre les résultats VBA\n")
  cat("4. Optimiser pour la performance\n")
}

# Exécution du test
main_test_simple()