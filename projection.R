library(tidyverse)
library(data.table)

source("scripts/fonctions.R")
source("scripts/calculs.R")

load("data/Rdata/hypotheses_2024.RData")

mp_2024 <- readRDS("data/Rdata/mp_2024.RDS")

proj_flex <- function(mp, 
                      date_valorisation, 
                      choix_choc = NULL, # Default NULL plus propre
                      hypotheses_morta, 
                      hypotheses_produits, 
                      hypotheses_rachat, 
                      hypotheses_rendement,
                      horizon_projection = 50
                      ) {
  
  # --- 1. Initialisation ---
  date_valorisation <- ymd(date_valorisation)
  annee_valorisation <- lubridate::year(date_valorisation)
  
  # Définition des chocs (Structure plus propre)
  choc_config <- list(
    mortalite = FALSE, longevite = FALSE, catastrophe = FALSE,
    rachat_hausse = FALSE, rachat_baisse = FALSE, rachat_massif = FALSE
  )
  
  vals_choc <- list("mortalite" = 0.15, "longevite" = -0.2, "catastrophe" = 0.0015, 
                    "rachat_hausse" = 0.5, "rachat_baisse" = -0.5, "rachat_massif" = 0.4)
  
  # Activation du choc sélectionné sans utiliser 'assign' (plus sûr)
  if (!is.null(choix_choc) && choix_choc %in% names(choc_config)) {
    choc_config[[choix_choc]] <- TRUE
  }
  
  # --- 2.  (Vectorisé) ---
  setDT(mp)
 
  # Copie de sécurité pour ne pas modifier l'objet global de l'utilisateur
  mp <- copy(mp) 
  
  # préparation des données
  mp <- process_mp(mp, annee_valorisation) 
  
  
  # Projection structurelle (création des lignes années futures)
  mp <- proj_init_dt_optim(mp, date_valorisation, annee_valorisation)
  
  # Clé primaire pour accélérer les jointures futures ou les tris
  setkeyv(mp, c("provenance", "numero", "annee_projection"))
  
  # --- 3. Jointures des hypothèses ---
  
  # Jointure rachat
  mp <- join_assumptions_rachat(mp, hypotheses_rachat, 
                                choc_config$rachat_hausse, 
                                choc_config$rachat_baisse, 
                                choc_config$rachat_massif, 
                                vals_choc)
  
  # Jointure autres hypothèses
  mp <- join_assumptions_dt_fast(mp, hypotheses_morta, hypotheses_produits, hypotheses_rendement, 
                                 choc_config$mortalite, 
                                 choc_config$longevite, 
                                 choc_config$catastrophe, 
                                 vals_choc)
  
  # Calculs pré-boucle (Probas, Primes)
  mp <- calcul_prob_dt(mp)
  mp <- calcul_prime_dt(mp)
  
  # --- 4. Gestion des NA (Optimisation Majeure) ---
  
  # Identification des colonnes PM
  cols_pm <- names(mp)[grepl("pm", names(mp), fixed = TRUE)]
  
  setnafill(mp, type = "const", fill = 0, cols = cols_pm)
  
  # --- 5. Boucle de Projection (Cœur du calcul) ---
  
  # Split par année : Création de la liste pour l'itération
  # keep.by = TRUE est important pour garder la colonne année
  dat <- split(mp, by = "annee_projection", keep.by = TRUE)
  
  for (compteur_annee in 1:horizon_projection) {
    
    # Pointers vers les data.tables (Pas de copie mémoire ici, c'est des pointeurs)
    bdd_prev <- dat[[compteur_annee]]
    bdd_curr <- dat[[compteur_annee + 1]]
    
    # Appel des fonctions optimisées (Modifient bdd_curr par référence)
    
    # --- FLUX EURO ---
    calculer_flux_euro(bdd_curr, bdd_prev)
    
    # --- FLUX UC ---
    calculer_flux_uc(bdd_curr, bdd_prev)
    
    # --- CLÔTURES ---
    calculer_cloture_euro(bdd_curr, bdd_prev)
    calculer_cloture_uc(bdd_curr, bdd_prev)
  }
  
  # Fusion rapide
  resultats_projection <- rbindlist(dat, fill = TRUE)
  
  # Génération des KPIs finaux
  flexing <- generer_flexings(resultats_projection, date_valorisation)
  
  return(list(flexing = flexing, projection = resultats_projection))
}

bench::system_time(
  proj_flex_2024_init <- proj_flex(mp_2024,
                                   date_valorisation = "2024-12-31", 
                                   choix_choc = NULL, 
                                   hypotheses_morta = hypotheses_morta, 
                                   hypotheses_produits = hypotheses_produits_2024, 
                                   hypotheses_rachat = hypotheses_rachat_2024, 
                                   hypotheses_rendement = hypotheses_rendement))


flexing_central_2024_init_apres_modif <- proj_flex_2024_init[[1]]

identical(flexing_central_2024_init_apres_modif, flexing_2024_init)


proj_central_2024_init <- proj_flex_2024_init[[2]]

mp <- mp_2024
date_valorisation <- "2024-12-31" 
choix_choc <- NULL
hypotheses_produits <- hypotheses_produits_2024
hypotheses_rachat <- hypotheses_rachat_2024 
horizon_projection <- 50


bench::system_time(proj_flex_2024 <- map(
  c(
    "mortalite",
    "longevite",
    "catastrophe",
    "rachat_hausse",
    "rachat_baisse",
    "rachat_massif"
  ),
  \(x) proj_flex(
    mp_2024,
    date_valorisation = "2024-12-31",
    choix_choc = x,
    hypotheses_morta = hypotheses_morta,
    hypotheses_produits = hypotheses_produits_2024,
    hypotheses_rachat = hypotheses_rachat_2024,
    hypotheses_rendement = hypotheses_rendement
  )
))


flexings <- map(proj_flex_2024, \(x) pluck(x, 1))
