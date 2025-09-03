library(tidyverse)
library(readxl)

transform_sexe <- function(sexe_texte) {
  sexe_texte_upper <- toupper(as.character(sexe_texte))
  dplyr::case_when(
    sexe_texte_upper == "H" ~ 1L,
    sexe_texte_upper == "F" ~ 2L,
    TRUE ~ NA_integer_
  )
}

#' Mapping des noms de produits VBA vers des codes numériques.
mapping_produits_r <- c(
  "AEP" = 1, "AEP2" = 2, "APA" = 3, "APA2" = 4, "APIU" = 5, "ARA" = 6,
  "CEP1" = 7, "CFC" = 8, "CFCP" = 9, "CFD" = 10, "CFDP" = 11, "CMDV" = 12,
  "COFE" = 13, "EPA" = 14, "EPR" = 15, "EPR2" = 16, "ERS" = 17, "FIDP" = 18,
  "FOFE" = 19, "AINV" = 20, "MAD" = 21, "OB10" = 22, "ORP" = 23, "PERL" = 24,
  "SPT" = 25, "SRP" = 26, "TRIA" = 27, "TRIP" = 28, "COE" = 29, "OBEP" = 30,
  "PRF2" = 31, "PRF5" = 32, "AEUC" = 33, "UCAF" = 34, "PRF1" = 35, "PERS" = 36,
  "PSVV" = 37, "PSVC" = 38, "PSVU" = 39, "PSUC" = 40, "OBER" = 41, "AEPA" = 42,
  "ALIB" = 43, "ASVI" = 44, "ASCA" = 45, "ASTR" = 46,
  "0" = 0, "-" = 0
)

#' Transforme un nom de produit en code numérique.
transform_nom_prod <- function(nom_prod_vba) {
  nom_prod_vba_char <- as.character(nom_prod_vba)
  code <- mapping_produits_r[nom_prod_vba_char]
  if(is.na(nom_prod_vba_char) || nom_prod_vba_char == "") {
    return(mapping_produits_r[""]) 
  }
  if (is.na(code) && !nom_prod_vba_char %in% names(mapping_produits_r[mapping_produits_r == 0])) {
    warning(paste("Nom de produit VBA non trouvé dans le mapping:", nom_prod_vba_char), call. = FALSE)
    return(NA_integer_) 
  }
  if (is.na(code) && nom_prod_vba_char %in% names(mapping_produits_r[mapping_produits_r == 0])) {
    return(0L)
  }
  return(as.integer(code))
}


params <- list()
params$date_valorisation <- ymd("2023-12-31")
params$annee_valorisation <- lubridate::year(params$date_valorisation)



model_point <- readxl::read_excel("data/modele_epargne.xlsm", sheet = "MODEL POINT", guess_max = 50000) |> 
  janitor::clean_names() |> 
  rename(tx_charg_pm_euro = tx_charg_sur_pm_euro) |> 
  pivot_longer(cols = ends_with(c("euro", "uc")), names_to = c(".value", "type"), names_pattern = "(.*)_(uc|euro)") |> 
  mutate(
    annee_naissance = year(date_de_naissance), .after = date_de_naissance
  ) |> 
  mutate(annee_echeance = year(date_echeance), .after = date_echeance) |> 
  mutate(nb_tetes = 1)





model_point <- readxl::read_excel("data/modele_epargne.xlsm", sheet = "MODEL POINT", guess_max = 50000) |> 
  janitor::clean_names() |> 
  rename(tx_charg_pm_euro = tx_charg_sur_pm_euro,
         annuites_avant_bonus_1 = annuites_avant_bonus1
  ) |> 
  pivot_longer(cols = ends_with(c("euro", "uc")), names_to = c(".value", "compartiment"), names_pattern = "(.*)_(uc|euro)") |> 
  pivot_longer(cols = ends_with(c("1", "2")), names_to = c(".value", "type_bonus"), names_pattern = "(.*)_(1|2)") |>
  mutate(
    nom_produit = purrr::map_int(nom_produit, transform_nom_prod),
    sexe_client = transform_sexe(sexe_client),
  ) |> 
  mutate(
    annee_naissance = year(date_de_naissance), .after = date_de_naissance
  ) |> 
  mutate(
    annee_effet = year(date_effet),
    mois_effet = month(date_effet),
    .after = date_effet
  ) |> 
  mutate(annee_echeance = year(date_echeance), .after = date_echeance) |> 
  mutate(nb_tetes = 1) |>
  mutate(tx_charg_pm = if_else(position_contrat == "En réduction" & tx_charg_pm > 0, tx_charg_pm + tx_charg_suspens_pp, tx_charg_pm)) |> 
  mutate(delai_bonus = annuites_avant_bonus - (params$annee_valorisation - annee_effet) - 1,
         delai_bonus = case_when(
           periodicite == "Semestriel" & mois_effet > 6 ~ delai_bonus + 1,
           periodicite == "Trimestriel" & mois_effet > 3 ~ delai_bonus + 1,
           delai_bonus
         )) |> 
  mutate(contrat_prorogeable = as.logical(contrat_prorogeable),
         indic_obseque = as.logical(indic_obseque)
         ) |> 
  mutate(duree_restante_prime = duree_versement - (params$annee_valorisation - annee_effet) - 1,
         duree_restante_prime = case_when(
           periodicite == "Semestriel" & mois_effet > 6 ~ duree_restante_prime + 1,
           periodicite == "Trimestriel" & mois_effet > 3 ~ duree_restante_prime + 1,
           .default = duree_restante_prime
         )
         ) |> 
  mutate(
    nb_tetes_bis = case_when(
      compartiment == "euro" & pm > 0 ~ 1,
      compartiment == "uc" & pm > 0 ~ 1,
      .default = 0
    )
  ) |> 
  mutate(
    cat_rachat_tot = case_when(
      periodicite %in% c("Libre", "Unique") & form_p == 1 ~ 2,
      periodicite %in% c("Libre", "Unique") & form_p == 2 ~ 3,
      !(periodicite %in% c("Libre", "Unique")) & form_p == 1 ~ 4,
      !(periodicite %in% c("Libre", "Unique")) & form_p == 2 ~ 5,
      indic_obseque ~ 1,
    )
  )





mp <- readxl::read_excel("data/modele_epargne.xlsm", sheet = "MODEL POINT", guess_max = 50000) |> 
  janitor::clean_names()


mp_sample <- mp |> 
  slice_sample(n = 1000)


openxlsx::write.xlsx(mp_sample, file = "data/mp_sample.xlsx")


mp_sample <- read_excel("data/mp_sample.xlsx")

test <- model_point |> 
  cross_join(tibble(annee_projection = 0:50))


test2 <- test |> 
  mutate(
    date_projection = params$date_valorisation + dyears(annee_projection),
    age_assure = params$annee_valorisation - annee_naissance + annee_projection,
    anciennete_contrat = params$annee_valorisation - annee_effet + annee_projection,
    reste_contrat = annee_echeance - params$annee_valorisation - annee_projection,
    indic_terme_contrat = params$date_valorisation + dyears(annee_projection) >= date_echeance & params$date_valorisation + dyears(annee_projection) < date_echeance + dyears(1),
    date_anniversaire = f_anniversaire(date_effet, params, annee_projection),
    duree_restante = max(year(date_echeance) - year(date_projection), 0), 
    duree_restante_prime = max(duree_restante_prime - annee_projection)
         )


f_anniversaire <- function(date_effet, params, proj){
  
  new_date_proj <- date_effet
  
  year(new_date_proj) <- year(params$date_valorisation + dyears(proj))
  
  if_else(new_date_proj >= params$date_valorisation + dyears(proj), new_date_proj, new_date_proj + dyears(1))
  
}

test3 <- test2 |> 
  filter(indic_terme_contrat)


test4 <- test2 |> 
  filter(numero == 401)





####################################################################
#Sample
####################################################################

f_anniversaire_bis <- function(date_effet, params, proj){
  
  new_date_proj <- date_effet
  
  year(new_date_proj) <- year(params$date_valorisation + dyears(proj))
  
  new_date_proj
  
}

#init
process_model_point_sample <- mp_sample |> 
  rename(tx_charg_pm_euro = tx_charg_sur_pm_euro,
         annuites_avant_bonus_1 = annuites_avant_bonus1
  ) |> 
  pivot_longer(cols = ends_with(c("euro", "uc")), names_to = c(".value", "compartiment"), names_pattern = "(.*)_(uc|euro)") |> 
  pivot_longer(cols = ends_with(c("1", "2")), names_to = c(".value", "type_bonus"), names_pattern = "(.*)_(1|2)") |>
  mutate(
    annee_naissance = year(date_de_naissance), .after = date_de_naissance
  ) |> 
  mutate(
    annee_effet = year(date_effet),
    mois_effet = month(date_effet),
    .after = date_effet
  ) |> 
  mutate(annee_echeance = year(date_echeance), .after = date_echeance) |> 
  mutate(nb_tetes = 1) |>
  mutate(tx_charg_pm = if_else(position_contrat == "En réduction" & tx_charg_pm > 0, tx_charg_pm + tx_charg_suspens_pp, tx_charg_pm)) |> 
  mutate(delai_bonus = annuites_avant_bonus - (params$annee_valorisation - annee_effet) - 1,
         delai_bonus = case_when(
           periodicite == "Semestriel" & mois_effet > 6 ~ delai_bonus + 1,
           periodicite == "Trimestriel" & mois_effet > 3 ~ delai_bonus + 1,
           .default = delai_bonus
         )) |> 
  mutate(contrat_prorogeable = as.logical(contrat_prorogeable),
         indic_obseque = as.logical(indic_obseque)
  ) |> 
  mutate(duree_restante_prime = duree_versement - (params$annee_valorisation - annee_effet) - 1,
         duree_restante_prime = case_when(
           periodicite == "Semestriel" & mois_effet > 6 ~ duree_restante_prime + 1,
           periodicite == "Trimestriel" & mois_effet > 3 ~ duree_restante_prime + 1,
           .default = duree_restante_prime
         )
  ) |> 
  mutate(
    nb_tetes_bis = case_when(
      compartiment == "euro" & pm > 0 ~ 1,
      compartiment == "uc" & pm > 0 ~ 1,
      .default = 0
    )
  ) |> 
  mutate(
    cat_rachat_tot = case_when(
      periodicite %in% c("Libre", "Unique") & form_p == 1 ~ 2,
      periodicite %in% c("Libre", "Unique") & form_p == 2 ~ 3,
      !(periodicite %in% c("Libre", "Unique")) & form_p == 1 ~ 4,
      !(periodicite %in% c("Libre", "Unique")) & form_p == 2 ~ 5,
      indic_obseque ~ 1,
    )
  )


#calculs préliminaires
process_model_point_sample <- process_model_point_sample |> 
  cross_join(tibble(annee_projection = 0:50))


process_model_point_sample <- process_model_point_sample |> 
  mutate(
    date_projection = params$date_valorisation + dyears(annee_projection),
    age_assure = params$annee_valorisation - annee_naissance + annee_projection,
    anciennete_contrat = params$annee_valorisation - annee_effet + annee_projection,
    reste_contrat = annee_echeance - params$annee_valorisation - annee_projection,
    indic_terme_contrat = params$date_valorisation + dyears(annee_projection) >= date_echeance & params$date_valorisation + dyears(annee_projection) < date_echeance + dyears(1),
    date_anniversaire = f_anniversaire_bis(date_effet, params, annee_projection),
    duree_restante = pmax(year(date_echeance) - year(date_projection) + 1, 0), 
    duree_restante_prime = pmax(duree_restante_prime - annee_projection + 1, 0),
    nb_cloture = nb_tetes_bis
    )


# jointure des qx 


process_model_point_sample <- process_model_point_sample |> 
  left_join(thtf |> select(sexe, age, qx_approx), by = join_by(age_assure == age, sexe_client == sexe))
  



# CalculsNbDeces_NbTirages_NbRachatsTot_NbRachatsPart_NbTermes_NbClotures()


a <- process_model_point_sample |> 
  left_join(hypotheses, join_by(nom_produit == produit, anciennete_contrat == anciennete)) |> 
  left_join(hypotheses2, join_by(anciennete_contrat == anciennete, categorie_rachat == categorie_rachat)) |> 
  mutate(
    p_deces = pmax(0, pmin(1, qx_approx)),
    p_tirage = pmin(tx_tirage, pmax(0, 1 - qx)),
    p_rachat_tot = pmin(tx_rachat_tot, pmax(0, 1 - qx_approx - tx_tirage)),
    p_rachat_part = pmin(tx_rachat_part, pmax(0, 1 - qx_approx - tx_tirage - tx_rachat_tot)),
    p_terme = indic_terme_contrat*(1 - p_deces - p_tirage - p_rachat_tot),
    p_clot = 1 - p_deces - p_tirage - p_rachat_tot - p_terme, .by = numero
  ) |> 
  filter(annee_projection > 0) |> 
  mutate(
    p_clot = replace_na(p_clot, 0),
    p_clot = cumprod(p_clot), .by= c(numero, compartiment, type_bonus)
    )




# CalculsCoeffPrimeAnneeEch()


# il faut utiliser cette fonction 


f_coeff <- function(mois, periode){
  
  nb_max <- 12 %/% periode 
  
  res <- ceiling((12 - mois + 1) / periode)
  
  return((nb_max - res)/nb_max)
  
}


map(1:12, \(x) f_coeff(x, 1))

map(1:12, \(x) f_coeff(x, 3))

map(1:12, \(x) f_coeff(x, 6))

map(1:12, \(x) f_coeff(x, 12))



coeff_prime_bonus <- function(mois, periode){
  
  case_match(periode,
        
             "Unique" ~ 0,
             "Annuel" ~ f_coeff(mois, 12),
             "Semestriel" ~ f_coeff(mois, 6),
             "Trimestriel" ~ f_coeff(mois, 3),
             "Mensuel" ~ f_coeff(mois, 1)
             )
  
}

# bizarre d'enlever 1 jour à la date d'anniversaire 

b <- a |>
  mutate(
    coeff_prime_annee_echeance = if_else(duree_restante_prime == 1,
      case_match(periodicite,
      "Unique"~ 0,
      "Annuel" ~ 1, 
      "Semestriel" ~ if_else(mois_effet <= 6, 1, 1/2),
      "Trimestriel" ~ case_when(mois_effet <= 3 ~ 1, mois_effet <= 6 ~ 1/4, mois_effet <= 9 ~ 1/2, .default = 3/4),
      "Mensuel" ~ (12 - (month(date_projection) - month(date_anniversaire - ddays(1))))/12
    ), 0)
  )


b_bis <- a |> 
  mutate(
    coeff_prime_annee_echeance = if_else(duree_restante_prime == 1, coeff_prime_bonus(mois_effet, periodicite), 0),
    coeff_prime_annee_bonus = if_else(delai_bonus == 0, coeff_prime_bonus(mois_effet, periodicite), 0)
  )


view(b |> filter(num_adh == 57031614, duree_restante_prime == 1) |> 
       select(numero, num_adh, periodicite, date_projection, date_anniversaire, duree_restante_prime, coeff_prime_annee_echeance))


view(test |> filter(num_adh_original == 57031614, duree_restante_prime_annee == 1))


test <- read_excel("data/modele_epargne_sample.xlsm", sheet = "BDD_Export_VBA")

r <- test |> 
  distinct(num_adh = num_adh_original, coeff_prime_annee_echeance)

v <- b |> 
  distinct(num_adh, coeff_prime_annee_echeance)

view(a |> select(numero, num_adh, annee_projection, delai_bonus))

sample_vba <- read_excel("data/modele_epargne_sample.xlsm", sheet = "BDD_Export_VBA") |> 
  janitor::clean_names() |> 
  filter(!(num_adh_original %in% c(3020007, 3020017, 3020024, 3020026, 3020028, 3020030, 3020036, 3020007)))


thtf <- thf_002 |> 
  mutate(qx_approx = (qx + lag(qx, default = 0))/2)


## CalculsPrime


params$prime_ou_sans <- "Sans Primes"


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


calcul_prime <- function(prime_ou_sans, ...){
  
  case_match(prime_ou_sans,
             
             "Sans Primes" ~ calcul_prime_sans(...),
             "Primes" ~ calcul_prime_avec(...)
             
             )

}




c <- b_bis |> 
  mutate(
    prime_comm = calcul_prime_sans(duree_restante_prime, delai_bonus, coeff_prime_annee_bonus, coeff_prime_annee_echeance, prime_commerciale, nb_cloture, nb_tetes)
  )


c <- b_bis |> 
  mutate(
    prime_comm = calcul_prime_avec(duree_restante_prime, coeff_prime_annee_echeance, prime_commerciale, p_clot, nb_tetes)
  )



ftest <- function(data_vba, data_r, variable_vba, variable_r, val){
  
  vba <- data_vba |> 
    filter(prime_commerciale_euro > 0) |> 
    distinct(num_adh_original, {{variable_vba}})
  
  r <- data_r |> 
    filter(prime_comm > 0, compartiment == "euro") |> 
    distinct(num_adh, {{variable_r}})
  

  
  if (val) return(list(vba, r)) else return(all(near(vba[[2]], r[[2]])))
  
  
}

ftest(sample_vba, c, prime_commerciale_euro, prime_comm)


## calculchargprime 

d <- c |> 
  mutate(
    chargement_prime = prime_comm*tx_charg_prime
  )



x <- ftest(sample_vba, d, chargements_prime_euro, chargement_prime, F)


## calculcomissionprime 

d <- d |> 
  mutate(
    commissions_primes = prime_comm*taux_com_prime
  )

ftest(sample_vba, d, commissions_prime_euro, commissions_primes, F)

## calculprimenette


d <- d |> 
  left_join(pb_uc, by = "annee_projection")


d <- d |> 
  mutate(
    prime_nette = prime_comm - chargement_prime,
    interets_prime = if_else(compartiment == "euro", prime_nette*((1 + taux_annuel_garanti)^0.5 - 1), 0), 
    pb_prime = if_else(compartiment == "euro", prime_nette*((1 + pb)^0.5 - 1), 0),
    rendement_uc_prime = if_else(compartiment == "uc", prime_nette*((1 + rendement_uc)^0.5 - 1), 0)
  )


ftest(sample_vba, d, prime_nette_euro, prime_nette, F)



## calcul pm mi periode 1 

e <- d |> 
  mutate(
    taux_interets_pm_mi_periode = (1 + taux_annuel_garanti)^0.5 - 1,
    taux_pb_mi_periode = (1 + pb)^0.5 - 1,
  )




f_pm_mi_periode_1_uc <- function(pm, rendement_uc, prime_nette, interets_prime, rendement_prime){
  
  interets_pm_mi_periode = 0
  
  rendement_mi_periode = pm*((1 + rendement_uc)^0.5 - 1)
  
  pm_mi_periode = prime_nette + interets_prime + pm + interets_pm_mi_periode + rendement_prime + rendement_mi_periode
  
  return(pm_mi_periode)
}


f_pm_mi_periode_1_euro <- function(pm, tmg, pb, prime_nette, interets_prime, pb_prime){
  
  interets_pm_mi_periode = pm*((1 + tmg)^0.5 - 1)
  
  pb_mi_periode = pm*((1 + pb)^0.5 - 1)
  
  pm_mi_periode = prime_nette + interets_prime + pb_prime + pm + interets_pm_mi_periode + pb_mi_periode
  
  return(pm_mi_periode)
}


f_pm_mi_periode_1 <- function(compartiment, ...){
  
  switch (compartiment,
          
    euro = f_pm_mi_periode_1_euro(...),
    
    uc = f_pm_mi_periode_1_uc(...)
  )
  
}



f_pm_mi_periode_2 <- function(pm, taux_deces_pm, taux_tirage_pm, taux_rachat_tot_pm, taux_rachat_part_pm, ch_deces, ch_tirage, ch_rachat_tot, ch_rachat_part){
  
  
  sin_deces = pm*taux_deces_pm
  
  sin_tirage = pm*taux_tirage_pm
  
  sin_rachat_tot = pm*taux_rachat_tot_pm
  
  sin_rachat_part = pm*taux_rachat_part_pm
  
  
  chargement_deces = sin_deces*ch_deces
  
  chargement_tirage = sin_tirage*ch_tirage
  
  chargement_rachat_tot = sin_rachat_tot*ch_rachat_tot
  
  chargement_rachat_part = sin_rachat_part*ch_rachat_part
  
  pm_mi_periode = pm - sin_deces - sin_tirage - sin_rachat_tot - sin_rachat_part - chargement_deces - chargement_tirage - chargement_rachat_tot - chargement_rachat_part
  
  
  return(pm_mi_periode)
  
}

f_pm_mi_periode_3_euro <- function(pm_mi_periode_2, pm_prev, ind_terme_contrat, ch_prest_terme, interets_pm_mi_periode, interets_prime, prime_nette, sin_deces, sin_tirage, sin_rachat_tot, sin_rachat_part, taux_chargement_pm){
  
  sin_terme = ind_terme_contrat*(pm_mi_periode_2/(1 + ch_prest_terme) - (pm_prev + interets_pm_mi_periode + interets_prime + (prime_nette - sin_deces - sin_tirage - sin_rachat_tot - sin_rachat_part - pm_mi_periode_2)/2)*taux_chargement_pm)
  
  chargement_terme = sin_terme*ch_prest_terme 
  
  pm_mi_periode_3 = pm_mi_periode_2 - sin_terme - chargement_terme
  
  return(pm_mi_periode_3)
                                 
                                 
}


f_pm_mi_periode_3_uc <- function(pm_mi_periode_2, pm_prev, ind_terme_contrat, ch_prest_terme, interets_pm_mi_periode, rendement_mi_periode, interets_prime, rendement_prime, prime_nette, sin_deces, sin_tirage, sin_rachat_tot, sin_rachat_part, taux_retro, taux_chargement_pm){
  
  sin_terme = ind_terme_contrat*(pm_mi_periode_2/(1 + ch_prest_terme) - (pm_prev + interets_pm_mi_periode + rendement_mi_periode + interets_prime + rendement_prime + (prime_nette - sin_deces - sin_tirage - sin_rachat_tot - sin_rachat_part - pm_mi_periode_2)/2)*(taux_retro + taux_chargement_pm))
  
  chargement_terme = sin_terme*ch_prest_terme 
  
  pm_mi_periode_3 = pm_mi_periode_2 - sin_terme - chargement_terme
  
  return(pm_mi_periode_3)
  
  
}



f_pm_mi_periode_3 <- function(compartiment, ...){
  
  switch (compartiment,
          
          euro = f_pm_mi_periode_3_euro(...),
          
          uc = f_pm_mi_periode_3_uc(...)
  )
  
}


f_pm_mi_periode_4_euro <- function(pm_mi_periode_3_euro, pm_mi_periode_3_uc, tx_cap_euro_uc, tx_cap_uc_euro, charb_euro_uc){
  
  cap_euro_uc = pm_mi_periode_3_euro*tx_cap_euro_uc
  
  charg_transf_euro_uc = cap_euro_uc*charb_euro_uc 
  
  cap_uc_euro = pm_mi_periode_3_uc*tx_cap_uc_euro
  
  pm_mi_periode_4_euro = pm_mi_periode_3_euro + cap_uc_euro - cap_euro_uc - charg_transf_euro_uc
  
  return(pm_mi_periode_4_euro)
  
}


f_pm_cloture_euro <- function(pm, tmg, pb, prime_nette, interets_prime, pb_prime){
  
  interets_pm_fin_periode = pm*((1 + tmg)^0.5 - 1)
  
  pb_fin_periode = pm*((1 + pb)^0.5 - 1)
  
  pm_cloture = ifelse(duree_restante == 1, 0, pm + interets_pm_fin_periode + pb_fin_periode)
  
  return(pm_mi_periode)
}


calcul_charg_pm_euro <- function(ind_terme_contrat, pm_euro_prev, interets_pm_mi_periode_euro, 
                                 pb_mi_periode_euro, interets_prime_euro, pb_prime_euro, 
                                 prime_nette_euro, sin_deces_euro, sin_tirage_euro, 
                                 sin_rachat_tot_euro, sin_rachat_part_euro, pm_mi_periode2_euro, 
                                 taux_chargement_pm_euro, interets_fin_periode_euro, 
                                 pb_fin_periode_euro, cap_uc_euro, cap_euro_uc, 
                                 charg_transf_euro_uc, sin_terme_euro, pm_cloture_euro) {
  
  if (ind_terme_contrat == 1) {
    charg_pm_euro <- (pm_euro_prev + interets_pm_mi_periode_euro + pb_mi_periode_euro + 
                        interets_prime_euro + pb_prime_euro + 
                        (prime_nette_euro - sin_deces_euro - sin_tirage_euro - 
                           sin_rachat_tot_euro - sin_rachat_part_euro - pm_mi_periode2_euro) / 2) * 
      taux_chargement_pm_euro
  } else {
    charg_pm_euro <- min(taux_chargement_pm_euro * 
                           (pm_euro_prev + interets_prime_euro + interets_fin_periode_euro + 
                              interets_pm_mi_periode_euro + pb_prime_euro + pb_fin_periode_euro + 
                              pb_mi_periode_euro + 0.5 * (prime_nette_euro + cap_uc_euro - 
                                                            cap_euro_uc - charg_transf_euro_uc - sin_deces_euro - 
                                                            sin_tirage_euro - sin_rachat_tot_euro - sin_rachat_part_euro - 
                                                            sin_terme_euro)), pm_cloture_euro)
  }
  
  return(charg_pm_euro)
}



calcul_commission_pm_euro <- function(tx_com_sur_encours_euro, pm_euro_prev, interets_prime_euro,
                                      interets_fin_periode_euro, interets_pm_mi_periode_euro,
                                      pb_prime_euro, pb_fin_periode_euro, pb_mi_periode_euro,
                                      prime_nette_euro, cap_uc_euro, cap_euro_uc,
                                      charg_transf_euro_uc, sin_deces_euro, sin_tirage_euro,
                                      sin_rachat_tot_euro, sin_rachat_part_euro, sin_terme_euro) {
  
  commissions_pm_euro <- tx_com_sur_encours_euro * 
    (pm_euro_prev + interets_prime_euro + interets_fin_periode_euro + 
       interets_pm_mi_periode_euro + pb_prime_euro + pb_fin_periode_euro + 
       pb_mi_periode_euro + 0.5 * (prime_nette_euro + cap_uc_euro - 
                                     cap_euro_uc - charg_transf_euro_uc - sin_deces_euro - 
                                     sin_tirage_euro - sin_rachat_tot_euro - sin_rachat_part_euro - 
                                     sin_terme_euro))
  
  return(commissions_pm_euro)
}

calcul_pm_euro <- function(pm_cloture_euro, charg_pm_euro) {
  
  pm_euro <- max(0, pm_cloture_euro - charg_pm_euro)
  
  return(pm_euro)
}



calcul_pm <- function(pm_prev, params){
  
  pm_mi_periode_euro = f_pm_mi_periode_1(params, params[""])
  
  pm_mi_periode_uc = f_pm_mi_periode_1("uc")
  
  pm_mi_periode_euro = f_pm_mi_periode_2()
  
  pm_mi_periode_uc = f_pm_mi_periode_2()
  
  
  
}


d |> 
  





###########################################
#phase de test pour init et caluls préliminaires
###########################################

##nom_produit sexe client

test1 <- process_model_point_sample |> 
  distinct(num_adh, nom_produit, sexe_client)

test2 <- test |> 
  distinct(num_adh = num_adh_original, nom_produit = nom_prod_code, sexe_client = sexe_code)

setdiff(test2, test1)

setdiff(test1, test2)


# taux chargement pm

test1 <- process_model_point_sample |> 
  distinct(num_adh, compartiment, tx_charg_pm) |> 
  pivot_wider(names_from = compartiment, values_from = tx_charg_pm) |> 
  rename(taux_chargement_pm_euro = euro,
         taux_chargement_pm_uc = uc
         )

test2 <- test |> 
  distinct(num_adh = num_adh_original, taux_chargement_pm_euro, taux_chargement_pm_uc)


identical(test1, test2)


## bonus, annuite_avant_bonus = DurVerBonus, delai_bonus = DelaiBonus

test1 <- process_model_point_sample |> 
  distinct(num_adh, type_bonus, delai_bonus) |> 
  pivot_wider(names_from = type_bonus, values_from = delai_bonus) |> 
  rename(
    delai_bonus1_annee = `1`,
    delai_bonus2_annee = `2`
  ) |> 
  arrange(num_adh)


test2 <- test |> 
  distinct(num_adh = num_adh_original, delai_bonus1_annee, delai_bonus2_annee) |> 
  group_by(num_adh) |> 
  slice(1, .preserve = TRUE) |> 
  ungroup() |> 
  arrange(num_adh)

test2 == test1

identical(test1, test2)




## duree restant prime


test1 <- process_model_point_sample |> 
  filter(annee_projection > 1) |> 
  distinct(num_adh, duree_restante_prime)
  

test2 <- test |> 
  filter(annee_projection > 1) |> 
  distinct(num_adh = num_adh_original, duree_restante_prime = duree_restante_prime_annee)

identical(test1, test2)

setdiff(test1, test2)

setdiff(test2, test1)

## duree restante

test1 <- process_model_point_sample |> 
  filter(annee_projection > 0) |> 
  distinct(num_adh, duree_restante)

test2 <- test |> 
  filter(annee_projection > 0) |> 
  distinct(num_adh = num_adh_original, duree_restante = duree_restante_annee)

identical(test1, test2)


#age_assure, anciennete_contrat, reste_contrat, date_anniversaire

a <- process_model_point_sample |> 
  select(num_adh, annee_projection, age_assure, anciennete_contrat, reste_contrat, date_anniversaire) |> 
  distinct()

b <- test |> 
  select(num_adh = num_adh_original, annee_projection, age_assure = age_assure_annee, anciennete_contrat = anciennete_contrat_annee,
         reste_contrat = reste_contrat_annee, date_anniversaire = date_anniversaire_annee
         )

identical(a$age_assure, b$age_assure)

identical(a$anciennete_contrat, b$anciennete_contrat)

setdiff(a, b)

diff <- setdiff(b, a)


#nbtetes

a <- process_model_point_sample |> 
  select(num_adh, annee_projection, nb_tetes) |> 
  distinct()

b <- test |> 
  select(num_adh = num_adh_original, annee_projection, nb_tetes = nb_tetes_initial) |> 
  distinct()

identical(b, a)

a == b

#nb cloture 

a <- process_model_point_sample |> 
  select(num_adh, annee_projection, compartiment, nb_tetes_bis) |> 
  distinct() |> 
  pivot_wider(names_from = compartiment, values_from = nb_tetes_bis) |> 
  rename(
    nb_clotures_euro = euro,
    nb_clotures_uc = uc
  ) |> 
  filter(annee_projection == 0)

b <- test |> 
  select(num_adh = num_adh_original, annee_projection, nb_clotures_euro = nb_clotures_euro_annee, nb_clotures_uc = nb_clotures_uc_annee) |> 
  distinct() |> 
  filter(annee_projection == 0)

identical(a, b)

setdiff(a ,b)
setdiff(b, a)

p_1 <- runif(20, min = 0, max = 10)

p_2 <- runif(20, min = 0, max = 10)

v <- runif(20, min = 0, max = 10)


u <- 1


for (i in 1:20){
  u[i + 1] <- v[i] + u[i]*(1 - p_1[i] - p_2[i])
}


l <- v/(1 - (1 - p_1 - p_2))

test <- cumprod(1-p_1-p_2)*(1-l) + l

test <- cumprod(1-p_1-p_2) + cumsum(rev(cumprod(1-p_1-p_2))*v)


P_raw <- cumprod(1-p_1-p_2)
# P contient (P_0, P_1, ..., P_N)
P <- c(1, P_raw)
N <- length(1-p_1-p_2)
# Calculer b_k / P_{k+1}
# b est (b_0, ..., b_{N-1})
# P[2:(N+1)] est (P_1, ..., P_N)
# b_k / P_{k+1} correspond à b / P[2:(N+1)]
ratio <- v / P[2:(N + 1)]

# Calculer la somme cumulative (sum_{k=0}^{n-1} ...)
# On ajoute 0 au début car la somme est vide pour n=0
sum_ratio <- c(0, cumsum(ratio))

# Calculer U_n = P_n * (U0 + sum_ratio[n+1])
# P est (P_0, ..., P_N)
# sum_ratio est (0, sum_0, sum_0_1, ..., sum_0_{N-1})
# On a besoin de P[n+1] * (U0 + sum_ratio[n+1])
U0 <- 1
U_seq_formula <- P * (U0 + sum_ratio)

cumprod(1-p_1-p_2)*(1 + cumsum(v/cumprod(1-p_1-p_2)))


