library(tidyverse)
library(readxl)


model_point <- readxl::read_excel("data/modele_epargne.xlsm", sheet = "MODEL POINT", guess_max = 50000) |> 
  janitor::clean_names() |> 
  rename(tx_charg_pm_euro = tx_charg_sur_pm_euro) |> 
  pivot_longer(cols = ends_with(c("euro", "uc")), names_to = c(".value", "type"), names_pattern = "(.*)_(uc|euro)") |> 
  mutate(
    annee_naissance = year(date_de_naissance), .after = date_de_naissance
  ) |> 
  mutate(annee_echeance = year(date_echeance), .after = date_echeance) |> 
  mutate(nb_tetes = 1)
  


params <- list()
params$date_valorisation <- ymd("2024-12-31")
params$annee_valorisation <- lubridate::year(params$date_valorisation)
  
  
model_point <- readxl::read_excel("data/modele_epargne.xlsm", sheet = "MODEL POINT", guess_max = 50000) |> 
  janitor::clean_names() |> 
  rename(tx_charg_pm_euro = tx_charg_sur_pm_euro,
         annuites_avant_bonus_1 = annuites_avant_bonus1
         ) |> 
  pivot_longer(cols = ends_with(c("euro", "uc")), names_to = c(".value", "compartiment"), names_pattern = "(.*)_(uc|euro)") |> 
  pivot_longer(cols = ends_with(c("1", "2")), names_to = c(".value", "type_bonus"), names_pattern = "(.*)_(1|2)") |>
  mutate(
    annee_naissance = year(date_de_naissance), .after = date_de_naissance
  ) |> 
  mutate(annee_echeance = year(date_echeance), .after = date_echeance) |> 
  mutate(nb_tetes = 1) |>
  mutate(tx_charg_pm = if_else(position_contrat == "En réduction" & tx_charg_pm > 0, tx_charg_pm + tx_charg_suspens_pp, tx_charg_pm)) |> 
  mutate(delai_bonus = annuites_avant_bonus - (params$annee_valorisation - annee_effet) - 1,
         delai_bonus = case_when(
           periodicie == "Semestriel" & mois_
         )
         )
  





test <- read_excel("data/modele_epargne.xlsm", sheet = "HYPOTHESES", range = "A11:CX70", col_names = c("Produit", as.character(0:100)))

test2 <- as.matrix(test)





read_hypotheses <- function(debut, fin, nom_variable) {
  
  # 1. Construction de la plage de cellules à lire
  # Cette partie de votre code est correcte.
  # Elle crée une chaîne de caractères comme "A1:CX100" si debut=1 et fin=100.
  range_excel <- glue::glue("A{debut}:CX{fin}")
  
  # Message pour l'utilisateur indiquant la plage lue
  print(paste("Lecture de la plage :", range_excel))
  
  # 2. Lecture des données depuis le fichier Excel
  # Les données sont d'abord stockées dans une variable temporaire.
  donnees_matrice <- as.matrix(
    read_excel(
      "data/modele_epargne.xlsm", 
      sheet = "HYPOTHESES", 
      range = range_excel, 
      col_names = c("Produit", as.character(0:100))
    )
  )
  
  # 3. Assignation des données à une variable dont le nom est fourni par l'argument 'nom_variable'
  # La fonction assign() permet de faire cela.
  # Le premier argument est le nom que vous voulez donner à la variable (sous forme de chaîne de caractères).
  # Le deuxième argument est la valeur que vous voulez assigner à cette nouvelle variable.
  # Par défaut, la variable est créée dans l'environnement actuel (ici, l'environnement de la fonction).
  assign(nom_variable, donnees_matrice)
  
  # Message pour l'utilisateur indiquant la création de la variable
  print(paste("Variable '", nom_variable, "' créée avec succès.", sep=""))
  
  # 4. Construction du chemin de fichier pour la sauvegarde
  # On utilise la chaîne de caractères 'nom_variable' pour nommer le fichier .RData.
  chemin_fichier <- glue::glue("Hypotheses/{nom_variable}.RData")
  
  # 5. Sauvegarde de la variable dans un fichier .RData
  # Pour sauvegarder un objet dont le nom est contenu dans une chaîne de caractères,
  # il faut utiliser l'argument 'list' de la fonction save().
  # 'list' prend un vecteur de chaînes de caractères, chaque chaîne étant le nom d'un objet à sauvegarder.
  save(list = nom_variable, file = chemin_fichier)
  
  # Message pour l'utilisateur indiquant que la sauvegarde est terminée
  print(paste("La variable '", nom_variable, "' a été sauvegardée dans le fichier : '", chemin_fichier, "'", sep=""))
  
}


import_table <- function(range, colnames, name){
  
  matrix <- as.matrix(
    read_excel(
      "data/modele_epargne.xlsm", 
      sheet = "HYPOTHESES", 
      range = range, 
      col_names = colnames
    )
  )
  
  assign(name, matrix)
  
  path <- glue::glue("Hypotheses/{name}.RData")
  
  save(list = name, file = path)
}





TxTirage <- read_excel("data/modele_epargne.xlsm", sheet = "HYPOTHESES", range = 'A730:CX776')

colnames(TxTirage) <- c("produit", 0:100)

TxTirage[, 2] <- 0

TxTirage <- TxTirage |> 
  pivot_longer(cols = !produit, names_to = "anciennete", values_to = "tx_tirage")
  


corrige_na <- function(x, replace){
  x[is.na(x)] <- replace 
  x
}

read_hypotheses <- function(debut, fin, nom_variable) {
  
  range_excel <- glue::glue("A{debut}:CX{fin}")

  donnees <- read_excel("data/modele_epargne.xlsm", sheet = "HYPOTHESES", range = range_excel, col_names = c("produit", 0:100)) |> 
    pivot_longer(cols = !produit, names_to = "anciennete", values_to = nom_variable) |> 
    mutate(across(!produit, \(x) corrige_na(x, 0)))
    
  assign(nom_variable, donnees)

  chemin_fichier <- glue::glue("Hypotheses_new/{nom_variable}.RDS")

  saveRDS(get(nom_variable), file = chemin_fichier)
}


read_hypotheses(11, 56, "ChPrime_Prev")
read_hypotheses(83, 128, "ChDeces")
read_hypotheses(227, 272, "ChTirage")
read_hypotheses(227, 272, "ChRachatTot")
read_hypotheses(299, 344, "ChRachatPart")
read_hypotheses(371, 416, "ChPrestTerme")
read_hypotheses(443, 488, "ChPM_Prev")
read_hypotheses(515, 560, "ChArb_Euro_UC")
read_hypotheses(587, 632,"ChArb_UC_Euro")
read_hypotheses(659, 704,"ChArb_UC_UC")
read_hypotheses(731, 776,"TxTirage")
read_hypotheses(1019, 1064,"TxPass_Euro_UC")
read_hypotheses(1091, 1136,"TxPass_UC_Euro")
read_hypotheses(1163, 1208,"TxPass_UC_UC")
read_hypotheses(1235, 1280,"TxTechPre2011")
read_hypotheses(1307, 1352,"TxTechPost2011")


create_table <- function(path_dir){
  
  files <- list.files(path_dir, full.names = TRUE)
  
  files |> 
    map(readRDS) |> 
    accumulate(\(x, y) left_join(x, y)) |> 
    pluck(-1)
  
}


hypotheses <- create_table("Hypotheses_new") |> 
  mutate(anciennete = as.numeric(anciennete)) |> 
  janitor::clean_names()


read_hypotheses2 <- function(debut, fin, nom_variable) {
  
  range_excel <- glue::glue("A{debut}:CX{fin}")
  
  donnees <- read_excel("data/modele_epargne.xlsm", sheet = "HYPOTHESES", range = range_excel, col_names = c("categorie_rachat", 0:100)) |> 
    pivot_longer(cols = !categorie_rachat, names_to = "anciennete", values_to = nom_variable) |> 
    mutate(across(!categorie_rachat, \(x) corrige_na(x, 0)))
  
  assign(nom_variable, donnees)
  
  chemin_fichier <- glue::glue("Hypotheses_2/{nom_variable}.RDS")
  
  saveRDS(get(nom_variable), file = chemin_fichier)
}


read_hypotheses2(803, 807, "TxRachatTot")
read_hypotheses2(947, 951, "TxRachatPart")
read_hypotheses2(875, 879, "TxRachatTot_Prev")


hypotheses2 <- create_table("Hypotheses_2") |> 
  mutate(anciennete = as.numeric(anciennete)) |> 
  janitor::clean_names()

import_table("A1434:C1479", c("prorogation", "taux", "duree") , "TableProrog")

vecteurPB <- rep(0L, 101)
names(vecteurPB) <- 0:100

rendementUC <- rep(0L, 101)
names(rendementUC) <- 0:100

save(vecteurPB, file = "Hypotheses/vecteurPB.RData")


tf002 <- read_excel("data/TH-TF-00-02.xls", range = "A2:C115") |> 
  janitor::clean_names()
  
th002 <- read_excel("data/TH-TF-00-02.xls", range = "F2:G113") |> 
  janitor::clean_names() |> 
  mutate(age = 0:(n()-1), .before = 1)
  
thf_002 <- list("H" = th002, "F" = tf002) |> 
  list_rbind(names_to = "sexe")


save(model_point, file = "data/model_point.RData")


model_point |> 
  group_by(nom_produit) |> 
  summarise(montant_pm_euro = sum(pm_euro),
            montant_pm_uc = sum(pm_uc),
            taux_chargement_moyen = weighted.mean(tx_charg_sur_pm_euro, w = pm_euro),
            taux_commissions_moyen = weighted.mean(tx_com_sur_encours_euro, w = pm_euro),
            taux_chargement_moyen
            )


a <- model_point |> 
  rename(tx_charg_pm_euro = tx_charg_sur_pm_euro) |> 
  pivot_longer(cols = ends_with(c("euro", "uc")), names_to = c(".value", "type"), names_pattern = "(.*)_(uc|euro)")


a |> 
  group_by(nom_produit, type) |> 
  summarise(montant_pm = sum(pm),
            taux_chargement_encours_moyen = weighted.mean(tx_charg_pm, w = pm),
            taux_commissions_encours_moyen = weighted.mean(tx_com_sur_encours, w = pm),
            taux_chargement_prime_moyen = weighted.mean(tx_charg_prime, w = prime_commerciale),
            taux_commissions_prime_moyen = weighted.mean(taux_com_prime, w = prime_commerciale)
  )


a |> 
  group_by(type, taux_annuel_garanti) |> 
  summarise(montant_pm = sum(pm))



b <- a |> 
  filter(type == "euro", taux_annuel_garanti == 0) |> 
  group_by(indic_obseque) |> 
  summarise(montant_pm = sum(pm))



pb_uc <- tibble(annee_projection = 0:100, pb = 0, rendement_uc = 0)


library(highcharter)

highchart() |> 
  # Data
  hc_add_series(
    b,
    "pie",
    hcaes(
      name = indic_obseque,
      y = montant_pm
    ),
    name = "Bars"
  )



a |> 
  group_by(type, taux_annuel_garanti, indic_obseque) |> 
  summarise(somme_pm = sum(pm),
            age_moyen = weighted.mean(age, w = pm)
            )


a |> 
  group_by(type, periodicite, indic_obseque) |> 
  summarise(somme_pm = sum(pm),
            somme_prime = sum(prime_commerciale)
            )


test <- a |> 
  group_by(position_contrat, indic_obseque) |> 
  summarise(somme_pm = sum(pm)
  )


hchart(
  test,
  "column",
  hcaes(x = position_contrat, y = somme_pm, group = indic_obseque)
) |> 
  hc_yAxis(type = "logarithmic") |> 
  hc_plotOptions(column = list(stacking = "normal"))





a <- model_point |> 
  group_by(type, taux_annuel_garanti, indic_obseque) |> 
  summarise(somme_pm = sum(pm)) |> 
  mutate(indic_obseque = if_else(indic_obseque == 1, "Epargne - Obseque", "Epargne - Classique")) |> 
  pivot_wider(names_from = c(type, indic_obseque), values_from = somme_pm, values_fill = 0)






test1 <- function() {
  model_point |> 
    reactable(
      columns = list(
         "position_contrat" = colDef(name = "position")
      )
    )
}


test1()



test2 <- function(variable, nom) {
  model_point |> 
    reactable(
      columns = list(
        !!sym(variable) := colDef(name = nom)
      )
    )
}


test2("position_contrat", "position")



test2 <- function(variable, nom) {
  col_list <- setNames(list(colDef(name = nom)), variable)
  model_point |>
    reactable(
      columns = col_list
    )
}

# Exemple d'utilisation
test2("position_contrat", "position")



calcul_impot <- function(revenu, part) {
  
  quotient <- revenu/part
  
  case_when(
    between(quotient, 0, 10777) ~ 0,
    between(quotient, 10778, 27478) ~ revenu*0.11 - 1185.45*part,
    between(quotient, 27479, 78570) ~ revenu*0.3 - 6406.29*part,
    between(quotient, 78571, 168994) ~ revenu*0.41 - 15048.99*part,
     quotient > 168994 ~ revenu*0.41 - 21808.75*part,
  )
  
  
}

montant_rachat <- function(valeur, valeur_initiale, valeur_rachat){
  
  valeur - (valeur_initiale*valeur/valeur_rachat)
  
}

hchart(pm1, "column", hcaes_("{x}" := names(pm1)[1], "{y}" := names(pm1)[4], "{group}" := names(pm1)[3])) |>
  hc_yAxis(title = list(text = "Total PM")) |>
  hc_xAxis(title = list(text = "Type de valeur")) |>
  hc_plotOptions(column = list(stacking = "normal"))





f <- expr(f(x = 1, y = 2))

# Add a new argument
f$z <- 3
f
#> f(x = 1, y = 2, z = 3)

# Or remove an argument:
f[[2]] <- NULL
f
#> f(y = 2, z = 3)


expr(!!expr(x + y) / !!expr(y + y))



f_coeff <- function(mois, periode){
  
  nb_max <- 12 %/% periode 
  
  res <- ceiling((12 - mois + 1) / periode)
  
  return((nb_max - res)/nb_max)
  
}


map(1:12, \(x) f_coeff(x, 1))

map(1:12, \(x) f_coeff(x, 3))

map(1:12, \(x) f_coeff(x, 6))

map(1:12, \(x) f_coeff(x, 12))


