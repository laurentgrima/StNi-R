# Syrph the Net Interactive version R 
# En cours de développement (version 10/08/2020), à utiliser avec précaution !
# Par Laurent Grima (grima.laurent@yahoo.com), pour le Groupe d'études des invertébrés Armoricains (GRETIA)

library("tidyverse")
library("openxlsx") # permet de créer le fichier Excel

# Données STN 2020 (fournies)

STN_habitats_especes <- read_delim("", ";", escape_double = FALSE, trim_ws = TRUE)
STN_especes_pheno <- read_delim("", ";", escape_double = FALSE, trim_ws = TRUE)

# Données liées au site (voir note pour le format)
data_site_habitats <- read_delim(".../ETUDE_habitats.csv", ";", escape_double = FALSE, trim_ws = TRUE)
data_filtre_geographique <- read_delim(".../ETUDE_filtre_geo.csv", ";", escape_double = FALSE, trim_ws = TRUE)
data_site_espece <- read_delim(".../ETUDE_especes.csv", ";", escape_double = FALSE, trim_ws = TRUE)

# filtre par secteur
# data_site_espece<- data_site_espece %>% 
#   filter(lieu == "") 

setwd("...") # Choix du répertoire où sera crée le classeur

# Création et mise en forme du classeur Excel
classeur_excel <- createWorkbook()
couleurs_1 <- createStyle(bgFill = "#FEFECE")
couleurs_2 <- createStyle(bgFill = "#B4F6BA")
couleurs_3 <- createStyle(bgFill = "#A8C8F6")
couleurs_rdv <- createStyle(bgFill = "#34CC34")
couleurs_abs <- createStyle(bgFill = "#FF0000")
couleurs_inna <- createStyle(bgFill = "#0070C0")
tres_faible <- createStyle(bgFill = "#000000", fontColour = "#FFFFFF")
faible <- createStyle(bgFill = "#FF0000")
moyenne <- createStyle(bgFill = "#FF7900")
bonne <- createStyle(bgFill = "#FFFF00")
tres_bonne <- createStyle(bgFill = "#79FF00")
excellente <- createStyle(bgFill = "#00FF2F")
# 

data_site_espece <- data_site_espece[!grepl("sp.", data_site_espece$specie),]
liste_totale <- STN_especes_pheno$species
liste_geographique <- data_filtre_geographique$specie
liste_predites_site <-c()
data_site_espece <- unique(data_site_espece$specie)

# vérifier s'il y a des especes non reconnues entre les fichiers ETUDE et STN
erreurs_referencement_site_espece <- setdiff(data_site_espece, liste_totale)
erreurs_referencement_liste_geographique <- setdiff(data_filtre_geographique$specie, liste_totale)
  if (all(is.na(erreurs_referencement_liste_geographique)) & all(is.na(erreurs_referencement_site_espece))){
    print("correspondance espèces : OK")  
} else {
  print("attention ! Certaines espèces ne sont pas reconnues dans la base de données STN, attention aux synonymes :")
  print(erreurs_referencement_liste_geographique) 
  print(erreurs_referencement_site_espece)
} 

# corriger si il y a une ligne habitats avec que des NA

data_site_habitats <- data_site_habitats %>% drop_na(Macrohabitat_STN)

### ANALYSE ###
n=1
# BOUCLE POUR L'ANALYSE PAR HABITAT
for (i in 1 : (length(data_site_habitats$Macrohabitat_STN) + 1)) {
  # 1 : liste des especes de françe et filtre habitat
  if (n == (length(data_site_habitats$Macrohabitat_STN) + 1)){ # partie analyse générale
      
      sp_attendues <- cbind(liste_predites_site)
      sp_attendues <- cbind.data.frame("specie"=sp_attendues)
      sp_attendues$filtre_habitat <- c(1)
      sp_attendues <- sp_attendues %>% mutate_if(is.factor, as.character)
      names(sp_attendues)[1] <- "specie"
      df_liste_totale <- cbind(specie = STN_especes_pheno$species)
      df_sp_habitat <- merge(sp_attendues,df_liste_totale,by="specie", all = T)
      df_sp_habitat <- unique(df_sp_habitat)
      df_sp_habitat <- df_sp_habitat$filtre_habitat
      df_sp_habitat_geo <- as.data.frame(df_sp_habitat)
      df_sp_habitat_geo <- cbind(df_liste_totale,df_sp_habitat_geo)
      names(df_sp_habitat_geo)[1] <- "specie"
      names(df_sp_habitat_geo)[2] <- "sp_attendues"
      habitat <- as.character("Analyse générale")
      addWorksheet(classeur_excel, habitat)
  }
  if (n < (length(data_site_habitats$Macrohabitat_STN) + 1)) {
  habitat <- pull(data_site_habitats[i,1], Macrohabitat_STN)
  addWorksheet(classeur_excel, habitat) # Creation de la feuille avec OpenXLSX
  habitat <- as.character(habitat)
  habitat_supp <- as.vector(as.matrix(data_site_habitats[n,2:(ncol(data_site_habitats))]))
  habitat_supp <- na.omit(habitat_supp)
  filtre_habitat_indiv <- subset(STN_habitats_especes, select=habitat)


  # 1.5 prise en compte des habitats supplémentaires
  # La boucle récupère les données de chaque sous habitats présents dans le vecteur puis ils sont aggrégés dans une seule colonne
  filtre_habitat_supp <- data.frame(liste_totale)
  if (!all(is.na(habitat_supp))) {
    for (i in 1 : length(habitat_supp)) {
      # for (c in 2 : ncol(data_site_habitats)){
      habitat_supp_i <- as.character(habitat_supp)
      habitat_supp_i <- habitat_supp_i[i]
      filtre_habitat_supp[,i+1] <- subset(STN_habitats_especes, select=habitat_supp_i)
    # }}
    }
    nbrow<- as.numeric(length(habitat_supp))
    
    if (nbrow>1){
    filtre_ss_hab <-rowSums(filtre_habitat_supp[,2:nbrow], na.rm=T)
    filtre_ss_hab[filtre_ss_hab > 0] <- 1
    filtre_habitat_indiv <- cbind(filtre_habitat_indiv,filtre_ss_hab)
    names(filtre_habitat_indiv)[2] <- "filtre_ss_hab"

    }
    else{
      filtre_ss_hab <- filtre_habitat_supp
      filtre_habitat_indiv <-cbind(filtre_habitat_indiv,filtre_ss_hab)
      names(filtre_habitat_indiv)[2] <- "filtre_ss_hab"
    }
  } 
  

  
  specie <- subset(STN_habitats_especes, select="specie") # colonne espece (toutes)
  df_sp_habitat <- cbind(specie,filtre_habitat_indiv)
  names(df_sp_habitat)[2] <- "filtre_habitat"

  
  # affiche le nom de l'habitat dans le classeur
  if (n < length(data_site_habitats$Macrohabitat_STN)+1){
  writeData(classeur_excel, sheet = habitat, x = "habitat :", startCol = 6, startRow = 1)
  writeData(classeur_excel, sheet = habitat, x = habitat, startCol = 6, startRow = 2)
  writeData(classeur_excel, sheet = habitat, x = "habitat supp.:", startCol = 7, startRow = 1)
  writeData(classeur_excel, sheet = habitat, x = habitat_supp, startCol = 7, startRow = 2)
}
  # 2 : filtre geographique 
  
  
  
  filtre_geographique <- cbind.data.frame("specie"=liste_geographique)
  filtre_geographique$filtre_geo. <- c(1)
  filtre_geographique <- filtre_geographique %>% mutate_if(is.factor, as.character)
  df_sp_geo. <- merge(filtre_geographique,specie,by="specie", all = T)
  df_sp_geo. <- unique(df_sp_geo.)

  df_sp_habitat_geo <- cbind(df_sp_habitat,filtre_geographique=df_sp_geo.$filtre_geo.)

  
  # 3 : especes predites
  
  df_sp_habitat_geo$sp_attendues <- 0
  df_sp_habitat_geo$sp_attendues[df_sp_habitat_geo$filtre_habitat >= 1 & df_sp_habitat_geo$filtre_geographique >= 1] <- 1
  
  df_sp_habitat_geo[is.na(df_sp_habitat_geo)] <- 0
  # si il y a des ss
  if (!all(is.na(habitat_supp))){
    
    df_sp_habitat_geo$sp_attendues[df_sp_habitat_geo$filtre_habitat == 1 & df_sp_habitat_geo$filtre_ss_hab == 0] <- 0
  }else {
    df_sp_habitat_geo$sp_attendues[df_sp_habitat_geo$filtre_habitat == 1] <- 0
  }
  
}  # 4 : especes observees
  liste_sp_observees <- data_site_espece
  df_sp_habitat_geo<- df_sp_habitat_geo %>% mutate(sp_observees= ifelse(is.na(pmatch(liste_totale, liste_sp_observees, duplicates.ok = T)),0, pmatch(liste_totale, liste_sp_observees, duplicates.ok = T)))
  df_sp_habitat_geo$sp_observees[df_sp_habitat_geo$sp_observees >= 1] <- 1 # convertis les donnnes en binaire
  

  # replacer toutes les NA des deux premieres colonnes par des 0
  df_sp_habitat_geo[is.na(df_sp_habitat_geo)] <- 0
  
  
  df_sp_habitat_geo$aurdv <- 0
  df_sp_habitat_geo$absentes <- 0
  df_sp_habitat_geo$inattendues <- 0
  
  df_sp_habitat_geo$aurdv[df_sp_habitat_geo$sp_observees >= 1 & df_sp_habitat_geo$sp_attendues >= 1] <- 1
  df_sp_habitat_geo$absentes[df_sp_habitat_geo$sp_observees == 0 & df_sp_habitat_geo$sp_attendues >= 1] <- 1
  df_sp_habitat_geo$inattendues[df_sp_habitat_geo$sp_observees !=0 & df_sp_habitat_geo$sp_attendues == 0] <- 1

  # donnnes phéno # boucle pour les calculs pheno
  
   pheno <- STN_especes_pheno[14:234]
   liste_id <- c("absentes", "aurdv", "sp_attendues", "inattendues")


    nb_row=1
    df_stat_pheno_combine <-as.data.frame(c())
   for (i in 1 : 4){ # boucle pour chaque parametre
     id_stat <- liste_id[i]
     vecteur_pheno <-c()
   for (i in 1 : ncol(pheno)){ #boucle pour chaque colonne du tableau de pheno
     plage_pheno <- pheno[i]
     plage_pheno <- cbind(df_sp_habitat_geo$specie, as.numeric(df_sp_habitat_geo[[id_stat]]),plage_pheno)
     df_pheno <- plage_pheno[plage_pheno[2] > 0, ]
     df_pheno <- df_pheno[,3]
     stat <- as.numeric(length(df_pheno[!is.na(df_pheno)]))
     vecteur_pheno <- c(vecteur_pheno, stat)
     writeData(classeur_excel, sheet = habitat, x = stat, startCol = 10+i, startRow = nb_row, colNames = F)
   }
     
     # stocke les valeurs pour le calcul de l'intégrité dans un dataframe
     if (id_stat=="aurdv"){ 
       aurdv  <- vecteur_pheno
     }
     if (id_stat=="sp_attendues"){ 
       attendues <-vecteur_pheno
            }
     nb_row=nb_row+1
   }
    integrite <- round(aurdv/attendues*100, digits=2)
    integrite <-as.data.frame(t((integrite))) 
    writeData(classeur_excel, sheet = habitat, x = integrite, startCol = 11, startRow = 5, colNames = F)
    
   
# tableau phéno des sp obs
    
   
    plage_pheno <- cbind(df_sp_habitat_geo$specie, df_sp_habitat_geo$sp_attendues, df_sp_habitat_geo$sp_observees,pheno) # mettre les espèces observées
    df_pheno <- plage_pheno[plage_pheno[2] | plage_pheno[3] > 0, ]
    names(df_pheno)[1] <- "specie"
    df_liste_totale <- cbind(specie = liste_totale)
    df_pheno <- merge(df_liste_totale,df_pheno,by="specie", all = TRUE)
    
    # filtre enlever les especes ni prédites ni observées
    df_pheno<- df_pheno %>% 
      filter(df_sp_habitat_geo$sp_attendues > 0 | df_sp_habitat_geo$sp_observees > 0) 
    df_pheno <- df_pheno[,-2]
    df_pheno <- df_pheno[,-2]
    
    writeData(classeur_excel, sheet = habitat, x = df_pheno, startCol = 10, startRow = 8)

  # arrangements graphiques
  
    setColWidths(classeur_excel, habitat, cols = c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10), widths = c(30, 10, 10, 10, 10, 10, 10, 10, 10, 30))

    # 1 à 3 codage liste
    
  conditionalFormatting(classeur_excel, habitat, cols = 2,
                        rows = 6:1000, rule = "=1", style = couleurs_1,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 2,
                        rows = 6:1000, rule = "=2", style = couleurs_2,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 2,
                        rows = 6:1000, rule = "=3", style = couleurs_3,
                        type = "expression")
  # 1 à 3 pat habitat
  
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 6:1000, rule = "=1", style = couleurs_1,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 6:1000, rule = "=2", style = couleurs_2,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 6:1000, rule = "=3", style = couleurs_3,
                        type = "expression")
  
   # sp inattendues, attendues, absentes
  
  conditionalFormatting(classeur_excel, habitat, cols = 6,
                        rows = 6:1000, rule = "=1", style = couleurs_rdv,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 7,
                        rows = 6:1000, rule = "=1", style = couleurs_abs,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 8,
                        rows = 6:1000, rule = "=1", style = couleurs_inna,
                        type = "expression")

  
  # etat_integrite
  
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 5, rule = ">=0", style = tres_faible,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 5, rule = ">21", style = faible,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 5, rule = ">41", style = moyenne,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 5, rule = ">51", style = bonne,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 5, rule = ">76", style = tres_bonne,
                        type = "expression")
  conditionalFormatting(classeur_excel, habitat, cols = 11:230,
                        rows = 5, rule = ">86", style = excellente,
                        type = "expression")
  
  # Tableau résultats bilan
  intitules <-   c("Prédites", "Observées", "Au rendez-vous", "Manquantes", "Inattendues")
  valeur <- c(nrow(filter(df_sp_habitat_geo, sp_attendues >= 1)),nrow(filter(df_sp_habitat_geo, sp_observees >= 1)), sum(df_sp_habitat_geo$aurdv == '1'),sum(df_sp_habitat_geo$absentes == '1'), sum(df_sp_habitat_geo$inattendues == '1'))
  df_recap_compte <- cbind(intitules,valeur)
  
  writeData(classeur_excel, sheet = habitat, x = df_recap_compte, startCol = 3, startRow = 1, colNames = F)
  writeData(classeur_excel, sheet = habitat, x = intitules, startCol = 2,colNames = F)

  
  rdv_predites <- round(as.numeric(df_recap_compte[3,2])/as.numeric(df_recap_compte[1,2])*100, digits = 2)
  manq_predites <- round(as.numeric(df_recap_compte[4,2])/as.numeric(df_recap_compte[1,2])*100, digits = 2)
  rdv_obs <- round(as.numeric(df_recap_compte[3,2])/as.numeric(df_recap_compte[2,2])*100, digits = 2)
  inna_obs <- round(as.numeric(df_recap_compte[5,2])/as.numeric(df_recap_compte[2,2])*100, digits = 2)
  
   # Tableau stats bilan
   intitules_2 <-   c("Intégrité écologique", "% rdv/prédites", "% manq/prédites", "Qualité du modèle", "% rdv/observées","% inattend/obs")
   valeur_2 <- c("",rdv_predites,manq_predites,"",rdv_obs,inna_obs)
   df_recap_compte_2 <- cbind(intitules_2,valeur_2)
   writeData(classeur_excel, sheet = habitat, x = df_recap_compte_2, startCol = 1, startRow = 1,colNames = F)

   # filtre enlever les especes ni prédites ni observées
   df_sp_habitat_geo<- df_sp_habitat_geo %>% 
     filter(sp_attendues == 1 | sp_observees == 1) 
   
   writeDataTable(classeur_excel, sheet = habitat,colNames = TRUE, withFilter = TRUE, x = df_sp_habitat_geo, startCol = 1, startRow = 8)
  
  
  liste_predites <- df_sp_habitat_geo$specie[df_sp_habitat_geo$sp_attendues >= 1]
  liste_predites_site <-c(liste_predites,liste_predites_site)
  liste_predites_site <- unique(liste_predites_site)
  n=n+1
  
  cat(n-1) # Affiche le numéro de l'habitat traité
  }
print("Tous les habitats ont été analysés")
saveWorkbook(classeur_excel, file = "classeur_excel.xlsx", overwrite = TRUE) # sauvegarde du classeur
