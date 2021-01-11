library("tidyverse")
library("openxlsx") # permet de créer le fichier Excel
classeur_excel <- createWorkbook()
setwd("C:/Users/Grima/Desktop/GRETIA STN/PLOUNERIN/resultats")
STN_habitats_especes <- read_delim("STN-habitats-especes.csv", ";", escape_double = FALSE, trim_ws = TRUE)
STN_especes_pheno <- read_delim("STN-especes-phenologies.csv", ";", escape_double = FALSE, trim_ws = TRUE)

# Données liées au site (voir note pour le format)
data_site_habitats <- read_delim("PLOUNERIN_habitats.csv", ";", escape_double = FALSE, trim_ws = TRUE)
data_filtre_geographique <- read_delim("PLOUNERIN_filtre_bretagne.csv", ";", escape_double = FALSE, trim_ws = TRUE)
data_site_espece <- read_delim("PLOUNERIN_syrphidae.csv", ";", escape_double = FALSE, trim_ws = TRUE)

erreurs_liste_sp <- setdiff(data_site_espece$specie, STN_habitats_especes$specie)
erreurs_filtre <- setdiff(data_filtre_geographique$specie, STN_habitats_especes$specie)

data_site_espece <- data_site_espece %>%
  filter(!(specie %in% erreurs_liste_sp))

data_filtre_geographique <- data_filtre_geographique %>%
  filter(!(specie %in% erreurs_filtre))
remove(erreurs_filtre)
remove(erreurs_liste_sp)

# Liste macrohabitats
liste_macrohabitats <- unique(na.omit(data_site_habitats$Macrohabitat_STN))


# BOUCLE POUR L'ANALYSE PAR HABITAT
for (i in 1 : length(liste_macrohabitats)) {

addWorksheet(classeur_excel, liste_macrohabitats[i])
df_filtres <- as.data.frame(cbind(specie = STN_especes_pheno$specie,filtre_geographique = NA))


df_filtres <- df_filtres %>% mutate(filtre_geographique = NA) 
df_filtres <- df_filtres %>% 
  rowwise() %>% 
  mutate(filtre_geographique = ifelse((any(specie == data_filtre_geographique$specie)),"1","0"))%>%
  ungroup()


macrohabitat <- data_site_habitats$Macrohabitat_STN[i]
df_filtres <- cbind(df_filtres, macrohabitat = subset(STN_habitats_especes, select = as.character(macrohabitat)))

microhabitats <- (as.vector(as.matrix(data_site_habitats[i,2:(ncol(data_site_habitats))])))


if (!is.na(microhabitats[1])){
  microhabitats <- na.omit(microhabitats)
  for (i in 1 : length(microhabitats)) {
    df_filtres <- cbind(df_filtres,microhabitat=subset(STN_habitats_especes, select=microhabitats[i]))
  }
  nb_microhabitats <- as.numeric(length(microhabitats))
} else{
  nb_microhabitats <- 0
}




df_filtres <- df_filtres %>% 
  mutate(filtre_habitat = case_when(rowSums(as.data.frame(df_filtres[,3:(3+nb_microhabitats)]),na.rm = TRUE) > 1 & df_filtres[,3]==1 ~ "1",df_filtres[,3]>1 ~ as.character(df_filtres[,3])))


df_filtres$attendues[df_filtres$filtre_geographique >= 1 & df_filtres$filtre_habitat >= 1] <- 1


df_filtres <- df_filtres %>% 
  rowwise() %>% 
  mutate(observees = ifelse((any(specie == data_site_espece$specie)),"1",NA))%>%
  ungroup()


df_filtres$aurdv <- NA
df_filtres$absentes <- NA
df_filtres$inattendues <- NA
df_filtres$aurdv[df_filtres$observees >= 1 & df_filtres$attendues >= 1] <- 1
df_filtres$absentes[is.na(df_filtres$observees) & df_filtres$attendues >= 1] <- 1
df_filtres$inattendues[df_filtres$observees >= 1 & is.na(df_filtres$attendues) >= 1] <- 1


## Enlever les espèces ni attendues ni observées

df_filtres<- df_filtres %>% 
  filter(attendues == 1 | observees == 1)


## Calculs


nb_attendues <- length(na.omit(df_filtres$attendues))
nb_observees <- length(na.omit(df_filtres$observees))
nb_aurdv <- length(na.omit(df_filtres$aurdv))
nb_absentes <- length(na.omit(df_filtres$absentes))
nb_inattendues <- length(na.omit(df_filtres$inattendues))

integrite <- round(length(na.omit(df_filtres$aurdv))/length(na.omit(df_filtres$attendues))*100, digits=2)
qualite_modele <- round(length(na.omit(df_filtres$aurdv))/length(na.omit(df_filtres$observees))*100, digits=2)
bilan1 <- as.data.frame(cbind(nb_attendues,nb_observees,nb_aurdv,nb_absentes,nb_inattendues))
bilan2 <- as.data.frame(cbind(integrite,qualite_modele))

##
traits <- cbind(STN_especes_pheno$specie,STN_especes_pheno[14:234])
colnames(traits)[1] = "specie"
traits <- left_join(df_filtres[,1:1],traits,by=c("specie"), all = TRUE)

##
couleurs_1 <- createStyle(bgFill = "#FEFECE")
couleurs_2 <- createStyle(bgFill = "#B4F6BA")
couleurs_3 <- createStyle(bgFill = "#A8C8F6")
couleurs_rdv <- createStyle(bgFill = "#34CC34")
couleurs_abs <- createStyle(bgFill = "#FF0000")
couleurs_inna <- createStyle(bgFill = "#0070C0")
# 1 à 3 codage liste

conditionalFormatting(classeur_excel, as.factor(macrohabitat), cols = 3,
                      rows = 6:1000, rule = "=1", style = couleurs_1,
                      type = "expression")
conditionalFormatting(classeur_excel, as.factor(macrohabitat), cols = 3,
                      rows = 6:1000, rule = "=2", style = couleurs_2,
                      type = "expression")
conditionalFormatting(classeur_excel, as.factor(macrohabitat), cols = 3,
                      rows = 6:1000, rule = "=3", style = couleurs_3,
                      type = "expression")

# sp inattendues, attendues, absentes

conditionalFormatting(classeur_excel, as.factor(macrohabitat), cols = (7+length(microhabitats)),
                      rows = 6:1000, rule = "=1", style = couleurs_rdv,
                      type = "expression")
conditionalFormatting(classeur_excel, as.factor(macrohabitat), cols = (8+length(microhabitats)),
                      rows = 6:1000, rule = "=1", style = couleurs_abs,
                      type = "expression")
conditionalFormatting(classeur_excel, as.factor(macrohabitat), cols = (9+length(microhabitats)),
                      rows = 6:1000, rule = "=1", style = couleurs_inna,
                      type = "expression")
##
writeData(classeur_excel, sheet = as.factor(macrohabitat), x = as.factor(macrohabitat), startCol = 7, startRow = 1, colNames = T)
writeData(classeur_excel, sheet = as.factor(macrohabitat), x = as.factor(microhabitats), startCol = 8, startRow = 1, colNames = T)
writeData(classeur_excel, sheet = as.factor(macrohabitat), x = df_filtres, startCol = 1, startRow = 6, colNames = T, withFilter = TRUE)
writeData(classeur_excel, sheet = as.factor(macrohabitat), x = traits, startCol = (ncol(df_filtres)+2), startRow = 6, colNames = T)
writeData(classeur_excel, sheet = as.factor(macrohabitat), x = bilan1, startCol = 1, startRow = 1, colNames = T)
writeData(classeur_excel, sheet = as.factor(macrohabitat), x = bilan2, startCol = 1, startRow = 3, colNames = T)

remove(df_filtres,microhabitats,macrohabitat,nb_attendues,nb_observees,nb_aurdv,nb_absentes,nb_inattendues,integrite,qualite_modele)
}
print("Tous les habitats ont été analysés")
saveWorkbook(classeur_excel, file = "PlounerinSTNv2.xlsx", overwrite = TRUE) # sauvegarde du classeur
