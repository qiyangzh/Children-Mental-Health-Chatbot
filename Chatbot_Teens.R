########################################################################################################
#Chatbot Teens Mental Health Interventions
########################################################################################################
# Authors: Qiyang Zhang
# Contact: qiyang39@nus.edu.sg
# Created: 2025/06/30

########################################################################################################
# Initial Set-up
########################################################################################################
# Clear workspace
rm(list=ls(all=TRUE))

# Load packages
test<-require(googledrive)   #all gs_XXX() functions for reading data from Google
if (test == FALSE) {
  install.packages("googledrive")
  require(googledrive)
}
test<-require(googlesheets4)   #all gs_XXX() functions for reading data from Google
if (test == FALSE) {
  install.packages("googlesheets4")
  require(googlesheets4)
}
test<-require(plyr)   #rename()
if (test == FALSE) {
  install.packages("plyr")
  require(plyr)
}
test<-require(metafor)   #escalc(); rma();
if (test == FALSE) {
  install.packages("metafor")
  require(metafor)
}
test<-require(robumeta)
if (test == FALSE) {
  install.packages("robumeta")
  require(robumeta)
}
test<-require(weightr) #selection modeling
if (test == FALSE) {
  install.packages("weightr")
  require(weightr)
}
test<-require(clubSandwich) #coeftest
if (test == FALSE) {
  install.packages("clubSandwich")
  require(clubSandwich)
}
test<-require(tableone)   #CreateTableOne()
if (test == FALSE) {
  install.packages("tableone")
  require(tableone)
}
test<-require(flextable)   
if (test == FALSE) {
  install.packages("flextable")
  require(flextable)
}
test<-require(officer)   
if (test == FALSE) {
  install.packages("officer")
  require(officer)
}
test<-require(tidyverse)   
if (test == FALSE) {
  install.packages("tidyverse")
  require(tidyverse)
}
test<-require(ggrepel)   
if (test == FALSE) {
  install.packages("ggrepel")
  require(ggrepel)
}
test<-require(readxl)   
if (test == FALSE) {
  install.packages("readxl")
  require(readxl)
}
test<-require(robumeta)   
if (test == FALSE) {
  install.packages("robumeta")
  require(robumeta)
}
test<-require(dplyr)   
if (test == FALSE) {
  install.packages("dplyr")
  require(dplyr)
}
test<-require(meta)   
if (test == FALSE) {
  install.packages("meta")
  require(meta)
}
rm(test)

########################################################################################################
# Load data
########################################################################################################
# set up to load from Google
drive_auth(email = "zhangqiyang0329@gmail.com")
id <- drive_find(pattern = "Adolescents_Chatbot_Well_being", type = "spreadsheet")$id[1]

# load findings and studies
gs4_auth(email = "zhangqiyang0329@gmail.com")
findings <- read_sheet(id, sheet = "Findings", col_types = "c")
studies <- read_sheet(id, sheet = "Studies", col_types = "c")   # includes separate effect sizes for each finding from a study
rm(id)

########################################################################################################
# Clean data
########################################################################################################
# remove any empty rows & columns
studies <- subset(studies, is.na(studies$Study)==FALSE)
findings <- subset(findings, is.na(findings$Study)==FALSE)

studies <- subset(studies, is.na(studies$Drop)==TRUE)
findings <- subset(findings, is.na(findings$Drop) == TRUE)

# merge dataframes
full <- merge(studies, findings, by = c("Study"), all = TRUE, suffixes = c(".s", ".f"))
full <- subset(full, is.na(full$Authors.s)!=TRUE)

# format to correct variable types
nums <- c("Treatment.N.original", "Control.N.original",
          "T_Mean_Pre", "T_SD_Pre", "C_Mean_Pre", 
          "C_SD_Pre", "T_Mean_Post", "T_SD_Post", 
          "C_Mean_Post", "C_SD_Post", "Clustered",
          "Treatment.Cluster","Control.Cluster",
          "Age.mean","Duration.weeks", "Follow-up",
          "Students", "Clinical", "Negative Outcomes",
          "Sample size", "Female", "HumanAssistance")

full[nums] <- lapply(full[nums], as.numeric)
rm(nums)

###############################################################
#Create unique identifiers (ES, study, program)
###############################################################
full$ESId <- as.numeric(rownames(full))
full$StudyID <- as.numeric(as.factor(full$Study))
summary(full$StudyID)
########################################################################################################
# Prep data
########################################################################################################
##### Calculate ESs #####
# calculate pretest ES, SMD is standardized mean difference
full <- escalc(measure = "SMD", m1i = T_Mean_Pre, sd1i = T_SD_Pre, n1i = Treatment.N.original,
               m2i = C_Mean_Pre, sd2i = C_SD_Pre, n2i = Control.N.original, data = full)
full$vi <- NULL
full <- plyr::rename(full, c("yi" = "ES_Pre"))

# calculate posttest ES
full <- escalc(measure = "SMD", m1i = T_Mean_Post, sd1i = T_SD_Post, n1i = Treatment.N.original,
               m2i = C_Mean_Post, sd2i = C_SD_Post, n2i = Control.N.original, data = full)
full$vi <- NULL
full <- plyr::rename(full, c("yi" = "ES_Post"))

# calculate DID (post ES - pre ES)
full$ES_DID <- full$ES_Post - full$ES_Pre

# put various ES together.  Options:
## 1) used reported ES (so it should be in the Effect.Size column, nothing to do)
## 2) Effect.Size is NA, and DID not missing, replace with that
full$ES_DID[which(is.na(full$ES_DID)==TRUE & is.na(full$ES)==FALSE)] <- full$ES[which(is.na(full$ES_DID)==TRUE & is.na(full$ES)==FALSE)]
full$Effect.Size <- as.numeric(full$ES_DID)
full$Effect.Size[full$Negative.Outcomes == 1 & !is.na(full$Negative.Outcomes)] <-
  full$Effect.Size[full$Negative.Outcomes == 1 & !is.na(full$Negative.Outcomes)] * -1
#View(full[c("Study","Negative.Outcomes", "Control", "Outcomes","Effect.Size", "ES_DID", "ES_Pre", "ES_Post")])

###############################################################
#Calculate meta-analytic variables: Correct ES for clustering (Hedges, 2007, Eq.19)
###############################################################
#first, create an assumed ICC
full$icc <- NA
full$icc[which(full$Clustered == 1)] <- 0.2

#find average students/cluster
full$Treatment.Cluster.n <- NA
full$Treatment.Cluster.n[which(full$Clustered == 1)] <- round(full$Treatment.N.original[which(full$Clustered == 1)]/full$Treatment.Cluster[which(full$Clustered == 1)], 0)
full$Control.Cluster.n <- NA
full$Control.Cluster.n[which(full$Clustered == 1)] <- round(full$Control.N.original[which(full$Clustered == 1)]/full$Control.Cluster[which(full$Clustered == 1)], 0)
# find other parts of equation
full$n.TU <- NA
full$n.TU <- ((full$Treatment.N.original * full$Treatment.N.original) - (full$Treatment.Cluster * full$Treatment.Cluster.n * full$Treatment.Cluster.n))/(full$Treatment.N.original * (full$Treatment.Cluster - 1))
full$n.CU <- NA
full$n.CU <- ((full$Control.N.original * full$Control.N.original) - (full$Control.Cluster * full$Control.Cluster.n * full$Control.Cluster.n))/(full$Control.N.original * (full$Control.Cluster - 1))

# next, calculate adjusted ES, save the originals, then replace only the clustered ones
full$Effect.Size.adj <- full$Effect.Size * (sqrt(1-full$icc*(((full$Sample.size-full$n.TU*full$Treatment.Cluster - full$n.CU*full$Control.Cluster)+full$n.TU + full$n.CU - 2)/(full$Sample.size-2))))

# save originals, replace for clustered
full$Effect.Size.orig <- full$Effect.Size
full$Effect.Size[which(full$Clustered==1)] <- full$Effect.Size.adj[which(full$Clustered==1)]

#reverse all effect sizes
full$Effect.Size[which(full$Clustered==1)] <- full$Effect.Size.adj[which(full$Clustered==1)]
num <- c("Effect.Size")
full[num] <- lapply(full[num], as.numeric)
rm(num)

###############################################################
#Calculate meta-analytic variables: Sample sizes
###############################################################
#create full sample/total clusters variables
full$Sample <- full$Treatment.N.original + full$Control.N.original

################################################################
# Calculate meta-analytic variables: Variances (Lipsey & Wilson, 2000, Eq. 3.23)
################################################################
#calculate standard errors
full$se<-sqrt(((full$Treatment.N.original+full$Control.N.original)/(full$Treatment.N.original*full$Control.N.original))+((full$Effect.Size*full$Effect.Size)/(2*(full$Treatment.N.original+full$Control.N.original))))

#calculate variance
full$var<-full$se*full$se

########################################
#meta-regression
########################################
full <- subset(full, is.na(full$Drop.s)==TRUE)

#Centering, when there is missing value, this won't work
full$FemalePercent <- full$Female/full$Sample.size
full$FiftyPercentFemale <- 0
full$FiftyPercentFemale[which(full$`FemalePercent` > 0.50)] <- 1

full$Clustered.c <- full$Clustered - mean(full$Clustered)
full$Follow.up.c <- full$Follow.up - mean(full$Follow.up)
full$Clinical.c <- full$Clinical - mean(full$Clinical)
full$FiftyPercentFemale.c <- full$FiftyPercentFemale - mean(full$FiftyPercentFemale)
full$HumanAssistance.c <- full$HumanAssistance - mean(full$HumanAssistance)

full$ESId <- as.numeric(rownames(full))
full$StudyID <- as.numeric(as.factor(full$Study))
summary(full$StudyID)


###########################
#forest plot
###########################
by_study_outcome <- full %>%
  dplyr::group_by(Study, Outcomes) %>%
  dplyr::summarise(
    Effect.Size = mean(Effect.Size, na.rm = TRUE),
    var = mean(var, na.rm = TRUE),
    n_rows = n()  # how many rows were averaged
  )

fit_HKSJ_SJ <- metafor::rma(
  yi     = Effect.Size,
  vi     = var,
  data = by_study_outcome,
  method = "SJ",        # Sidik–Jonkman tau^2
  test   = "knha"       # Hartung–Knapp with HA adjustment
)
# Create concise study labels (Study — Outcome)
by_study_outcome$slab <- paste0(by_study_outcome$Study, " — ", by_study_outcome$Outcomes)

forest(
  fit_HKSJ_SJ,
  slab       = by_study_outcome$slab,
  xlab       = "Effect size",
  refline    = 0,          # change to 1 for ratio measures on log scale
  showweights = TRUE,
  cex        = 0.9
)

# Add overall summary polygon (at the bottom)
addpoly(fit_HKSJ_SJ, row = -1, mlab = "Overall (RE, SJ + HK-Knha)")
pred <- predict(fit_HKSJ_SJ, digits = 3)  # pred$pi.lb, pred$pi.ub
title(sub = sprintf("Prediction interval: [%.3f, %.3f]", pred$pi.lb, pred$pi.ub))

########################################
#meta-regression
########################################
#Null Model
V_list <- impute_covariance_matrix(vi=full$var, cluster=full$StudyID, r=0.8)

MVnull <- rma.mv(yi=Effect.Size,
                 V=V_list,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full,
                 method="REML")
MVnull

#t-test of each covariate#
MVnull.coef <- coef_test(MVnull, cluster=full$StudyID, vcov="CR2")
MVnull.coef

#sensitivity test
# yi_uni, vi_uni should be one effect and its sampling variance per study
by_study_outcome <- full %>%
  dplyr::group_by(Study, Outcomes) %>%
  dplyr::summarise(
    Effect.Size = mean(Effect.Size, na.rm = TRUE),
    var = mean(var, na.rm = TRUE),
    n_rows = n()  # how many rows were averaged
  )
fit_HKSJ_SJ <- metafor::rma(
  yi     = Effect.Size,
  vi     = var,
  data = by_study_outcome,
  method = "SJ",        # Sidik–Jonkman tau^2
  test   = "knha"       # Hartung–Knapp with HA adjustment
)
summary(fit_HKSJ_SJ)
pred_HKSJ <- predict(fit_HKSJ_SJ)
pred_HKSJ

#Positive Outcome Model
Positive <- subset(full, full$Positive.Outcomes==1)

V_listPositive <- impute_covariance_matrix(vi=Positive$var, cluster=Positive$StudyID, r=0.8)

MVnullPositive <- rma.mv(yi=Effect.Size,
                         V=V_listPositive,
                         random=~1 | StudyID/ESId,
                         test="t",
                         data=Positive,
                         method="REML")
MVnullPositive

# #t-test of each covariate#
MVnull.coefPositive <- coef_test(MVnullPositive, cluster=Positive$StudyID, vcov="CR2")
MVnull.coefPositive

#sensitivity test
by_study_outcome_Positive <- Positive %>%
  dplyr::group_by(Study, Outcomes) %>%
  dplyr::summarise(
    Effect.Size = mean(Effect.Size, na.rm = TRUE),
    var = mean(var, na.rm = TRUE),
    n_rows = n()  # how many rows were averaged
  )
fit_HKSJ_SJ_Positive <- metafor::rma(
  yi     = Effect.Size,
  vi     = var,
  data = by_study_outcome_Positive,
  method = "SJ",        # Sidik–Jonkman tau^2
  test   = "knha"       # Hartung–Knapp with HA adjustment
)
summary(fit_HKSJ_SJ_Positive)
pred_HKSJ_Positive <- predict(fit_HKSJ_SJ_Positive)
pred_HKSJ_Positive

#Negative Outcome Model
Negative <- subset(full, full$Negative.Outcomes==1)

V_listNegative <- impute_covariance_matrix(vi=Negative$var, cluster=Negative$StudyID, r=0.8)

MVnullNegative <- rma.mv(yi=Effect.Size,
                         V=V_listNegative,
                         random=~1 | StudyID/ESId,
                         test="t",
                         data=Negative,
                         method="REML")
MVnullNegative

# #t-test of each covariate#
MVnull.coefNegative <- coef_test(MVnullNegative, cluster=Negative$StudyID, vcov="CR2")
MVnull.coefNegative

#sensitivity test
by_study_outcome_Negative <- Negative %>%
  dplyr::group_by(Study, Outcomes) %>%
  dplyr::summarise(
    Effect.Size = mean(Effect.Size, na.rm = TRUE),
    var = mean(var, na.rm = TRUE),
    n_rows = n()  # how many rows were averaged
  )
fit_HKSJ_SJ_Negative <- metafor::rma(
  yi     = Effect.Size,
  vi     = var,
  data = by_study_outcome_Negative,
  method = "SJ",        # Sidik–Jonkman tau^2
  test   = "knha"       # Hartung–Knapp with HA adjustment
)
summary(fit_HKSJ_SJ_Negative)
pred_HKSJ_Negative <- predict(fit_HKSJ_SJ_Negative)
pred_HKSJ_Negative
### Output prediction interval ###
dat_clean <- data.frame(yi = Negative$Effect.Size, se_g = Negative$se)
dat_clean <- na.omit(dat_clean)

yi_clean <- dat_clean$yi
se_g_clean <- dat_clean$se_g
install.packages("pimeta")
library(pimeta)

pima_result <- pima(yi_clean, se_g_clean, method = "HK")  # Using the Hartung-Knapp method

print(pima_result)


#Method: "ActiveorPassive"
#Unit:"Clinical", "Age", "FiftyPercentFemale"
#Treatment:"Duration.weeks", "Personalized", "Self.guided", "Modality","Social.function"
#Outcome:"Outcomes"
#Setting: "WEIRD"

# terms <- c("ActiveorPassive")
# terms <- c("Clinical", "Age", "FiftyPercentFemale")
# terms <- c("Duration.weeks", "Personalized", "HumanAssistance", "Modality","Social.function")
# terms <- c("Outcomes")
# terms <- c("WEIRD")

#Single model: not significant:"FiftyPercentFemale.c", "Negative.Outcomes", "Recruitment.setting","Clinical", "ActiveorPassive",
#not significant: "Targeted", "Sample.size", "Culture", "HumanAssistance", "Clinical", "Duration.weeks","Social.function", "Response.generation.approach", "WEIRD"
#marginally significant
terms <- c("Embodied")
terms <- c("Age")
#not significant
terms <- c("Age.mean")
terms <- c("Modality")
terms <- c("Age", "Embodied", "Modality")

#interact <- c("Social.function*Age")
#formula <- reformulate(termlabels = c(terms, interact))
formula <- reformulate(termlabels = c(terms))
formula

MVfull <- rma.mv(yi=Effect.Size,
                 V=V_list,
                 mods=formula,
                 random=~1 | StudyID/ESId,
                 test="t",
                 data=full,
                 method="REML")
MVfull
#t-test of each covariate#
MVfull.coef <- coef_test(MVfull, cluster=full$StudyID, vcov="CR2")
MVfull.coef

#######
#embodied
#######
# 1) Join Embodied (two categories) from full into by_study_outcome
by_study_outcome <- by_study_outcome %>%
  left_join(full %>% distinct(Study, Embodied), by = "Study")

# 2) Recode Embodied to a factor with explicit levels and reference
# Normalize common variants, then set levels exactly as requested
by_study_outcome <- by_study_outcome %>%
  mutate(
    Embodied = tolower(trimws(as.character(Embodied))),
    Embodied = case_when(
      Embodied %in% c("yes","y","true","1","embodied","avatar","robot") ~ "Yes",
      Embodied %in% c("no","n","false","0","non-embodied","none")       ~ "No",
      TRUE ~ NA_character_
    ),
    Embodied = factor(Embodied, levels = c("No", "Yes"))  # reference = No
  )

# 3) Build analysis datasets
dat_emb_all <- subset(by_study_outcome, !is.na(Embodied) & !is.na(Effect.Size) & !is.na(var))
stopifnot(nrow(dat_emb_all) > 0)

# If you have OutcomeDomain (recommended), use that to split:
has_outcome_domain <- "OutcomeDomain" %in% names(by_study_outcome)
if (has_outcome_domain) {
  dat_emb_neg <- subset(dat_emb_all, OutcomeDomain == "negative mental health")
  dat_emb_pos <- subset(dat_emb_all, OutcomeDomain == "positive well-being")
} else {
  # Fallback: use Negative.Outcomes indicator if present (1 = negative, 0 = positive)
  stopifnot("Negative.Outcomes" %in% names(by_study_outcome))
  dat_emb_neg <- subset(dat_emb_all, as.numeric(as.character(Negative.Outcomes)) == 1)
  dat_emb_pos <- subset(dat_emb_all, as.numeric(as.character(Negative.Outcomes)) == 0)
}

# 4) Fit three single-moderator models (overall, negative, positive)

# 4a) Overall (all outcomes)
fit_Emb_all <- rma(
  yi     = dat_emb_all$Effect.Size,
  vi     = dat_emb_all$var,
  mods   = ~ Embodied,     # coefficient = Yes vs No (reference)
  method = "SJ",           # Sidik–Jonkman tau^2
  test   = "knha",         # Hartung–Knapp with HA adjustment
  data   = dat_emb_all
)
summary(fit_Emb_all)

# CR2 robust SEs clustered by Study
cluster_vec_all <- dat_emb_all$Study[fit_Emb_all$not.na]
rob_Emb_all <- coef_test(fit_Emb_all,
                         cluster = cluster_vec_all,
                         vcov    = "CR2",
                         test    = "Satterthwaite")
rob_Emb_all

#########
#Modality
########
# 1) Join Modality from full into by_study_outcome (study-level dataset assumed already built)
by_study_outcome <- by_study_outcome %>%
  dplyr::left_join(full %>% dplyr::distinct(Study, Modality), by = "Study")

# 2) Recode Modality to a factor with explicit levels and reference = "text"
# Normalize common variants to the three target labels: "text", "audio", "multimedia"
by_study_outcome <- by_study_outcome %>%
  dplyr::mutate(
    Modality = tolower(trimws(as.character(Modality))),
    Modality = dplyr::case_when(
      Modality %in% c("text", "chat", "messaging", "sms", "web-text") ~ "text",
      Modality %in% c("audio", "voice", "phone", "ivr", "smart-speaker") ~ "audio",
      Modality %in% c("multimedia", "multimodal", "text+voice", "voice+text",
                      "app+chat", "avatar+text", "text+audio") ~ "multimedia",
      TRUE ~ NA_character_
    ),
    Modality = factor(Modality, levels = c("text", "audio", "multimedia"))
  )

# 3) Build analysis datasets
dat_mod_all <- subset(by_study_outcome, !is.na(Modality) & !is.na(Effect.Size) & !is.na(var))
stopifnot(nrow(dat_mod_all) > 0)

# If you have OutcomeDomain, use it to split; else fall back to Negative.Outcomes
if ("OutcomeDomain" %in% names(by_study_outcome)) {
  dat_mod_neg <- subset(dat_mod_all, OutcomeDomain == "negative mental health")
  dat_mod_pos <- subset(dat_mod_all, OutcomeDomain == "positive well-being")
} else {
  stopifnot("Negative.Outcomes" %in% names(by_study_outcome))
  # Ensure Negative.Outcomes is numeric (0/1) for filtering
  dat_mod_all <- dat_mod_all %>%
    dplyr::mutate(Negative.Outcomes = as.numeric(as.character(Negative.Outcomes)))
  dat_mod_neg <- subset(dat_mod_all, Negative.Outcomes == 1)
  dat_mod_pos <- subset(dat_mod_all, Negative.Outcomes == 0)
}

# 4) Fit three single-moderator models (overall, negative, positive)
#    Coefficients interpret differences vs reference level "text"

# 4a) Overall (all outcomes)
fit_Mod_all <- metafor::rma(
  yi     = dat_mod_all$Effect.Size,
  vi     = dat_mod_all$var,
  mods   = ~ Modality,    # audio vs text; multimedia vs text
  method = "SJ",
  test   = "knha",
  data   = dat_mod_all
)
summary(fit_Mod_all)

# CR2 robust SEs clustered by Study
cluster_vec_all <- dat_mod_all$Study[fit_Mod_all$not.na]
rob_Mod_all <- clubSandwich::coef_test(fit_Mod_all,
                                       cluster = cluster_vec_all,
                                       vcov    = "CR2",
                                       test    = "Satterthwaite")
rob_Mod_all
###########################
#forest plot
###########################
study_averages <- full %>%
  group_by(StudyID) %>%
  summarise(avg_effect_size = mean(Effect.Size, na.rm = TRUE),
            across(everything(), ~ first(.)))

MVnull <- robu(formula = Effect.Size ~ 1, studynum = StudyID, data = study_averages, var.eff.size = var)
m.gen <- metagen(TE = Effect.Size,
                 seTE = se,
                 studlab = Study,
                 data = study_averages,
                 sm = "SMD",
                 fixed = FALSE,
                 random = TRUE,
                 method.tau = "REML",
                 method.random.ci = "HK")
summary(m.gen)
png(file = "forestplot.png", width = 2800, height = 3000, res = 300)

meta::forest(m.gen,
             sortvar = TE,
             prediction = TRUE,
             print.tau2 = FALSE,
             leftlabs = c("Author", "g", "SE"))
dev.off()
#################################################################################
# Marginal Means
#################################################################################
# re-run model for each moderator to get marginal means for each #

# set up table to store results
means <- data.frame(moderator = character(0), group = character(0), beta = numeric(0), SE = numeric(0), 
                    tstat = numeric(0), df = numeric(0), p_Satt = numeric(0))

mods <- c("as.factor(Modality)", "as.factor(Social.function)", 
          "as.factor(Clinical)", "as.factor(ActiveorPassive)", "as.factor(WEIRD)",
          "as.factor(Age)", "as.factor(FiftyPercentFemale)", "as.factor(HumanAssistance)",
          "as.factor(Outcomes)")

for(i in 1:length(mods)){
  # i <- 1
  formula <- reformulate(termlabels = c(mods[i], terms, "-1"))   # Worth knowing - if you duplicate terms, it keeps the first one
  mod_means <- rma.mv(yi=Effect.Size, #effect size
                      V = V_list, #variance (tHIS IS WHAt CHANGES FROM HEmodel)
                      mods = formula, #ADD COVS HERE
                      random = ~1 | StudyID/ESId, #nesting structure
                      test= "t", #use t-tests
                      data=full, #define data
                      method="REML") #estimate variances using REML
  coef_mod_means <- as.data.frame(coef_test(mod_means,#estimation model above
                                            cluster=full$StudyID, #define cluster IDs
                                            vcov = "CR2")) #estimation method (CR2 is best)
  # limit to relevant rows (the means you are interested in)
  coef_mod_means$moderator <- gsub(x = mods[i], pattern = "as.factor", replacement = "")
  coef_mod_means$group <- rownames(coef_mod_means)
  rownames(coef_mod_means) <- c()
  coef_mod_means <- subset(coef_mod_means, substr(start = 1, stop = nchar(mods[i]), x = coef_mod_means$group)== mods[i])
  coef_mod_means$group <- substr(x = coef_mod_means$group, start = nchar(mods[i])+1, stop = nchar(coef_mod_means$group))
  means <- dplyr::bind_rows(means, coef_mod_means)
}
means
#################################################################################
# Heterogeneity
#################################################################################
# 95% prediction intervals
print(PI_upper <- MVfull$b[1] + (1.96*sqrt(MVfull$sigma2[1] + MVfull$sigma2[2])))
print(PI_lower <- MVfull$b[1] - (1.96*sqrt(MVfull$sigma2[1] + MVfull$sigma2[2])))

#################################################################################
#Create Descriptives Table
#################################################################################
# identify variables for descriptive tables (study-level and outcome-level)
#"Follow.up.Duration.weeks","Age.mean","FemalePercent",Duration.weeks",
vars_study <- c("WEIRD",
                "Clinical", "Age", "FiftyPercentFemale",
                "Personalized", "HumanAssistance", 
                "Modality","Social.function")
vars_outcome <- c("Outcomes", "Clustered", "ActiveorPassive", "Follow.up")

# To make this work, you will need a df that is at the study-level for study-level 
# variables (such as research design) you may have already created this (see above, with study-level ESs), but if you didn't, here is an easy way:
# 1) make df with *only* the study-level variables of interest and studyIDs in it
study_level_full <- full[c("StudyID", "WEIRD",
                           "Clinical", "Age", "FiftyPercentFemale",
                            "HumanAssistance", 
                           "Modality","Social.function")]
# 2) remove duplicated rows
study_level_full <- unique(study_level_full)
# 3) make sure it is the correct number of rows (should be same number of studies you have)
length(study_level_full$StudyID)==length(unique(study_level_full$StudyID))
# don't skip step 3 - depending on your data structure, some moderators can be
# study-level in one review, but outcome-level in another

# create the table "chunks"
table_study_df <- as.data.frame(print(CreateTableOne(vars = vars_study, data = study_level_full, 
                                                     includeNA = TRUE, 
                                                     factorVars = c("WEIRD",
                                                                    "Clinical", "Age", "FiftyPercentFemale",
                                                                    "HumanAssistance", 
                                                                    "Modality","Social.function")), 
                                      showAllLevels = TRUE))
table_outcome_df <- as.data.frame(print(CreateTableOne(vars = vars_outcome, data = full, includeNA = TRUE,
                                                       factorVars = c("Outcomes", "Clustered", "ActiveorPassive", "Follow.up")), 
                                        showAllLevels = TRUE))
rm(vars_study, vars_outcome)

################################
# Descriptives Table Formatting
################################
table_study_df$Category <- row.names(table_study_df)
rownames(table_study_df) <- c()
table_study_df <- table_study_df[c("Category", "level", "Overall")]
table_study_df$Category[which(substr(table_study_df$Category, 1, 1)=="X")] <- NA
table_study_df$Category <- gsub(pattern = "\\..mean..SD..", replacement = "", x = table_study_df$Category)
table_study_df$Overall <- gsub(pattern = "\\( ", replacement = "\\(", x = table_study_df$Overall)
table_study_df$Overall <- gsub(pattern = "\\) ", replacement = "\\)", x = table_study_df$Overall)
table_study_df$Category <- gsub(pattern = "\\.", replacement = "", x = table_study_df$Category)
table_study_df$Category[which(table_study_df$Category=="n")] <- "Total Studies"
table_study_df$level[which(table_study_df$level=="1")] <- "Yes"
table_study_df$level[which(table_study_df$level=="0")] <- "No                                                                              "
# fill in blank columns (to improve merged cells later)
for(i in 1:length(table_study_df$Category)) {
  if(is.na(table_study_df$Category[i])) {
    table_study_df$Category[i] <- table_study_df$Category[i-1]
  }
}

table_outcome_df$Category <- row.names(table_outcome_df)
rownames(table_outcome_df) <- c()
table_outcome_df <- table_outcome_df[c("Category", "level", "Overall")]
table_outcome_df$Category[which(substr(table_outcome_df$Category, 1, 1)=="X")] <- NA
table_outcome_df$Overall <- gsub(pattern = "\\( ", replacement = "\\(", x = table_outcome_df$Overall)
table_outcome_df$Overall <- gsub(pattern = "\\) ", replacement = "\\)", x = table_outcome_df$Overall)
table_outcome_df$Category <- gsub(pattern = "\\.", replacement = "", x = table_outcome_df$Category)
table_outcome_df$Category[which(table_outcome_df$Category=="n")] <- "Total Effect Sizes"
# fill in blank columns (to improve merged cells later)
for(i in 1:length(table_outcome_df$Category)) {
  if(is.na(table_outcome_df$Category[i])) {
    table_outcome_df$Category[i] <- table_outcome_df$Category[i-1]
  }
}

########################
#Output officer
########################
myreport<-read_docx()
# Descriptive Table
myreport <- body_add_par(x = myreport, value = "Table 4: Descriptive Statistics", style = "Normal")
descriptives_study <- flextable(head(table_study_df, n=nrow(table_study_df)))
descriptives_study <- add_header_lines(descriptives_study, values = c("Study Level"), top = FALSE)
descriptives_study <- theme_vanilla(descriptives_study)
descriptives_study <- merge_v(descriptives_study, j = c("Category"))
myreport <- body_add_flextable(x = myreport, descriptives_study)

descriptives_outcome <- flextable(head(table_outcome_df, n=nrow(table_outcome_df)))
descriptives_outcome <- delete_part(descriptives_outcome, part = "header")
descriptives_outcome <- add_header_lines(descriptives_outcome, values = c("Outcome Level"))
descriptives_outcome <- theme_vanilla(descriptives_outcome)
descriptives_outcome <- merge_v(descriptives_outcome, j = c("Category"))
myreport <- body_add_flextable(x = myreport, descriptives_outcome)
myreport <- body_add_par(x = myreport, value = "", style = "Normal")

########################
# MetaRegression Table
########################
MVnull.coef
str(MVnull.coef)
MVnull.coef$coef <- row.names(as.data.frame(MVnull.coef))
row.names(MVnull.coef) <- c()
MVnull.coef <- MVnull.coef[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]
MVnull.coef
str(MVnull.coef)

MVfull.coef$coef <- row.names(as.data.frame(MVfull.coef))
row.names(MVfull.coef) <- c()
MVfull.coef <- MVfull.coef[c("Coef", "beta", "SE", "tstat", "df_Satt", "p_Satt")]

# MetaRegression Table
model_null <- flextable(head(MVnull.coef, n=nrow(MVnull.coef)))
colkeys <- c("beta", "SE", "tstat", "df_Satt")
model_null <- colformat_double(model_null,  j = colkeys, digits = 2)
model_null <- colformat_double(model_null,  j = c("p_Satt"), digits = 3)
#model_null <- autofit(model_null)
model_null <- add_header_lines(model_null, values = c("Null Model"), top = FALSE)
model_null <- theme_vanilla(model_null)

myreport <- body_add_par(x = myreport, value = "Table 5: Model Results", style = "Normal")
myreport <- body_add_flextable(x = myreport, model_null)
#myreport <- body_add_par(x = myreport, value = "", style = "Normal")

model_full <- flextable(head(MVfull.coef, n=nrow(MVfull.coef)))
model_full <- colformat_double(model_full,  j = c("beta"), digits = 2)
model_full <- colformat_double(model_full,  j = c("p_Satt"), digits = 3)
#model_full <- autofit(model_full)
model_full <- delete_part(model_full, part = "header")
model_full <- add_header_lines(model_full, values = c("Meta-Regression"))
model_full <- theme_vanilla(model_full)

myreport <- body_add_flextable(x = myreport, model_full)
myreport <- body_add_par(x = myreport, value = "", style = "Normal")

# Marginal Means Table
marginalmeans <- flextable(head(means, n=nrow(means)))
colkeys <- c("moderator", "group", "SE", "tstat", "df")
marginalmeans <- colformat_double(marginalmeans,  j = colkeys, digits = 2)
marginalmeans <- colformat_double(marginalmeans,  j = c("p_Satt"), digits = 3)
rm(colkeys)
marginalmeans <- theme_vanilla(marginalmeans)
marginalmeans <- merge_v(marginalmeans, j = c("moderator"))
myreport <- body_add_par(x = myreport, value = "Table: Marginal Means", style = "Normal")
myreport <- body_add_flextable(x = myreport, marginalmeans)

# Write to word doc
file = paste("TableResults.docx", sep = "")
print(myreport, file)


#publication bias , selection modeling
full_y <- full$Effect.Size
full_v <- full$var
weightfunct(full_y, full_v)
weightfunct(full_y, full_v, steps = c(.025, .50, 1))

# source files for creating each visualization type
#source("/Users/apple/Desktop/SchoolbasedMentalHealth/Scripts/viz_MARC.R")

median(full$Sample.size)
# Sum sample size for unique StudyID
sum_unique <- full %>%
  distinct(StudyID, .keep_all = TRUE) %>%
  summarise(total_sample = sum(Sample.size, na.rm = TRUE))

print(sum_unique)
# set sample sizes based on median sample size (331) and existing % weights for bar plot
#            w_j = meta-analytic weight for study j (before rescaling)
#            w_j_perc = percent weight allocated to study j

full <- full %>% 
  mutate(w_j = 1/(se^2)) %>%
  mutate(w_j_perc = w_j/sum(w_j)) %>%
  mutate(N_j = floor(w_j_perc*331))

#needed b/c tibbles throw error in viz_forest
full <- full %>% 
  ungroup()
full <- as.data.frame(full)

sd(full$Age.mean, na.rm=TRUE)
sd(full$Sample.size, na.rm=TRUE)
sd(full$Duration.weeks, na.rm=TRUE)

MVfull.modeling <- rma(yi=Effect.Size,
                       vi=var,
                       test="t",
                       data=full,
                       slab = Study,
                       method="REML")

#funnel plot
metafor::funnel(MVfull.modeling) 
sel <- selmodel(MVfull.modeling, type = "stepfun", steps = 0.025)
plot(sel, ylim=c(0,5))

jpeg(file="countour_funnel_color.jpeg")

# contour-enhanced funnel plot
metafor::funnel(MVfull.modeling, level=c(90, 95, 99), shade=c("white", "gray55", "gray75"), refline=0, legend=FALSE)
dev.off()


##########
#heatmap for outcomes
#########
# Ensure variables are factors (set Social.function order: task-oriented left, social-oriented right)
full <- full %>%
  mutate(
    Negative.Outcomes = as.numeric(Negative.Outcomes),
    OutcomeDomain = factor(
      Negative.Outcomes,
      levels = c(1, 0),
      labels = c("Negative mental health", "Positive well-being")
    ),
    Study = as.factor(Study)
  )

by_study_sf <- full %>%
  dplyr::group_by(Study, OutcomeDomain) %>%
  dplyr::summarise(
    `Published Year`   = first(as.integer(`Published.Year.s`)),
    `Chatbot.Name`  = first(`Chatbot.Name`),
    Effect.Size        = mean(Effect.Size, na.rm = TRUE),
    var                = mean(var, na.rm = TRUE),
    n_rows             = n(),                # how many rows were averaged
    .groups = "drop"
  )

# Build ordered Study levels (newest first; use rev(...) if you want oldest on top)
study_levels <- by_study_sf %>%
  dplyr::distinct(Study, `Published Year`) %>%
  dplyr::arrange(dplyr::desc(`Published Year`), Study) %>%
  dplyr::pull(Study)

# Apply the order to Study; also ensure Social.function order (task-oriented left, social-oriented right)
by_study_sf <- by_study_sf %>%
  mutate(
    Study = factor(Study, levels = study_levels),
    OutcomeDomain = fct_relevel(OutcomeDomain, "Negative mental health", "Positive well-being")
  )

# Plot
ggplot(by_study_sf, aes(x = OutcomeDomain, y = Study, fill = Effect.Size)) +
  geom_tile(color = "white", linewidth = 0.3) +
  scale_fill_gradient2(
    low = "#C53030", mid = "#F2F2F2", high = "#2B6CB0", midpoint = 0,
    name = "Effect Size"
  ) +
  labs(
    x = "Outcomes Types",
    y = "Study",
    title = "Heatmap of Effect Sizes by Outcome Types"
  ) +
  geom_text(aes(label = sprintf("%.2f", Effect.Size)), size = 5.5) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid = element_blank(),
    axis.text.x = element_text(size = 14, angle = 45, hjust = 1),
    axis.title.x = element_text(size = 16, face = "bold"),
    axis.text.y = element_text(size = 14),         # increase y-axis tick labels
    axis.title.y = element_text(size = 16, face = "bold")
  )

##########
#heatmap for Embodied
#########
# Ensure variables are factors (set Social.function order: task-oriented left, social-oriented right)
full <- full %>%
  mutate(
    Embodied = factor(
      Embodied,
      levels = c("Yes", "No"),
      labels = c("Embodied", "Not Embodied")
    ),
    Study = as.factor(Study)
  )

by_study_sf <- full %>%
  dplyr::group_by(Study, Embodied) %>%
  dplyr::summarise(
    `Published Year`   = first(as.integer(`Published.Year.s`)),
    `Chatbot.Name`  = first(`Chatbot.Name`),
    Effect.Size        = mean(Effect.Size, na.rm = TRUE),
    var                = mean(var, na.rm = TRUE),
    n_rows             = n(),                # how many rows were averaged
    .groups = "drop"
  )

# Build ordered Study levels (newest first; use rev(...) if you want oldest on top)
study_levels <- by_study_sf %>%
  dplyr::distinct(Study, `Published Year`) %>%
  dplyr::arrange(dplyr::desc(`Published Year`), Study) %>%
  dplyr::pull(Study)

# Apply the order to Study; also ensure Social.function order (task-oriented left, social-oriented right)
by_study_sf <- by_study_sf %>%
  mutate(
    Study = factor(Study, levels = study_levels),
    Embodied = fct_relevel(Embodied, "Embodied", "Not Embodied")
  )

# Plot
ggplot(by_study_sf, aes(x = Embodied, y = Study, fill = Effect.Size)) +
  geom_tile(color = "white", linewidth = 0.3) +
  scale_fill_gradient2(
    low = "#C53030", mid = "#F2F2F2", high = "#2B6CB0", midpoint = 0,
    name = "Effect Size"
  ) +
  labs(
    x = "Embodied",
    y = "Study",
    title = "Heatmap of Effect Sizes by Embodied"
  ) +
  geom_text(aes(label = sprintf("%.2f", Effect.Size)), size = 5.5) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid = element_blank(),
    axis.text.x = element_text(size = 14, angle = 45, hjust = 1),
    axis.title.x = element_text(size = 16, face = "bold"),
    axis.text.y = element_text(size = 14),         # increase y-axis tick labels
    axis.title.y = element_text(size = 16, face = "bold")
  )
