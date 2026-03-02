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

##############
#sensitivity analysis
#############
# sensitiv <- full <- subset(full, Study != "Rossi et al., 2020")
# V_listsensitiv <- impute_covariance_matrix(vi=sensitiv$var, cluster=sensitiv$StudyID, r=0.8)
# 
# MVnullsensitiv <- rma.mv(yi=Effect.Size,
#                  V=V_listsensitiv,
#                  random=~1 | StudyID/ESId,
#                  test="t",
#                  data=sensitiv,
#                  method="REML")
# MVnullsensitiv
# 
# #t-test of each covariate#
# MVnull.coefsensitiv <- coef_test(MVnullsensitiv, cluster=sensitiv$StudyID, vcov="CR2")
# MVnull.coefsensitiv

# #Negative Outcome Model
# Negativesensitiv <- subset(sensitiv, sensitiv$Negative.Outcomes==1)
# 
# V_listNegativesensitiv <- impute_covariance_matrix(vi=Negativesensitiv$var, cluster=Negativesensitiv$StudyID, r=0.8)
# 
# MVnullNegativesensitiv <- rma.mv(yi=Effect.Size,
#                          V=V_listNegativesensitiv,
#                          random=~1 | StudyID/ESId,
#                          test="t",
#                          data=Negativesensitiv,
#                          method="REML")
# MVnullNegativesensitiv
# 
# # #t-test of each covariate#
# MVnull.coefNegativesensitiv <- coef_test(MVnullNegativesensitiv, cluster=Negativesensitiv$StudyID, vcov="CR2")
# MVnull.coefNegativesensitiv
#Children Outcome Model
Children <- subset(full, full$Age...28== "childhood")

V_listChildren <- impute_covariance_matrix(vi=Children$var, cluster=Children$StudyID, r=0.8)

MVnullChildren <- rma.mv(yi=Effect.Size,
                         V=V_listChildren,
                         random=~1 | StudyID/ESId,
                         test="t",
                         data=Children,
                         method="REML")
MVnullChildren

# #t-test of each covariate#
MVnull.coefChildren <- coef_test(MVnullChildren, cluster=Children$StudyID, vcov="CR2")
MVnull.coefChildren

#Adolescents Outcome Model
Adolescents <- subset(full, full$Age...28== "adolescence")

V_listAdolescents <- impute_covariance_matrix(vi=Adolescents$var, cluster=Adolescents$StudyID, r=0.8)

MVnullAdolescents <- rma.mv(yi=Effect.Size,
                         V=V_listAdolescents,
                         random=~1 | StudyID/ESId,
                         test="t",
                         data=Adolescents,
                         method="REML")
MVnullAdolescents

# #t-test of each covariate#
MVnull.coefAdolescents <- coef_test(MVnullAdolescents, cluster=Adolescents$StudyID, vcov="CR2")
MVnull.coefAdolescents

#Method: "ActiveorPassive"
#Unit:"Clinical", "Age", "FiftyPercentFemale"
#Treatment:"Duration.weeks", "Personalized", "Self.guided", "Modality","Social.function"
#Outcome:"Outcomes"
#Setting: "WEIRD"
# terms <- c("ActiveorPassive")
# terms <- c("Clinical")
# terms <- c("FiftyPercentFemale")
# terms <- c("Duration.weeks")
# terms <- c("Social.function")
# terms <- c("Recruitment.setting")
# terms <- c("Western")
# terms <- c("AI.based")
# terms <- c("Response.generation.approach")

# #yes
# terms <- c("Modality")
# terms <- c("HumanAssistance")
# terms <- c("Embodied")
# terms <- c("Age...28")
#Single model: not significant:"FiftyPercentFemale.c", "Negative.Outcomes", "Recruitment.setting","Clinical", "ActiveorPassive",
#not significant: "Targeted", "Sample.size", "Culture", "HumanAssistance", "Clinical", "Duration.weeks","Social.function", "Response.generation.approach", "WEIRD"
#marginally significant,  "HumanAssistance"
#View(full[c("Study","Age...28", "Modality", "Negative.Outcomes","StudyID", "ESId", "var", "Control", "Outcomes","Effect.Size", "ES_DID", "ES_Pre", "ES_Post")])
V_list <- impute_covariance_matrix(vi=full$var, cluster=full$StudyID, r=0.8)

terms <- c("Age...28", "Modality")
#terms <- c("Embodied", "Age...28")
#terms <- c("Modality", "Embodied", "Age...28")
#terms <- c("Modality", "HumanAssistance", "Embodied", "Age...28")
#interact <- c("Age...28*Modality")
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

#check for collinearity
dummy_model <- lm(Effect.Size ~ Embodied + Age...28, data = full)
alias(dummy_model)
library(car)
vif(dummy_model)

#This VIF output confirms the diagnosis: You have severe multicollinearity.
#Almost all embodied chatbots are audio, and non-embodied are multimedia or text
#A Generalized VIF (GVIF) of 8.11 is well above the safety threshold (typically 5). This confirms that Embodied and Modality are highly redundant—they are essentially fighting to explain the exact same variance in your data.

#check for collinearity
dummy_model <- lm(Effect.Size ~ Embodied + Modality, data = full)
alias(dummy_model)
library(car)
vif(dummy_model)
#A VIF of 2.75 falls comfortably into the "Moderate" category

###########################
#forest plot
###########################
study_averages <- full %>%
  group_by(StudyID) %>%
  summarise(
    avg_effect_size = mean(Effect.Size, na.rm = TRUE),
    across(everything(), ~ first(.))
  ) %>%
  ungroup() %>%
  arrange(Study)   # ← alphabetical order

#MVnull <- robu(formula = Effect.Size ~ 1, studynum = StudyID, data = study_averages, var.eff.size = var)
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
             sortvar = NULL, 
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

mods <- c("as.factor(Modality)", "as.factor(Embodied)", "as.factor(Social.function)", 
          "as.factor(Clinical)", "as.factor(ActiveorPassive)", "as.factor(Western)",
          "as.factor(Age...28)", "as.factor(FiftyPercentFemale)", "as.factor(HumanAssistance)",
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
#################################################################################
# 95% prediction intervals
print(PI_upper <- MVfull$b[1] + (1.96*sqrt(MVfull$sigma2[1] + MVfull$sigma2[2])))
print(PI_lower <- MVfull$b[1] - (1.96*sqrt(MVfull$sigma2[1] + MVfull$sigma2[2])))

#################################################################################
#Create Descriptives Table
#################################################################################
# identify variables for descriptive tables (study-level and outcome-level)
#"Follow.up.Duration.weeks","Age.mean","FemalePercent",Duration.weeks",
vars_study <- c("Western",
                "Clinical", "Age...28", "FiftyPercentFemale",
                 "HumanAssistance", 
                "Modality","Social.function")
vars_outcome <- c("Outcomes", "Clustered", "ActiveorPassive", "Follow.up")

# To make this work, you will need a df that is at the study-level for study-level 
# variables (such as research design) you may have already created this (see above, with study-level ESs), but if you didn't, here is an easy way:
# 1) make df with *only* the study-level variables of interest and studyIDs in it
study_level_full <- full[c("StudyID", "Western",
                           "Clinical", "Age...28", "FiftyPercentFemale",
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
                                                     factorVars = c("Western",
                                                                    "Clinical", "Age...28", "FiftyPercentFemale",
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
tiff("funnel_plot.tiff",
     units = "in", width = 6.5, height = 6.0, res = 600,
     compression = "lzw")

par(mar = c(4.5, 4.5, 1.5, 1), cex = 1.1)  # margins + slightly larger text

metafor::funnel(MVfull.modeling,
                level  = c(90, 95, 99),
                shade  = c("white", "gray55", "gray75"),
                refline = 0,
                legend = FALSE)

dev.off()

##########
#heatmap for outcomes
#########
by_study_sf <- by_study_sf %>%
  mutate(Study = as.character(Study)) %>%
  arrange(`Published Year`, Study) %>%   # sort by year first
  mutate(Study = factor(Study, levels = rev(unique(Study))))

# ---- Save as high-resolution TIFF (wider, less tall) ----
tiff("heatmap_outcome.tiff",
     units = "in",
     width = 12,
     height = 6.5,
     res = 600,
     type = "cairo",
     compression = "lzw")

outcome <- ggplot(by_study_sf,
                  aes(x = OutcomeDomain,
                      y = Study,
                      fill = Effect.Size)) +
  geom_tile(color = "white", linewidth = 0.3) +
  scale_fill_gradient2(
    low = "#C53030",
    mid = "#F2F2F2",
    high = "#2B6CB0",
    midpoint = 0,
    name = "Effect Size"
  ) +
  labs(
    x = "Outcome Types",
    y = "Study",
    title = "Heatmap of Effect Sizes by Outcome Types"
  ) +
  geom_text(aes(label = sprintf("%.2f", Effect.Size)))+ 
  theme_minimal(base_size = 12) + 
  theme(panel.grid = element_blank(), axis.text.x = element_text(size = 14, angle = 0, hjust = 1), axis.title.x = element_text(size = 16, face = "bold"), axis.text.y = element_text(size = 14)) # increase y-axis tick labels axis.title.y = element_text(size = 16, face = "bold") )
print(outcome)
dev.off()

##########
#heatmap for Age
#########
# Reorder Study chronologically (oldest at top via rev(unique(...)) after arrange)
by_study_sf <- full %>%
  mutate(
    Study = as.character(Study),
    Age...28 = as.factor(Age...28),
    PublishedYear = as.integer(Published.Year.s)   # make a clean numeric year
  ) %>%
  group_by(Study, Age...28) %>%
  summarise(
    PublishedYear = first(PublishedYear),
    Chatbot.Name  = first(Chatbot.Name),
    Effect.Size   = mean(Effect.Size, na.rm = TRUE),
    var           = mean(var, na.rm = TRUE),
    n_rows        = n(),
    .groups = "drop"
  ) %>%
  arrange(PublishedYear, Study) %>%                       # Old → New
  mutate(Study = factor(Study, levels = rev(unique(Study))))  # Oldest on TOP


# Make Age...28 a factor for discrete x-axis ordering (adjust levels as you prefer)
by_study_sf <- by_study_sf %>%
  mutate(
    Age...28 = as.character(Age...28),
    Age...28 = factor(Age...28, levels = sort(unique(Age...28)))
  )

by_study_sf <- by_study_sf %>%
  mutate(
    Age...28 = fct_recode(Age...28,
                          "Childhood" = "childhood", "Adolescence" = "adolescence")
  )

# ---- Save as high-resolution TIFF (wider, less tall) ----
tiff("heatmap_age28.tiff",
     units = "in",
     width = 12,
     height = 6.5,
     res = 600,
     type = "cairo",
     compression = "lzw")

age_plot <- ggplot(by_study_sf,
                   aes(x = Age...28,
                       y = Study,
                       fill = Effect.Size)) +
  geom_tile(color = "white", linewidth = 0.3) +
  scale_fill_gradient2(
    low = "#C53030",
    mid = "#F2F2F2",
    high = "#2B6CB0",
    midpoint = 0,
    name = "Effect Size"
  ) +
  labs(
    x = "Age Group",
    y = "Study",
    title = "Heatmap of Effect Sizes by Age Group"
  ) +
  geom_text(aes(label = sprintf("%.2f", Effect.Size))) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid = element_blank(),
    axis.text.x = element_text(size = 14, angle = 0, hjust = 0.5),
    axis.title.x = element_text(size = 16, face = "bold"),
    axis.text.y = element_text(size = 14),
    axis.title.y = element_text(size = 16, face = "bold")
  )

print(age_plot)
dev.off()

##########
#heatmap for Modality
#########
by_study_sf <- full %>%
  mutate(
    Study = as.character(Study),
    Modality = as.factor(Modality),
    PublishedYear = as.integer(Published.Year.s)   # make a clean numeric year
  ) %>%
  group_by(Study, Modality) %>%
  summarise(
    PublishedYear = first(PublishedYear),
    Chatbot.Name  = first(Chatbot.Name),
    Effect.Size   = mean(Effect.Size, na.rm = TRUE),
    var           = mean(var, na.rm = TRUE),
    n_rows        = n(),
    .groups = "drop"
  ) %>%
  arrange(PublishedYear, Study) %>%                       # Old → New
  mutate(Study = factor(Study, levels = rev(unique(Study))))  # Oldest on TOP


# Make Age...28 a factor for discrete x-axis ordering (adjust levels as you prefer)
by_study_sf <- by_study_sf %>%
  mutate(
    Modality = as.character(Modality),
    Modality = factor(Modality, levels = sort(unique(Modality)))
  )

by_study_sf <- by_study_sf %>%
  mutate(
    Modality = fct_recode(Modality,
                          "Audio" = "audio", "Text" = "text", "Multimedia" = "multimedia")
  )

# ---- Save as high-resolution TIFF (wider, less tall) ----
tiff("heatmap_modality.tiff",
     units = "in",
     width = 12,
     height = 6.5,
     res = 600,
     type = "cairo",
     compression = "lzw")

Modality_plot <- ggplot(by_study_sf,
                   aes(x = Modality,
                       y = Study,
                       fill = Effect.Size)) +
  geom_tile(color = "white", linewidth = 0.3) +
  scale_fill_gradient2(
    low = "#C53030",
    mid = "#F2F2F2",
    high = "#2B6CB0",
    midpoint = 0,
    name = "Effect Size"
  ) +
  labs(
    x = "Modality",
    y = "Study",
    title = "Heatmap of Effect Sizes by Modality"
  ) +
  geom_text(aes(label = sprintf("%.2f", Effect.Size))) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid = element_blank(),
    axis.text.x = element_text(size = 14, angle = 0, hjust = 0.5),
    axis.title.x = element_text(size = 16, face = "bold"),
    axis.text.y = element_text(size = 14),
    axis.title.y = element_text(size = 16, face = "bold")
  )

print(Modality_plot)
dev.off()

############
#traffic light
############
# Read the first sheet (or specify sheet = "Sheet1")
Rob <- read_excel("RoB.xlsx")

# Peek at your data
str(Rob)
names(Rob)
lapply(Rob, function(x) head(unique(x), 5))

library(dplyr)

standardize_judgement <- function(x) {
  x <- trimws(tolower(as.character(x)))
  dplyr::case_when(
    x %in% c("y") ~ "Low",
    x %in% c("u") ~ "Some concerns",
    x %in% c("n") ~ "High"
  )
}

# Apply to all columns except Study (assumes first column is Study)
Rob_mapped <- Rob %>%
  dplyr::mutate(across(2:(ncol(Rob) - 1), standardize_judgement))

# Optional: check unique values after mapping
# lapply(Rob_mapped, function(col) sort(unique(col)))

# 2) Create traffic light plot using robvis in Generic mode
# If you have an Overall column already as last column, it will be included automatically below.
# If not, you can skip it or create one.

# Select only existing columns; ensure first column is Study (rename if needed)
if (names(Rob_mapped)[1] != "Study") names(Rob_mapped)[1] <- "Study"

# install.packages("robvis")
# library(robvis)
# 
# install.packages("remotes")
# remotes::install_github("mcguinlu/robvis")
# # Assuming rob_generic_df is your mapped JBI data
# # First column must be Study; middle columns are JBI items; last column (if any) is Overall.
# # Update robvis and dependencies
# install.packages("robvis", dependencies = TRUE)
# 
# # Restart R, then:
# library(robvis)
# packageVersion("robvis")          # confirm it updated
# getAnywhere("rob_traffic_light")  # check the function signature
# library(robvis)
# library(ggplot2)

# 2) Define how to categorize Overall Appraisal Score (set thresholds as appropriate)
map_overall <- function(x) {
  # Example thresholds: adjust to your rubric
  dplyr::case_when(
    x >= 8 ~ "Low",
    x >= 5 ~ "Some concerns",
    TRUE    ~ "High"
  )
}

# 3) Build plotting df with a mapped overall, but do NOT change your original last column
last_name <- names(Rob_mapped)[ncol(Rob_mapped)]  # "Overall Appraisal Score"
rob_generic_df <- Rob_mapped %>%
  mutate(Overall_mapped = map_overall(.data[[last_name]])) %>%
  select(-all_of(last_name), everything()) %>%     # drop numeric overall from plotting df
  relocate(Overall_mapped, .after = dplyr::last_col()) %>%
  dplyr::rename(Overall = Overall_mapped)
rob_generic_df <- rob_generic_df %>% dplyr::select(-`Overall Appraisal Score`)
# 4) Plot
p_traffic <- robvis::rob_traffic_light(rob_generic_df, tool = "Generic", psize = 12) +
  ggplot2::scale_fill_manual(values = c(
    "Low" = "#02C3BD",
    "Some concerns" = "#F5BE41",
    "High" = "#E45756"
  ), drop = FALSE, na.translate = FALSE)

n_studies <- nrow(rob_generic_df)
ggplot2::ggsave("traffic_light_JBI_RCT.png", p_traffic,
                width = 10, height = max(6, 0.35*n_studies), dpi = 300)



