# To Successfully run post-hoc tests, first you must make sure you have Final_Task_Frame and Final_KSA_Frame created and
# loaded in the global environment.
# Set WD ----------------------------------------------------------
setwd(paste0("C:/Users/", Sys.info()["user"],"/Documents/"))

# Load data ---------------------------------------------------------------
# In case somebody wants to come back and re-analyze this, can do it without
# having to re-run RJAQ in order to create all relevant variables.
load(file= "~/RJAQ/OUTPUT/Dynamic_Reports/Data/Variables.RData")
 
# Check Assumptions -------------------------------------------------------
if(exists(x = "Final_Task_Frame") & exists(x="Final_KSA_Frame")){print(c("Datasets OK"))} else{stop("Missing Datasets! Have you executed RJAQ?")}

#Create Custom Variable
Statements_NoNA       <- na.omit(Statements)        #The JAQ.xlsx task & KSAO statements without attentiveness and without DA designators.
Statements_Tasks_NoNA <- na.omit(Statements_Tasks)  #The JAQ.xlsx task statements without attentiveness and without DA designators.

# Set Cut-Off Values and Significance Thresholds --------------------------
STAT_pvalue      <- .01  # ALL Tests P-Value.
STAT_L_pvalue    <- .01  # Levene's Test P-Value for T-Test.
STAT_SMD_Cut     <- .25  # Standardized Mean Difference cut-point.
options(scipen=99)       # Where to set cut-off for decimals.

# XLSX Parameter Settings for Output ----------------------------------------
# Create XLSX workbook to save output in a single XLSX file for Gender and a single XLSX file for Race. With a
# cover page that indicates how many of the analyses were significant.
SigAnalyses_Race   <- createWorkbook(type="xlsx")
SigAnalyses_Gender <- createWorkbook(type="xlsx")

#Create custom workbook formatting for columns and rows.
# Styles for the data table column names
TABLE_COLNAMES_STYLE_Gender <- CellStyle(SigAnalyses_Gender) + Fill(foregroundColor = "dodgerblue4")+ Font(SigAnalyses_Gender, isBold=TRUE,name = "Calibri",color = "azure") +
                               Alignment(wrapText=FALSE, horizontal="ALIGN_CENTER") + Border(color="lightgrey", position=c("TOP", "BOTTOM","LEFT","RIGHT"), 
                               pen=c("BORDER_THIN", "BORDER_THICK","BORDER_THIN","BORDER_THIN"))

TABLE_COLNAMES_STYLE_Race   <- CellStyle(SigAnalyses_Race) + Fill(foregroundColor = "dodgerblue4")+ Font(SigAnalyses_Race, isBold=TRUE,name = "Calibri",color = "azure") +
                               Alignment(wrapText=FALSE, horizontal="ALIGN_CENTER") + Border(color="lightgrey", position=c("TOP", "BOTTOM","LEFT","RIGHT"), 
                               pen=c("BORDER_THIN", "BORDER_THICK","BORDER_THIN","BORDER_THIN"))


ROWS_Gender    <- CellStyle(SigAnalyses_Gender) + Font(wb = SigAnalyses_Gender,name="Calibri",heightInPoints = 10) + Alignment(horizontal = "ALIGN_CENTER",wrapText = TRUE,vertical = "VERTICAL_CENTER") + 
                  Border(color="black",position=c("TOP","LEFT","RIGHT","BOTTOM"), pen=c("BORDER_THIN"))

ROWS_Race      <- CellStyle(SigAnalyses_Race) + Font(wb = SigAnalyses_Race,name="Calibri",heightInPoints = 10) + Alignment(horizontal = "ALIGN_CENTER",wrapText = TRUE,vertical = "VERTICAL_CENTER") + 
                  Border(color="black",position=c("TOP","LEFT","RIGHT","BOTTOM"), pen=c("BORDER_THIN"))


dfColIndex_CHI_Gender           <- rep(list(ROWS_Gender), 9) #There will always be 9 columns in all Chi-Square type analyses.
dfColIndex_TT_Gender            <- rep(list(ROWS_Gender), 11) #There will always be 11 columns in t-test related analyses.
names(dfColIndex_CHI_Gender)    <- seq(1, 9, by = 1)
names(dfColIndex_TT_Gender)     <- seq(1, 11, by = 1)
#dfColIndex_CHI_Gender$`2`$alignment$horizontal <- c("ALIGN_LEFT") #FIXTHIS
dfColIndex_CHI_Race             <- rep(list(ROWS_Race), 9) #There will always be 9 columns in all Chi-Square type analyses.
dfColIndex_TT_Race              <- rep(list(ROWS_Race), 11) #There will always be 11 columns in t-test related analyses.
names(dfColIndex_CHI_Race)      <- seq(1, 9, by = 1)
names(dfColIndex_TT_Race)       <- seq(1, 11, by = 1)
  
# Gender.REQU ---------------------------------------------------------
# Task + KSAO
Gender.REQU <- data.frame()
suppressWarnings(
for (i in Statements$Number){
  if(length(table(task[[paste0("REQU_",i)]])) == 2){
    test1 <- chisq.test(table(task$Gender,task[[paste0("REQU_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task$Gender,task[[paste0("REQU_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    Gender.REQU <- rbind(Gender.REQU,row)}
  else if(length(table(task[[paste0("REQU_",i)]])) == 1){
     row <- rep(x = "NA",times = 7)
     Gender.REQU <- rbind(Gender.REQU,row)
    }
})
suppressWarnings(Gender.REQU <- sapply(X = Gender.REQU,2,FUN = as.numeric))
Gender.REQU <- cbind(Statements_NoNA,Gender.REQU)
names(Gender.REQU) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Male Mean","Female Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
#XLSX stuff
Gender.REQU.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "CHI.REQU")
setColumnWidth(Gender.REQU.2,1,7)
setColumnWidth(Gender.REQU.2,2,115)
setColumnWidth(Gender.REQU.2,3:4,20)
setColumnWidth(Gender.REQU.2,5:7,15)
setColumnWidth(Gender.REQU.2,8:9,18)
addDataFrame(x = Gender.REQU, sheet = Gender.REQU.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
createFreezePane(sheet = Gender.REQU.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Gender.ESS ---------------------------------------------------------
# Task
PH_Essentiality_Frame <- cbind(Essential,task$Race,task$Gender)
Gender.ESS <- data.frame()
suppressWarnings(
for (i in Statements_Tasks$Number){
  if(length(table(PH_Essentiality_Frame[[paste0("C_",i)]])) == 2){
    test1 <- chisq.test(table(PH_Essentiality_Frame$`task$Gender`,PH_Essentiality_Frame[[paste0("C_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(PH_Essentiality_Frame$`task$Gender`,PH_Essentiality_Frame[[paste0("C_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    Gender.ESS <- rbind(Gender.ESS,row)} 
  else if(length(table(PH_Essentiality_Frame[[paste0("C_",i)]]) == 1)) {
    row   <- rep(x = "NA",times = 7)
    Gender.ESS <- rbind(Gender.ESS,row)} 
    })
suppressWarnings(Gender.ESS <- sapply(X = Gender.ESS,2,FUN = as.numeric))
Gender.ESS <- cbind(Statements_Tasks,Gender.ESS)
names(Gender.ESS) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Male Mean","Female Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
Gender.ESS.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "CHI.ESS")
setColumnWidth(Gender.ESS.2,1,7)
setColumnWidth(Gender.ESS.2,2,115)
setColumnWidth(Gender.ESS.2,3:4,20)
setColumnWidth(Gender.ESS.2,5:7,15)
setColumnWidth(Gender.ESS.2,8:9,18)
addDataFrame(x = Gender.ESS, sheet = Gender.ESS.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
createFreezePane(sheet = Gender.ESS.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Gender.APP ---------------------------------------------------------
# Task + KSAO
Gender.APP <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(length(table(task[[paste0("NA_",i)]])) == 2){
    test1 <- chisq.test(table(task$Gender,task[[paste0("NA_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task$Gender,task[[paste0("NA_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    Gender.APP <- rbind(Gender.APP,row)}
  else if(length(table(task[[paste0("NA_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    Gender.APP <- rbind(Gender.APP,row)
  }
})
suppressWarnings(Gender.APP <- sapply(X = Gender.APP,2,FUN = as.numeric))
Gender.APP <- cbind(Statements_NoNA,Gender.APP)
names(Gender.APP) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Male Mean","Female Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
Gender.APP.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "CHI.APP")
setColumnWidth(Gender.APP.2,1,7)
setColumnWidth(Gender.APP.2,2,115)
setColumnWidth(Gender.APP.2,3:4,20)
setColumnWidth(Gender.APP.2,5:7,15)
setColumnWidth(Gender.APP.2,8:9,18)
addDataFrame(x = Gender.APP, sheet = Gender.APP.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
createFreezePane(sheet = Gender.APP.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Gender.DIFF ---------------------------------------------------------
# KSAO
#CHI SQUARE IF DICHOT
if(ScaleType_DIFF == "DICHOT"){
Gender.DIFF <- data.frame()
suppressWarnings(
for (i in Statements_KSAOs$Number){
  if(length(table(task[[paste0("DIFF_",i)]])) == 2){
    test1 <- chisq.test(table(task$Gender,task[[paste0("DIFF_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task$Gender,task[[paste0("DIFF_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    Gender.DIFF <- rbind(Gender.DIFF,row)}
  else if(length(table(task[[paste0("DIFF_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    Gender.DIFF <- rbind(Gender.DIFF,row)
  }
})
suppressWarnings(Gender.DIFF <- sapply(X = Gender.DIFF,2,FUN = as.numeric))
Gender.DIFF <- cbind(Statements_KSAOs,Gender.DIFF)
names(Gender.DIFF) <- c("KSAO","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Male Mean","Female Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
Gender.DIFF.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "CHI.DIFF")
setColumnWidth(Gender.DIFF.2,1,7)
setColumnWidth(Gender.DIFF.2,2,115)
setColumnWidth(Gender.DIFF.2,3:4,20)
setColumnWidth(Gender.DIFF.2,5:7,15)
setColumnWidth(Gender.DIFF.2,8:9,18)
addDataFrame(x = Gender.DIFF, sheet = Gender.DIFF.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
createFreezePane(sheet = Gender.DIFF.2,colSplit = 3,rowSplit = 2) # Freeze Panes
}

#TTEST IF LIKERT
if(ScaleType_DIFF == "LIKERT"){
  Gender.DIFF <- data.frame()
  suppressWarnings(
    for (i in Statements_KSAOs$Number){
      if(min(rowMeans(table(task$Gender,task[[paste0("DIFF_",i)]]))) != 0 & min(rowSums(table(task$Gender,task[[paste0("DIFF_",i)]]))) > 1){
        test2 <- leveneTest(task[[paste0("DIFF_",i)]], as.factor(task$Gender), center=mean)
        row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
        x1    <- tapply(task[[paste0("DIFF_",i)]], task$Gender, mean, na.rm=TRUE)#SMD Calculation
        x2    <- tapply(task[[paste0("DIFF_",i)]], task$Gender, sd  , na.rm=TRUE)
        x3    <- table(task$Gender)
        eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
        SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
        SMD <- abs(SMD)
        sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
        
        ifelse(row2[2] <= STAT_L_pvalue,
               {test1 <- t.test(task[[paste0("DIFF_",i)]] ~ task$Gender,var.equal=FALSE)
               sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
               row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
               ,
               {test1 <- t.test(task[[paste0("DIFF_",i)]] ~ task$Gender,var.equal=TRUE)
               sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
               row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)})
        
        row3  <- c(row2,row)
        Gender.DIFF <- rbind(Gender.DIFF,row3)} 
      else if(min(rowMeans(table(task$Gender,task[[paste0("DIFF_",i)]]))) == 0 | min(rowSums(table(task$Gender,task[[paste0("DIFF_",i)]]))) <= 1)
      {Gender.DIFF <- rbind(Gender.DIFF,rep(x = "NA",times=9))}
    })
  suppressWarnings(Gender.DIFF <- sapply(X = Gender.DIFF,2,FUN = as.numeric))
  Gender.DIFF   <- round(x = Gender.DIFF,digits = 3)
  Gender.DIFF   <- cbind(Statements_KSAOs,Gender.DIFF)
  names(Gender.DIFF) <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Male Mean","Female Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
  Gender.DIFF.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "TT.DIFF")
  setColumnWidth(Gender.DIFF.2,1,7)
  setColumnWidth(Gender.DIFF.2,2,115)
  setColumnWidth(Gender.DIFF.2,3:4,11)
  setColumnWidth(Gender.DIFF.2,5:7,13)
  setColumnWidth(Gender.DIFF.2,8:9,15)
  addDataFrame(x = Gender.DIFF, sheet = Gender.DIFF.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
  createFreezePane(sheet = Gender.DIFF.2,colSplit = 3,rowSplit = 2) # Freeze Panes
  }
  
  
  
  
  
  
  






# Gender.IMP ---------------------------------------------------------
# Task + KSAO
Gender.IMP <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(min(rowMeans(table(task$Gender,task[[paste0("IMP_",i)]]))) != 0 & min(rowSums(table(task$Gender,task[[paste0("IMP_",i)]]))) > 1){
  test2 <- leveneTest(task[[paste0("IMP_",i)]], as.factor(task$Gender), center=mean)
  row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
  x1    <- tapply(task[[paste0("IMP_",i)]], task$Gender, mean, na.rm=TRUE)#SMD Calculation
  x2    <- tapply(task[[paste0("IMP_",i)]], task$Gender, sd  , na.rm=TRUE)
  x3    <- table(task$Gender)
  eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
  SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
  SMD <- abs(SMD)
  sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)

  ifelse(row2[2] <= STAT_L_pvalue,
       {test1 <- t.test(task[[paste0("IMP_",i)]] ~ task$Gender,var.equal=FALSE)
       sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
       row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
       ,
       {test1 <- t.test(task[[paste0("IMP_",i)]] ~ task$Gender,var.equal=TRUE)
       sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
       row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)})

row3  <- c(row2,row)
Gender.IMP <- rbind(Gender.IMP,row3)} 
 else if(min(rowMeans(table(task$Gender,task[[paste0("IMP_",i)]]))) == 0 | min(rowSums(table(task$Gender,task[[paste0("IMP_",i)]]))) <= 1)
  {Gender.IMP <- rbind(Gender.IMP,rep(x = "NA",times=9))}
})
suppressWarnings(Gender.IMP <- sapply(X = Gender.IMP,2,FUN = as.numeric))
Gender.IMP   <- round(x = Gender.IMP,digits = 3)
Gender.IMP   <- cbind(Statements_NoNA,Gender.IMP)
names(Gender.IMP) <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Male Mean","Female Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
Gender.IMP.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "TT.IMP")
setColumnWidth(Gender.IMP.2,1,7)
setColumnWidth(Gender.IMP.2,2,115)
setColumnWidth(Gender.IMP.2,3:4,11)
setColumnWidth(Gender.IMP.2,5:7,13)
setColumnWidth(Gender.IMP.2,8:9,15)
addDataFrame(x = Gender.IMP, sheet = Gender.IMP.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
createFreezePane(sheet = Gender.IMP.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Gender.FREQ ---------------------------------------------------------
# Task
Gender.FREQ <- data.frame()
suppressWarnings(
for (i in Statements_Tasks_NoNA$Number){
  if(min(rowMeans(table(task$Gender,task[[paste0("FREQ_",i)]]))) != 0 & min(rowSums(table(task$Gender,task[[paste0("FREQ_",i)]]))) > 1){
    test2 <- leveneTest(task[[paste0("FREQ_",i)]], as.factor(task$Gender), center=mean)
    row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
    x1    <- tapply(task[[paste0("FREQ_",i)]], task$Gender, mean, na.rm=TRUE)#SMD Calculation
    x2    <- tapply(task[[paste0("FREQ_",i)]], task$Gender, sd  , na.rm=TRUE)
    x3    <- table(task$Gender)
    eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
    SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
    SMD <- abs(SMD)
    sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
    
    ifelse(row2[2] <= STAT_L_pvalue,
           {test1 <- t.test(task[[paste0("FREQ_",i)]] ~ task$Gender,var.equal=FALSE)
           sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
           row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
           ,
           {test1 <- t.test(task[[paste0("FREQ_",i)]] ~ task$Gender,var.equal=TRUE)
           sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
           row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)})
    
    row3  <- c(row2,row)
    Gender.FREQ <- rbind(Gender.FREQ,row3)} 
  else if (min(rowMeans(table(task$Gender,task[[paste0("FREQ_",i)]]))) == 0 | min(rowSums(table(task$Gender,task[[paste0("FREQ_",i)]]))) <= 1)
  {Gender.FREQ <- rbind(Gender.FREQ,rep(x = "NA",times=9))}
})

suppressWarnings(Gender.FREQ <- sapply(X = Gender.FREQ,2,FUN = as.numeric))
Gender.FREQ   <- round(x = Gender.FREQ,digits = 3)
Gender.FREQ   <- cbind(Statements_Tasks_NoNA,Gender.FREQ)
names(Gender.FREQ) <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Male Mean","Female Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
Gender.FREQ.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "TT.FREQ")
setColumnWidth(Gender.FREQ.2,1,7)
setColumnWidth(Gender.FREQ.2,2,115)
setColumnWidth(Gender.FREQ.2,3:4,11)
setColumnWidth(Gender.FREQ.2,5:7,13)
setColumnWidth(Gender.FREQ.2,8:9,15)
addDataFrame(x = Gender.FREQ, sheet = Gender.FREQ.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
createFreezePane(sheet = Gender.FREQ.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Gender.RvR -----------------------------------------------------------
# Knowledge TTEST
if(ScaleType_RvR == "LIKERT"){
Gender.RvR <- data.frame()
suppressWarnings(
for (i in Statements_KSAOs$Number[1:length(VarLab_RvR_Task)]){
  if(min(rowMeans(table(task$Gender,task[[paste0("RvR_",i)]]))) != 0 & min(rowSums(table(task$Gender,task[[paste0("RvR_",i)]]))) > 1){
  test2 <- leveneTest(task[[paste0("RvR_",i)]], as.factor(task$Gender), center=mean)
  row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
  x1    <- tapply(task[[paste0("RvR_",i)]], task$Gender, mean, na.rm=TRUE)#SMD Calculation
  x2    <- tapply(task[[paste0("RvR_",i)]], task$Gender, sd  , na.rm=TRUE)
  x3    <- table(task$Gender)
  eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
  SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
  SMD <- abs(SMD)
  sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
  
  ifelse(row2[2] <= STAT_L_pvalue,
         {test1 <- t.test(task[[paste0("RvR_",i)]] ~ task$Gender,var.equal=FALSE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
         ,
         {test1 <- t.test(task[[paste0("RvR_",i)]] ~ task$Gender,var.equal=TRUE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
  )
  row3  <- c(row2,row)
  Gender.RvR <- rbind(Gender.RvR,row3)}  
  else if (min(rowMeans(table(task$Gender,task[[paste0("RvR_",i)]]))) == 0 | min(rowSums(table(task$Gender,task[[paste0("RvR_",i)]]))) <= 1)
  {Gender.RvR <- rbind(Gender.RvR,rep(x="NA",times=9))}
})
suppressWarnings(Gender.RvR <- sapply(X = Gender.RvR,2,FUN = as.numeric))
Gender.RvR   <- round(x = Gender.RvR,digits = 3)
Gender.RvR   <- cbind(Statements_KSAOs[1:length(VarLab_RvR_Task),],Gender.RvR)
names(Gender.RvR) <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Male Mean","Female Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
Gender.RvR.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "TT.RvR")
setColumnWidth(Gender.RvR.2,1,7)
setColumnWidth(Gender.RvR.2,2,115)
setColumnWidth(Gender.RvR.2,3:4,11)
setColumnWidth(Gender.RvR.2,5:7,13)
setColumnWidth(Gender.RvR.2,8:9,15)
addDataFrame(x = Gender.RvR, sheet = Gender.RvR.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
createFreezePane(sheet = Gender.RvR.2,colSplit = 3,rowSplit = 2) # Freeze Panes
saveWorkbook(wb = SigAnalyses_Gender,file = paste0(getwd(),"/","RJAQ/OUTPUT/Significance_Analyses/SignificanceAnalyses_Gender.xlsx"))
}

#DICHOT SCALE
if(ScaleType_RvR == "DICHOT"){
  Gender.RvR <- data.frame()
  suppressWarnings(
    for (i in Statements_KSAOs$Number[1:length(VarLab_RvR_Task)]){
      if(length(table(task[[paste0("RvR_",i)]])) == 2){
        test1 <- chisq.test(table(task$Gender,task[[paste0("RvR_",i)]]),correct = FALSE)
        test2 <- fisher.test(table(task$Gender,task[[paste0("RvR_",i)]]))
        sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
        sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
        a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
        a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
        row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
        Gender.RvR <- rbind(Gender.RvR,row)}
      else if(length(table(task[[paste0("RvR_",i)]])) != 2){
        row <- rep(x = "NA",times = 7)
        Gender.RvR <- rbind(Gender.RvR,row)
      }
    })
  suppressWarnings(Gender.RvR <- sapply(X = Gender.RvR,2,FUN = as.numeric))
  Gender.RvR <- cbind(Statements_KSAOs[1:length(VarLab_RvR_Task),],Gender.RvR)
  names(Gender.RvR) <- c("KSAO","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Male Mean","Female Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
  Gender.RvR.2     <- createSheet(wb = SigAnalyses_Gender,sheetName = "CHI.RvR")
  setColumnWidth(Gender.RvR.2,1,7)
  setColumnWidth(Gender.RvR.2,2,115)
  setColumnWidth(Gender.RvR.2,3:4,20)
  setColumnWidth(Gender.RvR.2,5:7,15)
  setColumnWidth(Gender.RvR.2,8:9,18)
  addDataFrame(x = Gender.RvR, sheet = Gender.RvR.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
  createFreezePane(sheet = Gender.RvR.2,colSplit = 3,rowSplit = 2) # Freeze Panes
  saveWorkbook(wb = SigAnalyses_Gender,file = paste0(getwd(),"/","RJAQ/OUTPUT/Significance_Analyses/SignificanceAnalyses_Gender.xlsx"))
  
}

# Race BW REQU --------------------------------------------------------
task_Race_BW <- task[which(task$Race == 1 | task$Race == 6),]
RaceBW.REQU <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(length(table(task_Race_BW[[paste0("REQU_",i)]])) == 2){
    test1 <- chisq.test(table(task_Race_BW$Race,task_Race_BW[[paste0("REQU_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task_Race_BW$Race,task_Race_BW[[paste0("REQU_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceBW.REQU <- rbind(RaceBW.REQU,row)}
  else if(length(table(task_Race_BW[[paste0("REQU_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceBW.REQU <- rbind(RaceBW.REQU,row)
  }
})
suppressWarnings(RaceBW.REQU <- sapply(X = RaceBW.REQU,2,FUN = as.numeric))
RaceBW.REQU <- cbind(Statements_NoNA,RaceBW.REQU)
names(RaceBW.REQU) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Black Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceBW.REQU.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "BW.CHI.REQU")
setColumnWidth(RaceBW.REQU.2,1,7)
setColumnWidth(RaceBW.REQU.2,2,115)
setColumnWidth(RaceBW.REQU.2,3:4,20)
setColumnWidth(RaceBW.REQU.2,5:7,15)
setColumnWidth(RaceBW.REQU.2,8:9,18)
addDataFrame(x = RaceBW.REQU, sheet = RaceBW.REQU.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceBW.REQU.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race BW ESS --------------------------------------------------------
PH_Essentiality_Frame <- cbind(Essential,task$Race,task$Gender)
PH_Essentiality_Frame <- PH_Essentiality_Frame[which(PH_Essentiality_Frame$`task$Race` == 1 | PH_Essentiality_Frame$`task$Race` == 6),]
RaceBW.ESS <- data.frame()
suppressWarnings(
for (i in Statements_Tasks$Number){
  if(length(table(PH_Essentiality_Frame[[paste0("C_",i)]])) == 2){
    test1 <- chisq.test(table(PH_Essentiality_Frame$`task$Race`,PH_Essentiality_Frame[[paste0("C_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(PH_Essentiality_Frame$`task$Race`,PH_Essentiality_Frame[[paste0("C_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceBW.ESS <- rbind(RaceBW.ESS,row)}
  else if(length(table(PH_Essentiality_Frame[[paste0("C_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceBW.ESS <- rbind(RaceBW.ESS,row)
  }
})
suppressWarnings(RaceBW.ESS <- sapply(X = RaceBW.ESS,2,FUN = as.numeric))
RaceBW.ESS <- cbind(Statements_Tasks,RaceBW.ESS)
names(RaceBW.ESS) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Black Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceBW.ESS.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "BW.CHI.ESS")
setColumnWidth(RaceBW.ESS.2,1,7)
setColumnWidth(RaceBW.ESS.2,2,115)
setColumnWidth(RaceBW.ESS.2,3:4,20)
setColumnWidth(RaceBW.ESS.2,5:7,15)
setColumnWidth(RaceBW.ESS.2,8:9,18)
addDataFrame(x = RaceBW.ESS, sheet = RaceBW.ESS.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceBW.ESS.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race BW APP --------------------------------------------------------
task_Race_BW <- task[which(task$Race == 1 | task$Race == 6),]
RaceBW.APP <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(length(table(task_Race_BW[[paste0("NA_",i)]])) == 2){
    test1 <- chisq.test(table(task_Race_BW$Race,task_Race_BW[[paste0("NA_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task_Race_BW$Race,task_Race_BW[[paste0("NA_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceBW.APP <- rbind(RaceBW.APP,row)}
  else if(length(table(task_Race_BW[[paste0("NA_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceBW.APP <- rbind(RaceBW.APP,row)
  }
})
suppressWarnings(RaceBW.APP <- sapply(X = RaceBW.APP,2,FUN = as.numeric))
RaceBW.APP <- cbind(Statements_NoNA,RaceBW.APP)
names(RaceBW.APP) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Black Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceBW.APP.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "BW.CHI.APP")
setColumnWidth(RaceBW.APP.2,1,7)
setColumnWidth(RaceBW.APP.2,2,115)
setColumnWidth(RaceBW.APP.2,3:4,20)
setColumnWidth(RaceBW.APP.2,5:7,15)
setColumnWidth(RaceBW.APP.2,8:9,18)
addDataFrame(x = RaceBW.APP, sheet = RaceBW.APP.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceBW.APP.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race BW DIFF --------------------------------------------------------
#CHISQUARE TEST
task_Race_BW <- task[which(task$Race == 1 | task$Race == 6),]
RaceBW.DIFF <- data.frame()
if(ScaleType_DIFF == "DICHOT") {
suppressWarnings(
for (i in Statements_KSAOs$Number){
  if(length(table(task_Race_BW[[paste0("DIFF_",i)]])) == 2){
    test1 <- chisq.test(table(task_Race_BW$Race,task_Race_BW[[paste0("DIFF_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task_Race_BW$Race,task_Race_BW[[paste0("DIFF_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceBW.DIFF <- rbind(RaceBW.DIFF,row)}
  else if(length(table(task_Race_BW[[paste0("DIFF_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceBW.DIFF <- rbind(RaceBW.DIFF,row)
  }
})
suppressWarnings(RaceBW.DIFF <- sapply(X = RaceBW.DIFF,2,FUN = as.numeric))
RaceBW.DIFF <- cbind(Statements_KSAOs,RaceBW.DIFF)
names(RaceBW.DIFF) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Black Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceBW.DIFF.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "BW.CHI.DIFF")
setColumnWidth(RaceBW.DIFF.2,1,7)
setColumnWidth(RaceBW.DIFF.2,2,115)
setColumnWidth(RaceBW.DIFF.2,3:4,20)
setColumnWidth(RaceBW.DIFF.2,5:7,15)
setColumnWidth(RaceBW.DIFF.2,8:9,18)
addDataFrame(x = RaceBW.DIFF, sheet = RaceBW.DIFF.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
}

#TTEST IF LIKERT
if(ScaleType_DIFF == "LIKERT"){
  RaceBW.DIFF <- data.frame()
  suppressWarnings(
    for (i in Statements_KSAOs$Number){
      if(min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("DIFF_",i)]]))) != 0 & min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("DIFF_",i)]]))) > 1){
        test2 <- leveneTest(task_Race_BW[[paste0("DIFF_",i)]], as.factor(task_Race_BW$Race), center=mean)
        row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
        x1    <- tapply(task_Race_BW[[paste0("DIFF_",i)]], task_Race_BW$Race, mean, na.rm=TRUE)#SMD Calculation
        x2    <- tapply(task_Race_BW[[paste0("DIFF_",i)]], task_Race_BW$Race, sd  , na.rm=TRUE)
        x3    <- table(task_Race_BW$Race)
        eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
        SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
        SMD <- abs(SMD)
        sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
        
        ifelse(row2[2] <= STAT_L_pvalue,
               {test1 <- t.test(task_Race_BW[[paste0("DIFF_",i)]] ~ task_Race_BW$Race,var.equal=FALSE)
               sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
               row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
               ,
               {test1 <- t.test(task_Race_BW[[paste0("DIFF_",i)]] ~ task_Race_BW$Race,var.equal=TRUE)
               sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
               row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)})
        
        row3  <- c(row2,row)
        RaceBW.DIFF <- rbind(RaceBW.DIFF,row3)} 
      else if(min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("DIFF_",i)]]))) == 0 | min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("DIFF_",i)]]))) <= 1)
      {RaceBW.DIFF <- rbind(RaceBW.DIFF,rep(x = "NA",times=9))}
    })
  suppressWarnings(RaceBW.DIFF <- sapply(X = RaceBW.DIFF,2,FUN = as.numeric))
  RaceBW.DIFF   <- round(x = RaceBW.DIFF,digits = 3)
  RaceBW.DIFF   <- cbind(Statements_KSAOs,RaceBW.DIFF)
  names(RaceBW.DIFF) <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Black Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
  RaceBW.DIFF.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "TT.DIFF")
  setColumnWidth(RaceBW.DIFF.2,1,7)
  setColumnWidth(RaceBW.DIFF.2,2,115)
  setColumnWidth(RaceBW.DIFF.2,3:4,11)
  setColumnWidth(RaceBW.DIFF.2,5:7,13)
  setColumnWidth(RaceBW.DIFF.2,8:9,15)
  addDataFrame(x = RaceBW.DIFF, sheet = RaceBW.DIFF.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
  createFreezePane(sheet = RaceBW.DIFF.2,colSplit = 3,rowSplit = 2) # Freeze Panes
}

# Race BW IMP ---------------------------------------------------------
# Task + KSAO
task_Race_BW <- task[which(task$Race == 1 | task$Race == 6),]
RaceBW.IMP <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("IMP_",i)]]))) != 0 & min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("IMP_",i)]]))) > 1){
  test2 <- leveneTest(task_Race_BW[[paste0("IMP_",i)]], as.factor(task_Race_BW$Race), center=mean)
  row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
  x1    <- tapply(task_Race_BW[[paste0("IMP_",i)]], task_Race_BW$Race, mean, na.rm=TRUE)#SMD Calculation
  x2    <- tapply(task_Race_BW[[paste0("IMP_",i)]], task_Race_BW$Race, sd  , na.rm=TRUE)
  x3    <- table(task_Race_BW$Race)
  eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
  SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
  SMD <- abs(SMD)
  sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
  
  ifelse(row2[2] <= STAT_L_pvalue,
         {test1 <- t.test(task_Race_BW[[paste0("IMP_",i)]] ~ task_Race_BW$Race,var.equal=FALSE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
         ,
         {test1 <- t.test(task_Race_BW[[paste0("IMP_",i)]] ~ task_Race_BW$Race,var.equal=TRUE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
  )
  row3  <- c(row2,row)
  RaceBW.IMP <- rbind(RaceBW.IMP,row3)}  
  else if (min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("IMP_",i)]]))) == 0 | min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("IMP_",i)]]))) <= 1)
  {RaceBW.IMP <- rbind(RaceBW.IMP,rep(x="NA",times=9))}
})
suppressWarnings(RaceBW.IMP <- sapply(X = RaceBW.IMP,2,FUN = as.numeric))
RaceBW.IMP                  <- round(x = RaceBW.IMP,digits = 3)
RaceBW.IMP                  <- cbind(Statements_NoNA,RaceBW.IMP)
names(RaceBW.IMP)          <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Black Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
RaceBW.IMP.2                    <- createSheet(wb = SigAnalyses_Race,sheetName = "BW.TT.IMP")
setColumnWidth(RaceBW.IMP.2,1,7)
setColumnWidth(RaceBW.IMP.2,2,115)
setColumnWidth(RaceBW.IMP.2,3:4,11)
setColumnWidth(RaceBW.IMP.2,5:7,13)
setColumnWidth(RaceBW.IMP.2,8:9,15)
addDataFrame(x = RaceBW.IMP, sheet = RaceBW.IMP.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceBW.IMP.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race BW FREQ ---------------------------------------------------------
# Task + KSAO
task_Race_BW <- task[which(task$Race == 1 | task$Race == 6),]
RaceBW.FREQ <- data.frame()
suppressWarnings(
for (i in Statements_Tasks_NoNA$Number){
  if(min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("FREQ_",i)]]))) != 0 & min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("FREQ_",i)]]))) > 1){
  test2 <- leveneTest(task_Race_BW[[paste0("FREQ_",i)]], as.factor(task_Race_BW$Race), center=mean)
  row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
  x1    <- tapply(task_Race_BW[[paste0("FREQ_",i)]], task_Race_BW$Race, mean, na.rm=TRUE)#SMD Calculation
  x2    <- tapply(task_Race_BW[[paste0("FREQ_",i)]], task_Race_BW$Race, sd  , na.rm=TRUE)
  x3    <- table(task_Race_BW$Race)
  eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
  SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
  SMD <- abs(SMD)
  sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
  
  ifelse(row2[2] <= STAT_L_pvalue,
         {test1 <- t.test(task_Race_BW[[paste0("FREQ_",i)]] ~ task_Race_BW$Race,var.equal=FALSE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
         ,
         {test1 <- t.test(task_Race_BW[[paste0("FREQ_",i)]] ~ task_Race_BW$Race,var.equal=TRUE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
  )
  row3  <- c(row2,row)
  RaceBW.FREQ <- rbind(RaceBW.FREQ,row3)}
  else if (min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("FREQ_",i)]]))) == 0 | min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("FREQ_",i)]]))) <= 1)
  {RaceBW.FREQ <- rbind(RaceBW.FREQ,rep(x="NA",times=9))}
})
suppressWarnings(RaceBW.FREQ <- sapply(X = RaceBW.FREQ,2,FUN = as.numeric))
RaceBW.FREQ                  <- round(x = RaceBW.FREQ,digits = 3)
RaceBW.FREQ                  <- cbind(Statements_Tasks_NoNA,RaceBW.FREQ)
names(RaceBW.FREQ)          <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Black Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
RaceBW.FREQ.2                    <- createSheet(wb = SigAnalyses_Race,sheetName = "BW.TT.FREQ")
setColumnWidth(RaceBW.FREQ.2,1,7)
setColumnWidth(RaceBW.FREQ.2,2,115)
setColumnWidth(RaceBW.FREQ.2,3:4,11)
setColumnWidth(RaceBW.FREQ.2,5:7,13)
setColumnWidth(RaceBW.FREQ.2,8:9,15)
addDataFrame(x = RaceBW.FREQ, sheet = RaceBW.FREQ.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceBW.FREQ.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race BW RvR ---------------------------------------------------------
# Knowledge
task_Race_BW <- task[which(task$Race == 1 | task$Race == 6),]
RaceBW.RvR <- data.frame()
if(ScaleType_RvR == "LIKERT"){
suppressWarnings(
for (i in Statements_KSAOs$Number[1:length(VarLab_RvR_Task)]){
  if(min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("RvR_",i)]]))) != 0 & min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("RvR_",i)]]))) > 1){
  test2 <- leveneTest(task_Race_BW[[paste0("RvR_",i)]], as.factor(task_Race_BW$Race), center=mean)
  row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
  x1    <- tapply(task_Race_BW[[paste0("RvR_",i)]], task_Race_BW$Race, mean, na.rm=TRUE)#SMD Calculation
  x2    <- tapply(task_Race_BW[[paste0("RvR_",i)]], task_Race_BW$Race, sd  , na.rm=TRUE)
  x3    <- table(task_Race_BW$Race)
  eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
  SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
  SMD <- abs(SMD)
  sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
  
  ifelse(row2[2] <= STAT_L_pvalue,
         {test1 <- t.test(task_Race_BW[[paste0("RvR_",i)]] ~ task_Race_BW$Race,var.equal=FALSE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
         ,
         {test1 <- t.test(task_Race_BW[[paste0("RvR_",i)]] ~ task_Race_BW$Race,var.equal=TRUE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
  )
  row3  <- c(row2,row)
  RaceBW.RvR <- rbind(RaceBW.RvR,row3)}
  else if (min(rowMeans(table(task_Race_BW$Race,task_Race_BW[[paste0("RvR_",i)]]))) == 0 | min(rowSums(table(task_Race_BW$Race,task_Race_BW[[paste0("RvR_",i)]]))) <= 1)
  {RaceBW.RvR <- rbind(RaceBW.RvR,rep(x="NA",times=9))}
})
suppressWarnings(RaceBW.RvR   <- sapply(X = RaceBW.RvR,2,FUN = as.numeric))
RaceBW.RvR                    <- round(x = RaceBW.RvR,digits = 3)
RaceBW.RvR                    <- cbind(Statements_KSAOs[1:length(VarLab_RvR_Task),],RaceBW.RvR)
names(RaceBW.RvR) <- c("KNOW","Description","Levene's F","Levene P","T Statistic","T P-Value","Black Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
RaceBW.RvR.2                      <- createSheet(wb = SigAnalyses_Race,sheetName = "BW.TT.RvR")
setColumnWidth(RaceBW.RvR.2,1,7)
setColumnWidth(RaceBW.RvR.2,2,115)
setColumnWidth(RaceBW.RvR.2,3:4,11)
setColumnWidth(RaceBW.RvR.2,5:7,13)
setColumnWidth(RaceBW.RvR.2,8:9,15)
addDataFrame(x = RaceBW.RvR, sheet = RaceBW.RvR.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceBW.RvR.2,colSplit = 3,rowSplit = 2) # Freeze Panes
}

#DICHOT SCALE
if(ScaleType_RvR == "DICHOT"){
  RaceBW.RvR <- data.frame()
  suppressWarnings(
    for (i in Statements_KSAOs$Number[1:length(VarLab_RvR_Task)]){
      if(length(table(task_Race_BW[[paste0("RvR_",i)]])) == 2){
        test1 <- chisq.test(table(task_Race_BW$Race,task_Race_BW[[paste0("RvR_",i)]]),correct = FALSE)
        test2 <- fisher.test(table(task_Race_BW$Race,task_Race_BW[[paste0("RvR_",i)]]))
        sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
        sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
        a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
        a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
        row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
        RaceBW.RvR <- rbind(RaceBW.RvR,row)}
      else if(length(table(task_Race_BW[[paste0("RvR_",i)]])) != 2){
        row <- rep(x = "NA",times = 7)
        RaceBW.RvR <- rbind(RaceBW.RvR,row)
      }
    })
  suppressWarnings(RaceBW.RvR <- sapply(X = RaceBW.RvR,2,FUN = as.numeric))
  RaceBW.RvR <- cbind(Statements_KSAOs[1:length(VarLab_RvR_Task),],RaceBW.RvR)
  names(RaceBW.RvR) <- c("KSAO","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Black Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
  RaceBW.RvR.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "CHI.RvR")
  setColumnWidth(RaceBW.RvR.2,1,7)
  setColumnWidth(RaceBW.RvR.2,2,115)
  setColumnWidth(RaceBW.RvR.2,3:4,20)
  setColumnWidth(RaceBW.RvR.2,5:7,15)
  setColumnWidth(RaceBW.RvR.2,8:9,18)
  addDataFrame(x = RaceBW.RvR, sheet = RaceBW.RvR.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
  createFreezePane(sheet = RaceBW.RvR.2,colSplit = 3,rowSplit = 2) # Freeze Panes
}

# Race HW REQU --------------------------------------------------------
task_Race_HW <- task[which(task$Race == 3 | task$Race == 6),]
RaceHW.REQU <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(length(table(task_Race_HW[[paste0("REQU_",i)]])) == 2){
    test1 <- chisq.test(table(task_Race_HW$Race,task_Race_HW[[paste0("REQU_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task_Race_HW$Race,task_Race_HW[[paste0("REQU_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceHW.REQU <- rbind(RaceHW.REQU,row)}
  else if(length(table(task_Race_HW[[paste0("REQU_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceHW.REQU <- rbind(RaceHW.REQU,row)
  }
})
suppressWarnings(RaceHW.REQU <- sapply(X = RaceHW.REQU,2,FUN = as.numeric))
RaceHW.REQU <- cbind(Statements_NoNA,RaceHW.REQU)
names(RaceHW.REQU) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Hispanic Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceHW.REQU.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "HW.CHI.REQU")
setColumnWidth(RaceHW.REQU.2,1,7)
setColumnWidth(RaceHW.REQU.2,2,115)
setColumnWidth(RaceHW.REQU.2,3:4,20)
setColumnWidth(RaceHW.REQU.2,5:7,15)
setColumnWidth(RaceHW.REQU.2,8:9,18)
addDataFrame(x = RaceHW.REQU, sheet = RaceHW.REQU.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceHW.REQU.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race HW ESS --------------------------------------------------------
PH_Essentiality_Frame <- cbind(Essential,task$Race,task$Gender)
PH_Essentiality_Frame <- PH_Essentiality_Frame[which(PH_Essentiality_Frame$`task$Race` == 3 | PH_Essentiality_Frame$`task$Race` == 6),]
RaceHW.ESS <- data.frame()
#xx <- NULL
suppressWarnings(
for (i in Statements_Tasks_NoNA$Number){
  if(length(table(PH_Essentiality_Frame[[paste0("C_",i)]])) == 2){
    test1 <- chisq.test(table(PH_Essentiality_Frame$`task$Race`,PH_Essentiality_Frame[[paste0("C_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(PH_Essentiality_Frame$`task$Race`,PH_Essentiality_Frame[[paste0("C_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceHW.ESS <- rbind(RaceHW.ESS,row)}
  else if(length(table(PH_Essentiality_Frame[[paste0("C_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceHW.ESS <- rbind(RaceHW.ESS,row)
  }
})
suppressWarnings(RaceHW.ESS <- sapply(X = RaceHW.ESS,2,FUN = as.numeric))
RaceHW.ESS <- cbind(Statements_Tasks,RaceHW.ESS)
names(RaceHW.ESS) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Hispanic Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceHW.ESS.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "HW.CHI.ESS")
setColumnWidth(RaceHW.ESS.2,1,7)
setColumnWidth(RaceHW.ESS.2,2,115)
setColumnWidth(RaceHW.ESS.2,3:4,20)
setColumnWidth(RaceHW.ESS.2,5:7,15)
setColumnWidth(RaceHW.ESS.2,8:9,18)
addDataFrame(x = RaceHW.ESS, sheet = RaceHW.ESS.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceHW.ESS.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race HW APP --------------------------------------------------------
task_Race_HW <- task[which(task$Race == 3 | task$Race == 6),]
RaceHW.APP <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(length(table(task_Race_HW[[paste0("NA_",i)]])) == 2){
    test1 <- chisq.test(table(task_Race_HW$Race,task_Race_HW[[paste0("NA_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task_Race_HW$Race,task_Race_HW[[paste0("NA_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceHW.APP <- rbind(RaceHW.APP,row)}
  else if(length(table(task_Race_HW[[paste0("NA_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceHW.APP <- rbind(RaceHW.APP,row)
  }
})
suppressWarnings(RaceHW.APP <- sapply(X = RaceHW.APP,2,FUN = as.numeric))
RaceHW.APP <- cbind(Statements_NoNA,RaceHW.APP)
names(RaceHW.APP) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Hispanic Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceHW.APP.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "HW.CHI.APP")
setColumnWidth(RaceHW.APP.2,1,7)
setColumnWidth(RaceHW.APP.2,2,115)
setColumnWidth(RaceHW.APP.2,3:4,20)
setColumnWidth(RaceHW.APP.2,5:7,15)
setColumnWidth(RaceHW.APP.2,8:9,18)
addDataFrame(x = RaceHW.APP, sheet = RaceHW.APP.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceHW.APP.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race HW DIFF --------------------------------------------------------
task_Race_HW <- task[which(task$Race == 3 | task$Race == 6),]
RaceHW.DIFF <- data.frame()
if(ScaleType_DIFF == "DICHOT"){
suppressWarnings(
for (i in Statements_KSAOs$Number){
  if(length(table(task_Race_HW[[paste0("DIFF_",i)]])) == 2){
    test1 <- chisq.test(table(task_Race_HW$Race,task_Race_HW[[paste0("DIFF_",i)]]),correct = FALSE)
    test2 <- fisher.test(table(task_Race_HW$Race,task_Race_HW[[paste0("DIFF_",i)]]))
    sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
    sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
    a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
    a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
    row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
    RaceHW.DIFF <- rbind(RaceHW.DIFF,row)}
  else if(length(table(task_Race_HW[[paste0("DIFF_",i)]])) != 2){
    row <- rep(x = "NA",times = 7)
    RaceHW.DIFF <- rbind(RaceHW.DIFF,row)
  }
})
suppressWarnings(RaceHW.DIFF <- sapply(X = RaceHW.DIFF,2,FUN = as.numeric))
RaceHW.DIFF <- cbind(Statements_KSAOs,RaceHW.DIFF)
names(RaceHW.DIFF) <- c("Task","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Hispanic Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
RaceHW.DIFF.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "HW.CHI.DIFF")
setColumnWidth(RaceHW.DIFF.2,1,7)
setColumnWidth(RaceHW.DIFF.2,2,115)
setColumnWidth(RaceHW.DIFF.2,3:4,20)
setColumnWidth(RaceHW.DIFF.2,5:7,15)
setColumnWidth(RaceHW.DIFF.2,8:9,18)
addDataFrame(x = RaceHW.DIFF, sheet = RaceHW.DIFF.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
}

#TTEST IF LIKERT
if(ScaleType_DIFF == "LIKERT"){
  RaceHW.DIFF <- data.frame()
  suppressWarnings(
    for (i in Statements_KSAOs$Number){
      if(min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("DIFF_",i)]]))) != 0 & min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("DIFF_",i)]]))) > 1){
        test2 <- leveneTest(task_Race_HW[[paste0("DIFF_",i)]], as.factor(task_Race_HW$Race), center=mean)
        row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
        x1    <- tapply(task_Race_HW[[paste0("DIFF_",i)]], task_Race_HW$Race, mean, na.rm=TRUE)#SMD Calculation
        x2    <- tapply(task_Race_HW[[paste0("DIFF_",i)]], task_Race_HW$Race, sd  , na.rm=TRUE)
        x3    <- table(task_Race_HW$Race)
        eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
        SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
        SMD <- abs(SMD)
        sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
        
        ifelse(row2[2] <= STAT_L_pvalue,
               {test1 <- t.test(task_Race_HW[[paste0("DIFF_",i)]] ~ task_Race_HW$Race,var.equal=FALSE)
               sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
               row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
               ,
               {test1 <- t.test(task_Race_HW[[paste0("DIFF_",i)]] ~ task_Race_HW$Race,var.equal=TRUE)
               sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
               row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)})
        
        row3  <- c(row2,row)
        RaceHW.DIFF <- rbind(RaceHW.DIFF,row3)} 
      else if(min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("DIFF_",i)]]))) == 0 | min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("DIFF_",i)]]))) <= 1)
      {RaceHW.DIFF <- rbind(RaceHW.DIFF,rep(x = "NA",times=9))}
    })
  suppressWarnings(RaceHW.DIFF <- sapply(X = RaceHW.DIFF,2,FUN = as.numeric))
  RaceHW.DIFF   <- round(x = RaceHW.DIFF,digits = 3)
  RaceHW.DIFF   <- cbind(Statements_KSAOs,RaceHW.DIFF)
  names(RaceHW.DIFF) <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Hispanic Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
  RaceHW.DIFF.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "TT.DIFF")
  setColumnWidth(RaceHW.DIFF.2,1,7)
  setColumnWidth(RaceHW.DIFF.2,2,115)
  setColumnWidth(RaceHW.DIFF.2,3:4,11)
  setColumnWidth(RaceHW.DIFF.2,5:7,13)
  setColumnWidth(RaceHW.DIFF.2,8:9,15)
  addDataFrame(x = RaceHW.DIFF, sheet = RaceHW.DIFF.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
  createFreezePane(sheet = RaceHW.DIFF.2,colSplit = 3,rowSplit = 2) # Freeze Panes
}




# Race HW IMP ---------------------------------------------------------
# Task + KSAO
task_Race_HW <- task[which(task$Race == 3 | task$Race == 6),]
RaceHW.IMP <- data.frame()
suppressWarnings(
for (i in Statements_NoNA$Number){
  if(min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("IMP_",i)]]))) != 0 & min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("IMP_",i)]]))) > 1){
  test2 <- leveneTest(task_Race_HW[[paste0("IMP_",i)]], as.factor(task_Race_HW$Race), center=mean)
  row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
  x1    <- tapply(task_Race_HW[[paste0("IMP_",i)]], task_Race_HW$Race, mean, na.rm=TRUE)#SMD Calculation
  x2    <- tapply(task_Race_HW[[paste0("IMP_",i)]], task_Race_HW$Race, sd  , na.rm=TRUE)
  x3    <- table(task_Race_HW$Race)
  eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
  SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
  SMD <- abs(SMD)
  sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
  
  ifelse(row2[2] <= STAT_L_pvalue,
         {test1 <- t.test(task_Race_HW[[paste0("IMP_",i)]] ~ task_Race_HW$Race,var.equal=FALSE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
         ,
         {test1 <- t.test(task_Race_HW[[paste0("IMP_",i)]] ~ task_Race_HW$Race,var.equal=TRUE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
  )
  row3  <- c(row2,row)
  RaceHW.IMP <- rbind(RaceHW.IMP,row3)}
  else if (min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("IMP_",i)]]))) == 0 | min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("IMP_",i)]]))) <= 1)
  {RaceHW.IMP <- rbind(RaceHW.IMP,rep(x="NA",times=9))}
})
suppressWarnings(RaceHW.IMP <- sapply(X = RaceHW.IMP,2,FUN = as.numeric))
RaceHW.IMP                  <- round(x = RaceHW.IMP,digits = 3)
RaceHW.IMP                  <- cbind(Statements_NoNA,RaceHW.IMP)
names(RaceHW.IMP)          <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Hispanic Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
RaceHW.IMP.2                    <- createSheet(wb = SigAnalyses_Race,sheetName = "HW.TT.IMP")
setColumnWidth(RaceHW.IMP.2,1,7)
setColumnWidth(RaceHW.IMP.2,2,115)
setColumnWidth(RaceHW.IMP.2,3:4,11)
setColumnWidth(RaceHW.IMP.2,5:7,13)
setColumnWidth(RaceHW.IMP.2,8:9,15)
addDataFrame(x = RaceHW.IMP, sheet = RaceHW.IMP.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceHW.IMP.2,colSplit = 3,rowSplit = 2) # Freeze Panes

# Race HW FREQ ---------------------------------------------------------
# Task
task_Race_HW <- task[which(task$Race == 3 | task$Race == 6),]
RaceHW.FREQ <- data.frame()
suppressWarnings(
for (i in Statements_Tasks_NoNA$Number){
  if(min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("FREQ_",i)]]))) != 0 & min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("FREQ_",i)]]))) > 1){
  test2 <- leveneTest(task_Race_HW[[paste0("FREQ_",i)]], as.factor(task_Race_HW$Race), center=mean)
  row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
  x1    <- tapply(task_Race_HW[[paste0("FREQ_",i)]], task_Race_HW$Race, mean, na.rm=TRUE)#SMD Calculation
  x2    <- tapply(task_Race_HW[[paste0("FREQ_",i)]], task_Race_HW$Race, sd  , na.rm=TRUE)
  x3    <- table(task_Race_HW$Race)
  eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
  SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
  SMD <- abs(SMD)
  sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
  
  ifelse(row2[2] <= STAT_L_pvalue,
         {test1 <- t.test(task_Race_HW[[paste0("FREQ_",i)]] ~ task_Race_HW$Race,var.equal=FALSE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
         ,
         {test1 <- t.test(task_Race_HW[[paste0("FREQ_",i)]] ~ task_Race_HW$Race,var.equal=TRUE)
         sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
         row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
  )
  row3  <- c(row2,row)
  RaceHW.FREQ <- rbind(RaceHW.FREQ,row3)}
  else if (min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("FREQ_",i)]]))) == 0 | min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("FREQ_",i)]]))) <= 1)
  {RaceHW.FREQ <- rbind(RaceHW.FREQ,rep(x="NA",times=9))}
})
suppressWarnings(RaceHW.FREQ <- sapply(X = RaceHW.FREQ,2,FUN = as.numeric))
RaceHW.FREQ                  <- round(x = RaceHW.FREQ,digits = 3)
RaceHW.FREQ                  <- cbind(Statements_Tasks_NoNA,RaceHW.FREQ)
names(RaceHW.FREQ)          <- c("Task","Description","Levene's F","Levene P","T Statistic","T P-Value","Hispanic Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
RaceHW.FREQ.2                    <- createSheet(wb = SigAnalyses_Race,sheetName = "HW.TT.FREQ")
setColumnWidth(RaceHW.FREQ.2,1,7)
setColumnWidth(RaceHW.FREQ.2,2,115)
setColumnWidth(RaceHW.FREQ.2,3:4,11)
setColumnWidth(RaceHW.FREQ.2,5:7,13)
setColumnWidth(RaceHW.FREQ.2,8:9,15)
addDataFrame(x = RaceHW.FREQ, sheet = RaceHW.FREQ.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceHW.FREQ.2,colSplit = 3,rowSplit = 2) # Freeze Panes


# Race HW RvR ---------------------------------------------------------
# Knowledge
task_Race_HW <- task[which(task$Race == 3 | task$Race == 6),]
RaceHW.RvR <- data.frame()
if(ScaleType_RvR == "LIKERT") {
suppressWarnings(
for (i in Statements_KSAOs$Number[1:length(VarLab_RvR_Task)]){
  if(min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("RvR_",i)]]))) != 0 & min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("RvR_",i)]]))) > 1){
    test2 <- leveneTest(task_Race_HW[[paste0("RvR_",i)]], as.factor(task_Race_HW$Race), center=mean)
    row2  <- c(test2$`F value`[1],test2$`Pr(>F)`[1])
    x1    <- tapply(task_Race_HW[[paste0("RvR_",i)]], task_Race_HW$Race, mean, na.rm=TRUE)#SMD Calculation
    x2    <- tapply(task_Race_HW[[paste0("RvR_",i)]], task_Race_HW$Race, sd  , na.rm=TRUE)
    x3    <- table(task_Race_HW$Race)
    eq1 <- ((x3[1] - 1) * x2[1]) + ((x3[2] - 1) * x2[2]);
    SMD <- (x1[1] - x1[2]) / sqrt(eq1 / (x3[1] + x3[2] - 2))
    SMD <- abs(SMD)
    sig_SMD <- ifelse(SMD >= STAT_SMD_Cut, 1,0)
    
    ifelse(row2[2] <= STAT_L_pvalue,
           {test1 <- t.test(task_Race_HW[[paste0("RvR_",i)]] ~ task_Race_HW$Race,var.equal=FALSE)
           sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
           row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
           ,
           {test1 <- t.test(task_Race_HW[[paste0("RvR_",i)]] ~ task_Race_HW$Race,var.equal=TRUE)
           sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
           row   <- c(round(test1$statistic,digits=4),round(test1$p.value,digits=4),test1$estimate,sig, SMD,sig_SMD)}
    )
    row3  <- c(row2,row)
    RaceHW.RvR <- rbind(RaceHW.RvR,row3)}
  else if (min(rowMeans(table(task_Race_HW$Race,task_Race_HW[[paste0("RvR_",i)]]))) == 0 | min(rowSums(table(task_Race_HW$Race,task_Race_HW[[paste0("RvR_",i)]]))) <= 1)
  {RaceHW.RvR <- rbind(RaceHW.RvR,rep(x="NA",times=9))}
})
suppressWarnings(RaceHW.RvR   <- sapply(X = RaceHW.RvR,2,FUN = as.numeric))
RaceHW.RvR                    <- round(x = RaceHW.RvR,digits = 3)
RaceHW.RvR                    <- cbind(Statements_KSAOs[1:length(VarLab_RvR_Task),],RaceHW.RvR)
names(RaceHW.RvR) <- c("KNOW","Description","Levene's F","Levene P","T Statistic","T P-Value","Hispanic Mean","White Mean",paste0("Sig @ ", STAT_pvalue), "SMD",paste0("Cut @ ",STAT_SMD_Cut))
RaceHW.RvR.2                      <- createSheet(wb = SigAnalyses_Race,sheetName = "HW.TT.RvR")
setColumnWidth(RaceHW.RvR.2,1,7)
setColumnWidth(RaceHW.RvR.2,2,115)
setColumnWidth(RaceHW.RvR.2,3:4,11)
setColumnWidth(RaceHW.RvR.2,5:7,13)
setColumnWidth(RaceHW.RvR.2,8:9,15)
addDataFrame(x = RaceHW.RvR, sheet = RaceHW.RvR.2, startRow=1, startColumn=1,colStyle = dfColIndex_TT_Race,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Race,row.names = FALSE)
createFreezePane(sheet = RaceHW.RvR.2,colSplit = 3,rowSplit = 2) # Freeze Panes
saveWorkbook(wb = SigAnalyses_Race,file = paste0(getwd(),"/","RJAQ/OUTPUT/Significance_Analyses/SignificanceAnalyses_Race.xlsx"))
}

#DICHOT SCALE
if(ScaleType_RvR == "DICHOT"){
  RaceHW.RvR <- data.frame()
  suppressWarnings(
    for (i in Statements_KSAOs$Number[1:length(VarLab_RvR_Task)]){
      if(length(table(task_Race_HW[[paste0("RvR_",i)]])) == 2){
        test1 <- chisq.test(table(task_Race_HW$Race,task_Race_HW[[paste0("RvR_",i)]]),correct = FALSE)
        test2 <- fisher.test(table(task_Race_HW$Race,task_Race_HW[[paste0("RvR_",i)]]))
        sig   <- ifelse(test1$p.value <= STAT_pvalue,1,0)
        sig2  <- ifelse(test2$p.value <= STAT_pvalue,1,0)
        a1    <- round((test1$observed[3] / (test1$observed[1] + test1$observed[3])) * 100, digits=2)
        a2    <- round((test1$observed[4] / (test1$observed[2] + test1$observed[4])) * 100, digits=2)
        row   <- c(round(test1$statistic,digits=3),round(test1$p.value,digits=3),round(test2$p.value,digits=3),a1,a2,sig,sig2)
        RaceHW.RvR <- rbind(RaceHW.RvR,row)}
      else if(length(table(task_Race_HW[[paste0("RvR_",i)]])) != 2){
        row <- rep(x = "NA",times = 7)
        RaceHW.RvR <- rbind(RaceHW.RvR,row)
      }
    })
  suppressWarnings(RaceHW.RvR <- sapply(X = RaceHW.RvR,2,FUN = as.numeric))
  RaceHW.RvR <- cbind(Statements_KSAOs[1:length(VarLab_RvR_Task),],RaceHW.RvR)
  names(RaceHW.RvR) <- c("KSAO","Description","Pearson's Chi-Square","Chi-Square P-Value","FET P-Value","Black Mean","White Mean",paste0("Pearson Sig @ ", STAT_pvalue),paste0("Fisher's Sig @ ",STAT_pvalue))
  RaceHW.RvR.2     <- createSheet(wb = SigAnalyses_Race,sheetName = "CHI.RvR")
  setColumnWidth(RaceHW.RvR.2,1,7)
  setColumnWidth(RaceHW.RvR.2,2,115)
  setColumnWidth(RaceHW.RvR.2,3:4,20)
  setColumnWidth(RaceHW.RvR.2,5:7,15)
  setColumnWidth(RaceHW.RvR.2,8:9,18)
  addDataFrame(x = RaceHW.RvR, sheet = RaceHW.RvR.2, startRow=1, startColumn=1,colStyle = dfColIndex_CHI_Gender,showNA = TRUE, colnamesStyle = TABLE_COLNAMES_STYLE_Gender,row.names = FALSE)
  createFreezePane(sheet = RaceHW.RvR.2,colSplit = 3,rowSplit = 2) # Freeze Panes
  saveWorkbook(wb = SigAnalyses_Race,file = paste0(getwd(),"/","RJAQ/OUTPUT/Significance_Analyses/SignificanceAnalyses_Race.xlsx"))
}

# Clean Up Variables 
# OVERALL Significance Calculations --------------------------------

Findings <- data.frame(Scales = c("Required Scale","Essentiality Scale","Applicability Scale","Differentiation Scale","Importance Scale","Frequency Scale","Reference v. Recall Scale"),
                       Scale_Suffix = c("REQU","ESS","APP","DIFF","IMP","FREQ","RvR"),
                       Type_of_Test = c("CHI","CHI","CHI",ifelse(ScaleType_DIFF == "DICHOT","CHI","T-TEST"),"T-TEST","T-TEST",ifelse(ScaleType_RvR == "DICHOT","CHI","T-TEST")),
                       JAQ_Sections = c("TASK/KSAO","TASK","TASK/KSAO","KSAO","TASK/KSAO","TASK","KNOWLEDGE"),
                       Gender_Findings = c(sum(Gender.REQU$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(Gender.ESS$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(Gender.APP$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(Gender.DIFF$`Pearson Sig @ 0.01`,na.rm = TRUE)
                                            ,sum(Gender.IMP$`Sig @ 0.01`,na.rm = TRUE),sum(Gender.FREQ$`Sig @ 0.01`,na.rm = TRUE),sum(Gender.RvR$`Sig @ 0.01`,na.rm = TRUE)),
                       RaceBW_Findings = c(sum(RaceBW.REQU$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(RaceBW.ESS$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(RaceBW.APP$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(RaceBW.DIFF$`Pearson Sig @ 0.01`,na.rm = TRUE)
                                           ,sum(RaceBW.IMP$`Sig @ 0.01`,na.rm = TRUE),sum(RaceBW.FREQ$`Sig @ 0.01`,na.rm = TRUE),sum(RaceBW.RvR$`Sig @ 0.01`,na.rm = TRUE)),
                       RaceHW_Findings = c(sum(RaceHW.REQU$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(RaceHW.ESS$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(RaceHW.APP$`Pearson Sig @ 0.01`,na.rm = TRUE),sum(RaceHW.DIFF$`Pearson Sig @ 0.01`,na.rm = TRUE)
                                           ,sum(RaceHW.IMP$`Sig @ 0.01`,na.rm = TRUE),sum(RaceHW.FREQ$`Sig @ 0.01`,na.rm = TRUE),sum(RaceHW.RvR$`Sig @ 0.01`,na.rm = TRUE)))
