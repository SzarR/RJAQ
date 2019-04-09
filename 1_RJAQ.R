# 0: Set WD / Packages / Options -----------------------------------------------
rm(list=ls()) #Clear environment to ensure no conflict with previous datasets.
setwd(paste0("C:/Users/", Sys.info()["user"],"/Documents/")) #Set the saving points to the user's local directory. 
#File clean-up.
file.remove(list.files("~/RJAQ/OUTPUT",full.names = TRUE,pattern=".xlsx"))
file.remove(list.files("~/RJAQ/OUTPUT/Significance_Analyses/Gender",full.names = TRUE,pattern=".xlsx"))
file.remove(list.files("~/RJAQ/OUTPUT/Significance_Analyses/Race",full.names = TRUE,pattern=".xlsx"))

#For Haven SPSS read in issues, run this.
#devtools::install_github("hadley/haven", force=TRUE)

RJAQ_Version <- c("2.0.1") #RJAQ version control.
list.of.packages <- c("car","reshape","ggplot2","haven","xlsx","readxl") #xlsx is a new package, that helps dramatically with stuff.
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages); sapply(list.of.packages,require,character.only=TRUE)
rm(list.of.packages,new.packages) #Clean Up

# 1: Set QC Options/Settings ---------------------------------------------------
QC1_ScaleCheck_Tasks    <- TRUE #Default:TRUE, If task was indicated as NA, automatically NA all other relevant scales for that task statement for that SME. 
QC1_ScaleCheck_KSAs     <- TRUE #Default:TRUE, If KSA was indicated as NA, automatically NA all other relevant scales for that KSA statement for that SME.
QC2_Attentive_JAQ       <- FALSE #Default:FALSE, If Attentiveness items detected, automatically calculates percentage correct among SMEs by utilizing MODE for correct response.
QC3_Missing_JAQ         <- FALSE #If Missing > QC3_Cutoff_Missing, then "DROP" candidate. JAQ.
QC3_Missing_LAQ         <- FALSE #If Missing > QC3_Cutoff_Missing, then "DROP" candidate. LAQ.
QC4_Variance_JAQ        <- FALSE #If Variance > QC4_Cutoff_VAR, then "DROP" candidate. JAQ.
QC4_Variance_LAQ        <- FALSE #If Variance > QC4_Cutoff_VAR, then "DROP" candidate. LAQ.
QC2_Cutoff_ATT          <- 30.00 #Percentage cut-off for attentiveness items (IMP/FREQ/NA only)
QC3_Cutoff_Missing      <- 70.00 #Percentage cut-off for Missing data across IMP/FREQ/NA scales.
#QC4 runs in EITHER/OR fashion. If one of the two gets hit with TRUE, then removal.
QC4_Cutoff_VAR          <- 0.00  #Cut-off for variance row-wise by scale (IMP/FREQ/NA) or -4SD
QC4_Cutoff_Z            <- -4.00 #Cut-off for variance row-wise by scale (IMP/FREQ/NA) or -4SD
SignificanceTesting     <- FALSE #Default: FALSE. Significance Testing Race/Gender. This links to the 2_SignificanceTesting R Script.

# 2: Read Task/Link/XLSX Files ---------------------------------------------------
# The SPSS files can now be named anything. We check against the colnames to determine
# if it is a linkage file or a task analysis file. The key variables for this are 
# "FREQ_1 to determine JAQ and "SAAL_IMP_1" to determine linkage file.

#Extracts any .SAV files from the userData folder in My Documents.
SAV_Files <- list.files(path = paste0("~/RJAQ/userData/"),pattern = ".sav",ignore.case = TRUE)
if(length(SAV_Files) == 2){File_1 <- read_sav(file= paste0("~/RJAQ/userData/",SAV_Files[1])); File_2 <- read_sav(file=paste0("~/RJAQ/userData/", SAV_Files[2]))}
if(length(SAV_Files) == 1){File_1 <- read_sav(file= paste0("~/RJAQ/userData/",SAV_Files[1]))}

#When 2 SAV files are loaded in the userData folder, this code executes.
if(length(SAV_Files) == 2){
  if("FREQ_1" %in% colnames(File_1)){task <- File_1; rm(File_1)} else {link <- File_1; rm(File_1)}
  if("SAAL_IMP_1" %in% colnames(File_2)){link <- File_2; rm(File_2)} else {task <- File_2; rm(File_2)}
}

#When 1 SAV file is loaded in the userData folder, this code executes.
if(length(SAV_Files) == 1){
  if("FREQ_1" %in% colnames(File_1)){task <- File_1; rm(File_1)} else {link <- File_1; rm(File_1)}
}

#Easier to call JAQ_Presence or LAQ_Presence to check for existence of file.
JAQ_Presence <- exists("task")
LAQ_Presence <- exists("link")

#Read in of JAQ.XLSX file here.
#New code imputed that set the endRow to be equal to the first row that has NA for columns 1 & 2.
#Must read  in twice in order to be able to specify endRow.
if(JAQ_Presence == TRUE){
  Statements    <- read.xlsx(file = "~/RJAQ/userData/JAQ.xlsx",sheetIndex = 1,header = TRUE,colClasses = c("numeric","character"),stringsAsFactors=FALSE,encoding = "UTF-8")
  Statements    <- read.xlsx(file = "~/RJAQ/userData/JAQ.xlsx",sheetIndex = 1,header = TRUE,colClasses = c("numeric","character"),stringsAsFactors=FALSE,encoding = "UTF-8",
                              endRow = if(any(rowSums(is.na(Statements)) == 2) == TRUE) {min(which(rowSums(is.na(Statements)) == 2))})
  suppressWarnings(if(length(attr(task$NA_1,"labels"))==2 | "q00" %in% colnames(task)) {Survey_Monkey = TRUE} else {Survey_Monkey = FALSE})
}

#Read in of LAQ.xlsx file here.
#Must read  in twice in order to be able to specify endRow.
#New code imputed that set the endRow to be equal to the first row that has NA for columns 1 & 2.
if(LAQ_Presence == TRUE){
  Statements_LAQ        <- read.xlsx(file = "~/RJAQ/userData/LAQ.xlsx",header = TRUE,sheetIndex = 1,colClasses = c("numeric","character"), stringsAsFactors=FALSE,encoding = "UTF-8")
  Statements_LAQ        <- read.xlsx(file = "~/RJAQ/userData/LAQ.xlsx",header = TRUE,sheetIndex = 1,colClasses = c("numeric","character"), stringsAsFactors=FALSE,encoding = "UTF-8",
                                     endRow = if(any(rowSums(is.na(Statements_LAQ)) == 2) == TRUE) {min(which(rowSums(is.na(Statements_LAQ)) == 2))})
                                       
  }
  
# 3: XLSX Output Booklet for JAQ -------------------------------------------------
# Creates the column formatting and XLSX booklets for eventual storage of dataset.

if(JAQ_Presence == TRUE){
JAQ_Workbook <- createWorkbook(type="xlsx")
# Styles for the data table row/column names
TABLE_ROWNAMES_STYLE <- CellStyle(JAQ_Workbook) + Font(JAQ_Workbook, isBold=TRUE) + Alignment(horizontal= "ALIGN_CENTER") + 
                        Border(color="black", position=c("TOP", "BOTTOM","LEFT","RIGHT"), pen=c("BORDER_THIN"))
TABLE_COLNAMES_STYLE <- CellStyle(JAQ_Workbook) + Fill(foregroundColor = "dodgerblue4")+ Font(JAQ_Workbook, isBold=TRUE,name = "Calibri",color = "azure") +
                        Alignment(wrapText=FALSE, horizontal="ALIGN_CENTER") + Border(color="lightgrey", position=c("TOP", "BOTTOM","LEFT","RIGHT"), 
                        pen=c("BORDER_THIN", "BORDER_THICK","BORDER_THIN","BORDER_THIN"))
ROWS                 <- CellStyle(JAQ_Workbook) + Font(wb = JAQ_Workbook,name="Calibri",heightInPoints = 10) + Alignment(horizontal = "ALIGN_CENTER",wrapText = TRUE,vertical = "VERTICAL_CENTER") + 
                        Border(color="black",position=c("TOP","LEFT","RIGHT","BOTTOM"), pen=c("BORDER_THIN"))
}

# 4: Set Specifications for read-in -----------------------------------------------------------------------------------------------

Count_Total_Tasks        <- ncol(task[,grepl("FREQ_",names(task))]) #Since frequency is ONLY used for tasks. We count how many times we observe a variable with name "FREQ_X"
Count_Total_KSAOs        <- ncol(task[,grepl("IMP_",names(task))]) - ncol(task[,grepl("FREQ_",names(task))]) #IMP - FREQ = TOTAL KSAO Tasks.
Count_RvR_Task           <- as.numeric(ncol(task[,grepl("RvR_",names(task))])) #Reference v. Recall Scale
Count_IMP_Task           <- as.numeric(ncol(task[,grepl("IMP_",names(task))])) #Importance Scale
Count_NA_Task            <- as.numeric(ncol(task[,grepl("NA_",names(task))]))  #Applicability Scale
Count_FREQ_Task          <- as.numeric(ncol(task[,grepl("FREQ_",names(task))])) #Frequency Scale
Count_REQU_Task          <- as.numeric(ncol(task[,grepl("REQU_",names(task))])) #REQU Scale
Count_DIFF_Task          <- as.numeric(ncol(task[,grepl("DIFF_",names(task))])) #Differentiation Scale

#Create Variable Names to call up the variables within the dataset.
VarLab_NA_Task           <- paste0("NA_",1:Count_NA_Task)
VarLab_IMP_Task          <- paste0("IMP_",1:Count_IMP_Task)
VarLab_FREQ_Task         <- paste0("FREQ_",1:Count_FREQ_Task)
VarLab_REQU_Task         <- paste0("REQU_",1:Count_REQU_Task)
if(Count_Total_KSAOs == Count_REQU_Task) {VarLab_REQU_Task <- paste0("REQU_",(Count_Total_Tasks + 1):Count_Total_KSAOs)} #If REQU not used for tasks but only KSAOs. This executes.
VarLab_DIFF_Task         <- names(task[,grepl("DIFF_",names(task))])
VarLab_RvR_Task          <- names(task[,grepl("RvR_",names(task))])

# Determine Scale Types
ScaleType_REQU  <- ifelse(max(task[,VarLab_REQU_Task],na.rm=T) - min(task[,VarLab_REQU_Task],na.rm=T) == 1, "DICHOT","LIKERT")
ScaleType_DIFF  <- ifelse(max(task[,VarLab_DIFF_Task],na.rm=T) - min(task[,VarLab_DIFF_Task],na.rm=T) == 1, "DICHOT","LIKERT")
if(Count_RvR_Task > 0){ScaleType_RvR   <- ifelse(max(task[,VarLab_RvR_Task],na.rm=T) - min(task[,VarLab_RvR_Task],na.rm=T) == 1, "DICHOT","LIKERT")}

# 5: Process Attentiveness Scale Questions -----------------------------------
# Always turned on. If no attentiveness items present, then no transformations to the datasets will occur.
ATT_All_Row           <- grep(pattern = "Are you paying attention",x = Statements$Description,ignore.case = TRUE,value = FALSE) #Are there any attentiveness items in JAQ.XLSX?

#Print warning if ATT_All_Row > 0 and QC2_Attentive_JAQ == FALSE
if(length(ATT_All_Row) > 0 & QC2_Attentive_JAQ == FALSE){warning("There seem to be Attentiveness Scale Questions in your JAQ.XLSX file, but you have not turned on QC2_Attentive_JAQ in line 18 above.")}
ATT_Tasks_Row         <- NULL
ATT_Tasks_Num         <- NULL
if(length(ATT_All_Row) > 0){
ATT_All_Num           <- Statements$Number[ATT_All_Row]  
ATT_KSAOs_Row         <- ATT_All_Row[ATT_All_Row > which(Statements$Number == Count_Total_Tasks,arr.ind=TRUE)] #Pull out the relevant KSAO number as presented on the JAQ. KSAO Statements ONLY!
ATT_KSAOs_Num         <- Statements$Number[ATT_KSAOs_Row]
ATT_Ks_Row            <- ATT_KSAOs_Row[ATT_KSAOs_Row <= (Count_Total_Tasks + Count_RvR_Task)]
ATT_Ks_Num            <- Statements$Number[ATT_Ks_Row]
ATT_Tasks_Row         <- ATT_All_Row[ATT_All_Row <= which(Statements$Number == Count_Total_Tasks,arr.ind=TRUE)] #Corresponds to the row of the task statement number that is the attenvieness scale item.
ATT_Tasks_Num         <- Statements$Number[ATT_Tasks_Row] 
Statements            <- Statements[-ATT_All_Row,] 
}

DutyAreas_Count          <- sum(is.na(Statements$Number)) - 2 # We are subtracting two. One for NA's associated with Knowledge, and one for the NAs associated with SAOs.
if(any(grepl(pattern = "Knowledge Areas", x = Statements[which(is.na(Statements$Number)),2])) == FALSE) {DutyAreas_Count <- sum(is.na(Statements$Number)) - 1} # We are subtracting one. If there are no Knowledge areas in JAQ. #Added 6.5.2018. #Update 11.29.2018 - Cannot do this. What if you have KA without RVR?
if(any(grepl(pattern = "Skills and Abilities", x = Statements[which(is.na(Statements$Number)),2])) == FALSE) {DutyAreas_Count <- sum(is.na(Statements$Number)) - 1} 
DutyAreas_Names_Full     <- Statements[which(is.na(Statements$Number)),2][1:DutyAreas_Count]
DutyAreas_Names_Acro     <- gsub("[:a-z:]","",DutyAreas_Names_Full) #Remove first capital letter.
DutyAreas_Names_Acro     <- gsub("[[:space:]]", "", DutyAreas_Names_Acro) # remove all the white spaces within strings. 
DutyAreas_Start          <- Statements[which(is.na(Statements$Number)) + 1,]$Number #This vector identifies where a new duty area begins in terms of task statement number.
DutyAreas_Tasks         <- NULL; for (i in 1:DutyAreas_Count) { DutyAreas_Tasks[i] <- DutyAreas_Start[i+1] - DutyAreas_Start[i] ;rm(i) }
Duty_Area_Outline        <- data.frame(Duty.Area = DutyAreas_Names_Full, DA.Acronym = DutyAreas_Names_Acro, DA.Tasks = DutyAreas_Tasks, Percent.Total.JAQ = paste0(round(((DutyAreas_Tasks/Count_Total_Tasks)*100),digits = 2),"%"))             
Statements_Tasks         <- subset(x = Statements,subset = Statements$Number <= Count_Total_Tasks) #Does not include attentiveness questions.
Statements_KSAOs         <- subset(x = Statements,subset = Statements$Number > Count_Total_Tasks)  #Does not include attentiveness questions.

if(length(ATT_All_Row) > 0) {
  VarLab_NA_Task      <- VarLab_NA_Task[-ATT_All_Num]
  VarLab_IMP_Task     <- VarLab_IMP_Task[-ATT_All_Num]
  if(length(ATT_Tasks_Num) > 0){VarLab_FREQ_Task  <- VarLab_FREQ_Task[-ATT_Tasks_Num]} #Update RJAQ v2.0.2 - If NULL, original code would wipe out all variable labels.
  VarLab_REQU_Task    <- VarLab_REQU_Task[-ATT_All_Num]
  if(length(ATT_KSAOs_Num) > 0){VarLab_DIFF_Task    <- VarLab_DIFF_Task[-ATT_KSAOs_Num + max(Statements_Tasks$Number)]} #must add the constant. 
  if(length(ATT_Ks_Num) > 0) {VarLab_RvR_Task     <- VarLab_RvR_Task[-ATT_Ks_Num + max(Statements_Tasks$Number)]} #Update RJAQ v2.0.2 - If NULL, original code would wipe out all variable labels.
}

# -------Survey Monkey Recoding Procedure ----------------
if (Survey_Monkey == TRUE) {
  #NA Recoding. Three possible levels in SurveyMonkey, 1 = YES, 2 = NO, NaN = MISSING. 
  #Remark is coded 999 = APPLICABLE, 1 = NOT APPLICABLE. There is no option for MISSING in Remark.
  is.nan.data.frame <- function(x) {do.call(cbind, lapply(x, is.nan))}    #Create custom function to detect NaN in a dataframe.
  task[,VarLab_NA_Task][is.nan(task[,VarLab_NA_Task])] <- NA              #Step 1: Recode NaN to NA. 
  myRecoder <- function(x){recode(x,"1=1; 2=0")}                          #Step 2: 1 -> 1 and 2 -> 0. Coercing to Remark Format.
  task[,VarLab_NA_Task] <- sapply(task[,VarLab_NA_Task],myRecoder)
} else {
  # --------Relevancy Recoder ------
  task[,VarLab_NA_Task][is.na(task[,VarLab_NA_Task])] <- 0 #Replace NAs to 0s for NA variables. #Step 1: NA -> 0
  myRecoder <- function(x){recode(x,"0=1; 1=0")}                          #Step 2: 0 -> 1 and 1 -> 0
  task[,VarLab_NA_Task] <- sapply(task[,VarLab_NA_Task],myRecoder)
}

# 6: Quality Control Module -------------------------------------------
# Runs only if QC2, QC3, or QC4 are TRUE.
if(QC2_Attentive_JAQ == TRUE | QC3_Missing_JAQ == TRUE  | QC4_Variance_JAQ == TRUE){
QualityControl              <- subset(x = task,select = c("Gender","Race","Rank","Tenure"))
QualityControl$Tenure       <- as.numeric(QualityControl$Tenure)
QualityControl$Gender       <- as_factor(QualityControl$Gender)
QualityControl$Race         <- as_factor(QualityControl$Race)
QualityControl$Rank         <- as_factor(QualityControl$Rank)
}
# QC2: Attentiveness Scale Questions ----------------------------------
#Create Mode Function
#If they get these three then we can safely assume they are paying attention.
Mode <- function(x) {
  ux <- unique(x)
  ux[which.max(tabulate(match(x, ux)))]}

#Run Attentiveness Mode Answer Choice Extraction Technique
#Changed it to two conditions that need to evaluate to TRUE for execution.
if(QC2_Attentive_JAQ == TRUE & length(ATT_All_Row) > 0){

#Importance Scale QC
QC2_IMP              <- paste0("IMP_",ATT_All_Num) #Task + KSAO
QC2_IMP_Key          <- sapply(na.omit(task[QC2_IMP]),Mode)

#REQU Scale QC
QC2_REQU             <- paste0("REQU_",ATT_All_Num) #Task + KSAO
QC2_REQU_Key         <- sapply(na.omit(task[QC2_REQU]),Mode)

#Applicability Scale QC
QC2_NA               <- paste0("NA_",ATT_All_Num) #Task + KSAO
QC2_NA_Key           <- sapply(na.omit(task[QC2_NA]),Mode) 

QC2_Questions        <- c(QC2_IMP,QC2_NA,QC2_REQU)
QC2_Key              <- c(QC2_IMP_Key, QC2_NA_Key, QC2_REQU_Key)

#Create a subset dataframe consisting of QC questions
QC2_Task <- task[QC2_Questions]

#Recode NA to 0.
QC2_Task <- replace(QC2_Task,is.na(QC2_Task),0)

#Scoring Algorithm
QC2_Scoring <- NULL
for(i in names(QC2_Task)){
  Iteration <- ifelse(QC2_Task[i] == QC2_Key[i],1,0)
  QC2_Scoring <- cbind(QC2_Scoring,Iteration)}
rm(Iteration)
QC2_Scoring <- as.data.frame(QC2_Scoring)
QC2_Scoring <- cbind(QC2_Scoring,rowSums(QC2_Scoring),((rowSums(QC2_Scoring) / length(QC2_Questions))*100))
colnames(QC2_Scoring) <- c(QC2_Questions,"ATT_Total_Raw","ATT_Total_Percent")

#Make a CUT!
QC2_Scoring$ATT_QC_2 <- ifelse(QC2_Scoring[,"ATT_Total_Percent"] < QC2_Cutoff_ATT,"DROP","KEEP") #RJAQ v2.0.2 inserted comma into data frame.

#Update QC_Overall
QualityControl <- cbind(QualityControl,subset(x = QC2_Scoring,select = c("ATT_Total_Percent","ATT_QC_2")))}

# QC3: Missingness Cut -------------------------------------------------------
if(QC3_Missing_JAQ == TRUE){
# IMP/NA/FREQ Scales
QualityControl$QC3_IMP_PercentNA     <- round(((rowSums(is.na(task[VarLab_IMP_Task]))/(Count_Total_Tasks+Count_Total_KSAOs))*100),digits=2)
QualityControl$QC3_FREQ_PercentNA    <- round(((rowSums(is.na(task[VarLab_FREQ_Task]))/Count_Total_Tasks)*100),digits=2)
if(Count_REQU_Task > 0){QualityControl$QC3_REQU_PercentNA     <- round(((rowSums(is.na(task[VarLab_REQU_Task]))/(Count_Total_Tasks+Count_Total_KSAOs))*100),digits=2)}
if(Count_DIFF_Task > 0) {QualityControl$QC3_DIFF_PercentNA      <- round(((rowSums(is.na(task[VarLab_DIFF_Task]))/Count_Total_KSAOs)*100),digits=2)}

QualityControl$MIS_QC_3 <- ifelse(QualityControl$QC3_IMP_PercentNA <= QC3_Cutoff_Missing | QualityControl$QC3_FREQ_PercentNA <= QC3_Cutoff_Missing,  "KEEP", "DROP")
if(Count_REQU_Task > 0){QualityControl$MIS_QC_3 <- ifelse(QualityControl$QC3_IMP_PercentNA <= QC3_Cutoff_Missing | QualityControl$QC3_FREQ_PercentNA <= QC3_Cutoff_Missing | QualityControl$QC3_REQU_PercentNA <= QC3_Cutoff_Missing,  "KEEP", "DROP")} 
if(Count_DIFF_Task > 0){QualityControl$MIS_QC_3 <- ifelse(QualityControl$QC3_IMP_PercentNA <= QC3_Cutoff_Missing | QualityControl$QC3_FREQ_PercentNA <= QC3_Cutoff_Missing | QualityControl$QC3_DIFF_PercentNA <= QC3_Cutoff_Missing,  "KEEP", "DROP")} 
if(Count_DIFF_Task > 0 & Count_REQU_Task > 0){QualityControl$MIS_QC_3 <- ifelse(QualityControl$QC3_IMP_PercentNA <= QC3_Cutoff_Missing | QualityControl$QC3_FREQ_PercentNA <= QC3_Cutoff_Missing | 
                                              QualityControl$QC3_DIFF_PercentNA <= QC3_Cutoff_Missing | QualityControl$QC3_REQU_PercentNA <= QC3_Cutoff_Missing,  "KEEP", "DROP")}
}

# QC4: Variance Cut -------------------------------------------------------
if(QC4_Variance_JAQ == TRUE){
QualityControl$QC4_VAR_IMP    <- round(apply(task[VarLab_IMP_Task],1,var,na.rm=T),digits=3)
QualityControl$QC4_VAR_IMP[is.na(QualityControl$QC4_VAR_IMP)] <- 0
QualityControl$QC4_VAR_IMP_Z  <- as.numeric(round(scale(QualityControl$QC4_VAR_IMP),digits=3))
QualityControl$QC4_VAR_FREQ   <- round(apply(task[VarLab_FREQ_Task],1,var,na.rm=T),digits=3)
QualityControl$QC4_VAR_FREQ[is.na(QualityControl$QC4_VAR_FREQ)] <- 0
QualityControl$QC4_VAR_FREQ_Z <- as.numeric(round(scale(QualityControl$QC4_VAR_FREQ),digits=3))
if(Count_REQU_Task > 0) {QualityControl$QC4_VAR_REQU <- round(apply(task[VarLab_REQU_Task],1,var,na.rm=T),digits=3);
          QualityControl$QC4_VAR_REQU[is.na(QualityControl$QC4_VAR_REQU)] <- 0
          QualityControl$QC4_VAR_REQU_Z <- as.numeric(round(scale(QualityControl$QC4_VAR_REQU),digits=3))}
if(Count_DIFF_Task > 0)  {QualityControl$QC4_VAR_DIFF  <- round(apply(task[VarLab_DIFF_Task],1,var,na.rm=T),digits=3);
          QualityControl$QC4_VAR_DIFF[is.na(QualityControl$QC4_VAR_DIFF)] <- 0
          QualityControl$QC4_VAR_DIFF_Z <- as.numeric(round(scale(QualityControl$QC4_VAR_DIFF),digits=3))}

QualityControl$VAR_QC_4  <- ifelse(QualityControl$QC4_VAR_IMP == QC4_Cutoff_VAR | QualityControl$QC4_VAR_IMP_Z <= QC4_Cutoff_Z | QualityControl$QC4_VAR_FREQ == QC4_Cutoff_VAR | QualityControl$QC4_VAR_FREQ_Z <= QC4_Cutoff_Z,  "DROP", "KEEP")
if(Count_REQU_Task > 0) {QualityControl$VAR_QC_4      <- ifelse(QualityControl$QC4_VAR_IMP == QC4_Cutoff_VAR | QualityControl$QC4_VAR_FREQ == QC4_Cutoff_VAR | QualityControl$QC4_VAR_REQU == QC4_Cutoff_VAR | QualityControl$QC4_VAR_REQU_Z <= QC4_Cutoff_Z,  "DROP", "KEEP")}
if(Count_DIFF_Task > 0)  {QualityControl$VAR_QC_4     <- ifelse(QualityControl$QC4_VAR_IMP == QC4_Cutoff_VAR | QualityControl$QC4_VAR_FREQ == QC4_Cutoff_VAR | QualityControl$QC4_VAR_DIFF == QC4_Cutoff_VAR | QualityControl$QC4_VAR_DIFF_Z <= QC4_Cutoff_Z,  "DROP", "KEEP")}
if(Count_DIFF_Task > 0 & Count_REQU_Task > 0) {QualityControl$VAR_QC_4     <- ifelse(QualityControl$QC4_VAR_IMP == QC4_Cutoff_VAR | QualityControl$QC4_VAR_FREQ == QC4_Cutoff_VAR | QualityControl$QC4_VAR_DIFF == QC4_Cutoff_VAR | QualityControl$QC4_VAR_DIFF_Z <= QC4_Cutoff_Z | QualityControl$QC4_VAR_REQU == QC4_Cutoff_VAR | QualityControl$QC4_VAR_REQU_Z <= QC4_Cutoff_Z,  "DROP", "KEEP")}
}

# QC Conclusion -----------------------------------------------------------
# Only QC_2
if(QC2_Attentive_JAQ == TRUE & QC3_Missing_JAQ == FALSE & QC4_Variance_JAQ == FALSE){
  QualityControl$QC_2 <- ifelse(QualityControl$ATT_QC_2 == "KEEP","KEEP","DROP")
  Inattentive_Cases <- which(QualityControl$QC_2 == "KEEP")
  task <- task[Inattentive_Cases,]
}

# QC_2 and QC_3
if(QC2_Attentive_JAQ == TRUE & QC3_Missing_JAQ == TRUE & QC4_Variance_JAQ == FALSE){
  QualityControl$QC_2_3 <- ifelse(QualityControl$ATT_QC_2 == "KEEP" & QualityControl$MIS_QC_3 == "KEEP","KEEP","DROP")
  Inattentive_Cases <- which(QualityControl$QC_2_3 == "KEEP")
  task <- task[Inattentive_Cases,]
}

# QC_2, QC_3, QC_4
if(QC2_Attentive_JAQ == TRUE & QC3_Missing_JAQ == TRUE & QC4_Variance_JAQ == TRUE){
  QualityControl$QC_2_3_4 <- ifelse(QualityControl$ATT_QC_2 == "KEEP" & QualityControl$MIS_QC_3 == "KEEP" & QualityControl$VAR_QC_4 == "KEEP","KEEP","DROP")
  Inattentive_Cases <- which(QualityControl$QC_2_3_4 == "KEEP")
  task <- task[Inattentive_Cases,]
}

# QC_3 and QC_4
if(QC2_Attentive_JAQ == FALSE & QC3_Missing_JAQ == TRUE & QC4_Variance_JAQ == TRUE){
  QualityControl$QC_3_4 <- ifelse(QualityControl$MIS_QC_3 == "KEEP" & QualityControl$VAR_QC_4 == "KEEP","KEEP","DROP")
  Inattentive_Cases <- which(QualityControl$QC_3_4 == "KEEP")
  task <- task[Inattentive_Cases,]
}

# QC_3
if(QC2_Attentive_JAQ == FALSE & QC3_Missing_JAQ == TRUE & QC4_Variance_JAQ == FALSE){
  QualityControl$QC_3 <- ifelse(QualityControl$MIS_QC_3 == "KEEP","KEEP","DROP")
  Inattentive_Cases <- which(QualityControl$QC_3 == "KEEP")
  task <- task[Inattentive_Cases,]
}

# QC_4
if(QC2_Attentive_JAQ == FALSE & QC3_Missing_JAQ == FALSE & QC4_Variance_JAQ == TRUE){
  QualityControl$QC_4 <- ifelse(QualityControl$VAR_QC_4 == "KEEP","KEEP","DROP")
  Inattentive_Cases <- which(QualityControl$QC_4 == "KEEP")
  task <- task[Inattentive_Cases,]
}

# QC XLSX Output -----------------------------------------------------------
if(QC2_Attentive_JAQ == TRUE | QC3_Missing_JAQ == TRUE | QC3_Missing_LAQ == TRUE | QC4_Variance_JAQ == TRUE | QC4_Variance_LAQ == TRUE){
QC_Analysis <- createSheet(JAQ_Workbook, sheetName = "QC_Analysis")
setColumnWidth(QC_Analysis,3,35)
setColumnWidth(QC_Analysis,4,25)
dfColIndex           <- rep(list(ROWS), dim(QualityControl)[2]) 
names(dfColIndex)    <- seq(1, dim(QualityControl)[2], by = 1)

addDataFrame(x = QualityControl, sheet = QC_Analysis, startRow=1, startColumn=1,colStyle = dfColIndex,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)
createFreezePane(sheet = QC_Analysis,colSplit = 3,rowSplit = 2) # Freeze Panes
}

# XLSX Demographic Output -------------------------------------------------
Demo_Sheet  <- createSheet(JAQ_Workbook, sheetName = "Demographics")
if("Assignment" %in% colnames(task))   {addDataFrame(x = table(task$Assignment,dnn = "Assignment"), sheet = Demo_Sheet,startColumn = 1, row.names=FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)}
if("Gender" %in% colnames(task))       {task$Gender <- ifelse(task$Gender == 1,"Male",ifelse(task$Gender == 2,"Female","N/A"))
                                        ;addDataFrame(x = table(task$Gender, dnn = "Gender"),sheet = Demo_Sheet, startColumn = 4,row.names = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)}
if("Race" %in% colnames(task))         {task$Race2   <- ifelse(task$Race == 1, "Black",ifelse(task$Race == 2, "Asian",ifelse(task$Race == 3, "Hispanic", 
                                        ifelse(task$Race == 4, "Pacific", ifelse(task$Race == 5, "Native",ifelse(task$Race == 6, "White",ifelse(task$Race == 7, "Two+", "N/A")))))))
                                        ;addDataFrame(x = table(task$Race2, dnn = "Race"), sheet = Demo_Sheet, startColumn = 7, row.names = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)}
if("AgencyType" %in% colnames(task))   {addDataFrame(x = table(task$AgencyType, dnn = "AgencyType"), sheet = Demo_Sheet, startColumn = 13, row.names = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)}
if("AgencySize" %in% colnames(task))   {addDataFrame(x = table(task$AgencySize, dnn = "AgencySize"), sheet = Demo_Sheet, startColumn = 16, row.names = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)}
if("AgencyRegion" %in% colnames(task)) {addDataFrame(x = table(task$AgencyRegion, dnn = "AgencyRegion"), sheet = Demo_Sheet, startColumn = 16, row.names = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)}

# QC1: Automatic NA Recode (Tasks) ------------------------------------------------
if (QC1_ScaleCheck_Tasks == TRUE) {
  #1 means APP. 0 Means NA. If NA = 0 and IMP, FREQ scales are NA, we are good. Otherwise, print "0" for the QC_Task output.
  QC_Analysis_Tasks      <- ifelse(task[,VarLab_NA_Task][1:length(Statements_Tasks$Number)] == 1, 1,
                            ifelse(task[,VarLab_NA_Task][1:length(Statements_Tasks$Number)] == 0
                            & is.na(task[,VarLab_FREQ_Task][1:length(Statements_Tasks$Number)])
                            & is.na(task[,VarLab_IMP_Task][1:length(Statements_Tasks$Number)]),1,0))

  if(Count_REQU_Task > 0){
    QC_Analysis_Tasks    <- ifelse(task[,VarLab_NA_Task][1:length(Statements_Tasks$Number)] == 1, 1,
                            ifelse(task[,VarLab_NA_Task][1:length(Statements_Tasks$Number)] == 0
                            & is.na(task[,VarLab_FREQ_Task][1:length(Statements_Tasks$Number)])
                            & is.na(task[,VarLab_REQU_Task][1:length(Statements_Tasks$Number)])
                            & is.na(task[,VarLab_IMP_Task][1:length(Statements_Tasks$Number)]),1,0))}

  #Name the matrix.
  dimnames(QC_Analysis_Tasks)[[2]] <- as.list(paste0("QC_",Statements_Tasks$Number))

  #What needs changing?
  QC_Replacements_Tasks <- which(QC_Analysis_Tasks == 0, arr.ind=TRUE)

  #Use the arr ind to make the replacements to NA.
  task[,VarLab_IMP_Task] <- replace(x = task[,VarLab_IMP_Task],list = QC_Replacements_Tasks,values = NA)
  task[,VarLab_FREQ_Task] <- replace(x = task[,VarLab_FREQ_Task],list = QC_Replacements_Tasks,values = NA)
  if(Count_REQU_Task > 0) {task[,VarLab_REQU_Task] <- replace(x = task[,VarLab_REQU_Task],list = QC_Replacements_Tasks,values = NA)}}

if (QC1_ScaleCheck_KSAs == TRUE){

  #1 means APP. 0 Means NA. If NA = 0 and IMP, FREQ scales are NA, we are good. Otherwise, print "0" for the QC_Task output.
  #A 0 is bad news in this analysis, a 0 needs to get acted upon. A 1 is in the clear.
  QC_Analysis_KSA      <- ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 1, 1,
                          ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 0
                          & is.na(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))]),1,0))

  if(Count_REQU_Task > length(Statements_Tasks$Number)){

  QC_Analysis_KSA      <- ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 1, 1,
                          ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 0
                          & is.na(task[,VarLab_REQU_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))]
                          & is.na(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))])),1,0))}

  if(Count_DIFF_Task > 0 & Count_REQU_Task > length(Statements_Tasks$Number)){
  QC_Analysis_KSA      <- ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 1, 1,
                          ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 0
                          & is.na(task[,VarLab_REQU_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))])
                          & is.na(task[,VarLab_IMP_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))])
                          & is.na(task[,VarLab_DIFF_Task]),1,0))}

  if(Count_DIFF_Task > 0){
  QC_Analysis_KSA      <- ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 1, 1,
                          ifelse(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))] == 0
                          & is.na(task[,VarLab_DIFF_Task])
                          & is.na(task[,VarLab_NA_Task][(length(Statements_Tasks$Number)+1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))]),1,0))}

  #Name the matrix.
  dimnames(QC_Analysis_KSA)[[2]] <- as.list(paste0("QC_",(length(Statements_Tasks$Number) + 1):(length(Statements_KSAOs$Number) + length(Statements_Tasks$Number))))

  #What needs changing?
  QC_Replacements_KSA <- which(QC_Analysis_KSA == 0, arr.ind=TRUE)
  QC_Replacements_KSA_2 <- QC_Replacements_KSA
  QC_Replacements_KSA_2[,2] <- QC_Replacements_KSA[,2] + nrow(Statements_Tasks)
  task[,VarLab_IMP_Task] <- replace(x = task[,VarLab_IMP_Task],list = QC_Replacements_KSA_2,values = NA)
  if(Count_REQU_Task > nrow(Statements_Tasks)) {task[,VarLab_REQU_Task] <- replace(x = task[,VarLab_REQU_Task],list = QC_Replacements_KSA_2,values = NA)}
  if(Count_DIFF_Task  > nrow(Statements_Tasks))  {task[,VarLab_DIFF_Task] <- replace(x = task[,VarLab_DIFF_Task],list = QC_Replacements_KSA,values = NA)}}

# Composite Scale ---------------------------------------------------
Composite        <- (task[,paste0("IMP_",Statements_Tasks$Number)] * 2 + task[,paste0("FREQ_",Statements_Tasks$Number)])/3
names(Composite) <- paste0("C_",Statements_Tasks$Number)

# Essentiality Scale ------------------------------------------------
if(max(task[,VarLab_IMP_Task],na.rm=T) == 4 & max(task[,VarLab_FREQ_Task],na.rm=T) == 4) {Essential <- as.data.frame(ifelse(Composite[1:length(Composite)] >= 2.333,1,0))}
if(max(task[,VarLab_IMP_Task],na.rm=T) == 5 & max(task[,VarLab_FREQ_Task],na.rm=T) == 5) {Essential <- as.data.frame(ifelse(Composite[1:length(Composite)] >= 3.000,1,0))}

# -------Required Upon ____ Scale (REQU) -------
if(Count_REQU_Task > 0 & ScaleType_REQU == "DICHOT") {Task_REQU_Sum <- (apply(task[,VarLab_REQU_Task],2,function(x) (sum(x == 1,na.rm=T)) / (sum(x == 2,na.rm=T) + sum(x==1,na.rm=T)))*100); Task_REQU_SD <- sapply(task[,VarLab_REQU_Task],sd,2)}
if(Count_REQU_Task > 0 & ScaleType_REQU == "LIKERT") {Task_REQU_Sum <- apply(task[,VarLab_REQU_Task],2,mean,na.rm=T); Task_REQU_SD <- sapply(task[,VarLab_REQU_Task],sd,2)}

# Reference v. Recall Scale -----------------------------------------------
if(Count_RvR_Task > 0){
if(ScaleType_RvR == "DICHOT") {Task_RvR_Sum  <- apply(task[,VarLab_RvR_Task],2,function(x) (sum(x == 1,na.rm=T)) / (sum(x == 2,na.rm=T) + sum(x==1,na.rm=T))*100); length(Task_RvR_Sum) <- nrow(Statements_KSAOs); Task_RvR_SD <- sapply(task[,VarLab_RvR_Task],sd,2);length(Task_RvR_SD) <- nrow(Statements_KSAOs)}
if(ScaleType_RvR == "LIKERT") {Task_RvR_Sum  <- apply(task[,VarLab_RvR_Task],2,mean,na.rm=T); Task_RvR_SD <- sapply(task[,VarLab_RvR_Task],sd,2); length(Task_RvR_Sum) <- nrow(Statements_KSAOs); length(Task_RvR_SD) <- nrow(Statements_KSAOs)}
} #Modified 6.5.2018. First needed to check Count_RvR_Task >0 and then it launches the second stage of checks.
# Differentiation Scale ------------------------------------------------
if(Count_DIFF_Task > 0 & ScaleType_DIFF == "DICHOT") {Task_DIFF_Sum  <- (apply(task[,VarLab_DIFF_Task],2,function(x) (sum(x == 1,na.rm=T)) / (sum(x == 2,na.rm=T) + sum(x==1,na.rm=T)))*100); Task_DIFF_SD <- sapply(task[,VarLab_DIFF_Task],sd,2)}
if(Count_DIFF_Task > 0 & ScaleType_DIFF == "LIKERT") {Task_DIFF_Sum  <- apply(task[,VarLab_DIFF_Task],2,mean,na.rm=T); Task_DIFF_SD <- sapply(task[,VarLab_DIFF_Task],sd,2)}

# For Remark 
if(Survey_Monkey == FALSE & ScaleType_REQU == "DICHOT"){{Task_REQU_Sum <- (apply(task[,VarLab_REQU_Task],2,function(x) (sum(x == 2,na.rm=T)) / (sum(x == 2,na.rm=T) + sum(x==1,na.rm=T)))*100); Task_REQU_SD <- sapply(task[,VarLab_REQU_Task],sd,2)}}
if(Count_RvR_Task > 0){
if(Survey_Monkey == FALSE & ScaleType_RvR == "DICHOT"){Task_RvR_Sum  <- apply(task[,VarLab_RvR_Task],2,function(x) (sum(x == 2,na.rm=T)) / (sum(x == 2,na.rm=T) + sum(x==1,na.rm=T))*100); length(Task_RvR_Sum) <- nrow(Statements_KSAOs); Task_RvR_SD <- sapply(task[,VarLab_RvR_Task],sd,2);length(Task_RvR_SD) <- nrow(Statements_KSAOs)}
}
if(Survey_Monkey == FALSE & ScaleType_DIFF == "DICHOT"){Task_DIFF_Sum  <- (apply(task[,VarLab_DIFF_Task],2,function(x) (sum(x == 2,na.rm=T)) / (sum(x == 2,na.rm=T) + sum(x==1,na.rm=T)))*100); Task_DIFF_SD <- sapply(task[,VarLab_DIFF_Task],sd,2)}

# --------(ALL) Calculation: NA I F (KSA) -------
if(Count_Total_KSAOs > 0){
KSA_NA_Sum         <- colMeans((task[,VarLab_NA_Task][(nrow(Statements_Tasks) + 1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))])*100, na.rm=TRUE)
KSA_NA_SD          <- sapply(task[,VarLab_NA_Task][(nrow(Statements_Tasks) + 1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))],sd,2)
KSA_IMP_Sum        <- colMeans(task[,VarLab_IMP_Task][(nrow(Statements_Tasks) + 1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))],na.rm=TRUE)
KSA_IMP_SD         <- sapply(task[,VarLab_IMP_Task][(nrow(Statements_Tasks) + 1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))],sd,2)
KSA_ESS_Sum        <- ifelse(KSA_NA_Sum >= 66.67 & KSA_IMP_Sum >= 3.00,1,0)
}

# --------(ALL) Calculation NA I F C E (TASK) ------
Task_NA_Sum        <- (colMeans(task[,VarLab_NA_Task][1:nrow(Statements_Tasks)],na.rm=T)*100)
Task_NA_SD         <- sapply(task[,VarLab_NA_Task][1:nrow(Statements_Tasks)],sd,2)
Task_FREQ_Sum      <- colMeans(task[,VarLab_FREQ_Task][1:nrow(Statements_Tasks)],na.rm=T)
Task_FREQ_SD       <- sapply(task[,VarLab_FREQ_Task][1:nrow(Statements_Tasks)],sd,2)
Task_IMP_Sum       <- colMeans(task[,VarLab_IMP_Task][1:nrow(Statements_Tasks)],na.rm=T)
Task_IMP_SD        <- sapply(task[,VarLab_IMP_Task][1:nrow(Statements_Tasks)],sd,2)
Task_COMP_Sum      <- colMeans(Composite,na.rm=TRUE)
Task_COMP_SD       <- sapply(Composite,sd,2)
Task_ESS_Sum       <- colMeans(Essential,na.rm=T)*100
Task_ESS_SD        <- sapply(Essential,sd,2)

# --------(ALL) Zone Matrix Results -------
Zones    <- ifelse(Task_NA_Sum >= 66.67 & Task_ESS_Sum >= 66.67, 1.1,
            ifelse((Task_NA_Sum >= 66.67 & Task_ESS_Sum >= 50.00 & Task_ESS_Sum < 66.67), 1.2,
            ifelse((Task_NA_Sum >= 50.00 & Task_NA_Sum < 66.67 & Task_ESS_Sum >= 66.67), 1.3,
            ifelse((Task_NA_Sum >= 50.00 & Task_NA_Sum < 66.67 & Task_ESS_Sum >= 50.00 & Task_ESS_Sum < 66.67), 1.4,
            ifelse((Task_NA_Sum >= 66.67 & Task_ESS_Sum >= 33.33 & Task_ESS_Sum < 50), 2.1,
            ifelse((Task_NA_Sum >= 66.67 & Task_ESS_Sum >= 0.00 & Task_ESS_Sum < 33.32), 2.2,
            ifelse((Task_NA_Sum >= 50.00 & Task_NA_Sum < 66.67 & Task_ESS_Sum >= 33.33 & Task_ESS_Sum < 50.00), 2.3,
            ifelse((Task_NA_Sum >= 50.00 & Task_NA_Sum < 66.67 & Task_ESS_Sum >= 0 & Task_ESS_Sum < 33.33), 2.4,
            ifelse((Task_NA_Sum >= 33.33 & Task_NA_Sum < 50.00 & Task_ESS_Sum >= 66.67), 3.1,
            ifelse((Task_NA_Sum >= 33.33 & Task_NA_Sum < 50.00 & Task_ESS_Sum >= 50.00 & Task_ESS_Sum < 66.67), 3.2,
            ifelse((Task_NA_Sum >= 0 & Task_NA_Sum < 33.33 & Task_ESS_Sum >= 66.67), 3.3,
            ifelse((Task_NA_Sum >= 0 & Task_NA_Sum < 33.33 & Task_ESS_Sum >= 50.00 & Task_ESS_Sum < 66.67), 3.4,
            ifelse((Task_NA_Sum >= 33.33 & Task_NA_Sum < 50.00 & Task_ESS_Sum >= 33.33 & Task_ESS_Sum < 50.00), 4.1,
            ifelse((Task_NA_Sum >= 33.33 & Task_NA_Sum < 50.00 & Task_ESS_Sum >= 0 & Task_ESS_Sum < 33.33), 4.2,
            ifelse((Task_NA_Sum >= 0 & Task_NA_Sum < 33.33 & Task_ESS_Sum >= 33.33 & Task_ESS_Sum < 50.00), 4.3,
            ifelse((Task_NA_Sum >= 0 & Task_NA_Sum < 33.33 & Task_ESS_Sum >= 0 & Task_ESS_Sum < 33.33), 4.4, "ATTENTION"
            )))))))))))))))) #BOOM Matrix created by Bob.

# KSA Saving Excel File ---------------------------------------------------
if((nrow(Statements_Tasks) + nrow(Statements_KSAOs)) > nrow(Statements_Tasks)){

# Standard Output KSA -----------------------------------------------------
Final_KSA_Frame               <- as.data.frame(cbind(Statements_KSAOs$Description, round(KSA_NA_Sum,digits = 2),round(KSA_NA_SD,digits=2),round(KSA_IMP_Sum, digits = 2),round(KSA_IMP_SD,digits=2),round(KSA_ESS_Sum,digits=2)),stringsAsFactors = FALSE)
colnames(Final_KSA_Frame)     <- c("Description","APP","APP_SD","IMP","IMP_SD","ESS")

# Standard Output + Reference ---------------------------------------------
if(Count_RvR_Task > 0){
  Final_KSA_Frame              <- as.data.frame(cbind(Statements_KSAOs$Description,round(KSA_NA_Sum,digits=2),round(KSA_NA_SD,digits=2),round(KSA_IMP_Sum,digits=2),round(KSA_IMP_SD,digits=2),round(Task_RvR_Sum,digits=2),round(Task_RvR_SD,digits=2),round(KSA_ESS_Sum,digits=2)),stringsAsFactors = FALSE)
  colnames(Final_KSA_Frame)    <- c("Description","APP","APP_SD","IMP","IMP_SD","REF","REF_SD","ESS")}

# Standard Output + Differentiation ---------------------------------------
if(Count_DIFF_Task > 0){
  Final_KSA_Frame             <- as.data.frame(cbind(Statements_KSAOs$Description,round(KSA_NA_Sum,digits=2),round(KSA_NA_SD,digits=2),round(KSA_IMP_Sum,digits=2),round(KSA_IMP_SD,digits=2),round(Task_DIFF_Sum,digits=2),round(Task_DIFF_SD,digits=2),round(KSA_ESS_Sum,digits=2)),stringsAsFactors = FALSE)
  colnames(Final_KSA_Frame)   <- c("Description","APP","APP_SD","IMP","IMP_SD","DIFF","DIFF_SD","ESS")}

# Standard Output + Differentiation + Reference ---------------------------
if(Count_DIFF_Task > 0 & Count_RvR_Task > 0){
  Final_KSA_Frame              <- as.data.frame(cbind(Statements_KSAOs$Description,round(KSA_NA_Sum,digits=2),round(KSA_NA_SD,digits=2),round(KSA_IMP_Sum,digits=2),round(KSA_IMP_SD,digits=2),round(Task_RvR_Sum,digits=2),round(Task_RvR_SD,digits=2),round(Task_DIFF_Sum,digits=2),round(Task_DIFF_SD,digits=2),round(KSA_ESS_Sum,digits=2)),stringsAsFactors = FALSE)
  colnames(Final_KSA_Frame)   <- c("Description","APP","APP_SD","IMP","IMP_SD","REF","REF_SD","DIFF","DIFF_SD","ESS")}

# Standard Output + Differentiation + Required ----------------------------
if(Count_DIFF_Task > 0 & Count_REQU_Task > 0){
  Final_KSA_Frame              <- as.data.frame(cbind(Statements_KSAOs$Description,round(KSA_NA_Sum,digits=2),round(KSA_NA_SD,digits=2),round(KSA_IMP_Sum,digits=2),round(KSA_IMP_SD,digits=2),round(Task_REQU_Sum[(nrow(Statements_Tasks)+1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))],digits=2),round(Task_REQU_SD[(nrow(Statements_Tasks)+1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))],digits=2),round(Task_DIFF_Sum,digits=2),round(Task_DIFF_SD,digits=2),round(KSA_ESS_Sum,digits=2)),stringsAsFactors = FALSE)
  colnames(Final_KSA_Frame)    <- c("Description","APP","APP_SD","IMP","IMP_SD","REQU","REQU_SD","DIFF","DIFF_SD","ESS")}

# Standard Output + Reference + Required ----------------------------------
if(Count_RvR_Task > 0 & Count_REQU_Task > 0){
  Final_KSA_Frame               <- as.data.frame(cbind(Statements_KSAOs$Description,round(KSA_NA_Sum,digits=2),round(KSA_NA_SD,digits=2),round(KSA_IMP_Sum,digits=2),round(KSA_IMP_SD,digits=2),round(Task_REQU_Sum[(nrow(Statements_Tasks)+1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))],digits=2),round(Task_REQU_SD[(nrow(Statements_Tasks)+1):(nrow(Statements_Tasks) + nrow(Statements_KSAOs))],digits=2),round(Task_RvR_Sum,digits=2),round(Task_RvR_SD,digits=2),round(KSA_ESS_Sum,digits=2)),stringsAsFactors = FALSE)
  colnames(Final_KSA_Frame)     <- c("Description","APP","APP_SD","IMP","IMP_SD","REQU","REQU_SD","REF","REF_SD","ESS")}

# Standard Output + Required + Differentiation + Reference ----------------
if(Count_REQU_Task > 0 & Count_DIFF_Task > 0 & Count_RvR_Task > 0){
  Final_KSA_Frame              <- as.data.frame(cbind(Statements_KSAOs$Description,round(KSA_NA_Sum,digits=2),round(KSA_NA_SD,digits=2),round(KSA_IMP_Sum,digits=2),round(KSA_IMP_SD,digits=2),round(Task_RvR_Sum,digits=2),round(Task_RvR_SD,digits=2),round(Task_DIFF_Sum,digits=2),round(Task_DIFF_SD,digits=2),round(Task_REQU_Sum[(nrow(Statements_Tasks)+1):(nrow(Statements_Tasks)+ nrow(Statements_KSAOs))],digits=2),round(Task_REQU_SD[(nrow(Statements_Tasks)+1):(nrow(Statements_Tasks)+ nrow(Statements_KSAOs))],digits=2),round(KSA_ESS_Sum,digits=2)),stringsAsFactors = FALSE)
  colnames(Final_KSA_Frame)    <- c("Description","APP","APP_SD","IMP","IMP_SD","REF","REF_SD","DIFF","DIFF_SD","REQU","REQU_SD","ESS")}

Final_KSA_Frame[2:ncol(Final_KSA_Frame)]   <- sapply(X = Final_KSA_Frame[2:ncol(Final_KSA_Frame)],FUN = as.numeric)
row.names(Final_KSA_Frame)                 <-  paste0("KSAO_",Statements_KSAOs$Number)

KSAO_Analysis        <- createSheet(JAQ_Workbook, sheetName = "KSAO_Analysis")

dfColIndex           <- rep(list(ROWS), dim(Final_KSA_Frame)[2]) 
names(dfColIndex)    <- seq(1, dim(Final_KSA_Frame)[2], by = 1)

# Add a table
addDataFrame(x = Final_KSA_Frame, sheet = KSAO_Analysis, startRow=1, startColumn=1,colStyle = dfColIndex,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)
setColumnWidth(sheet = KSAO_Analysis,colIndex = 2,colWidth = 90) # Change column width
createFreezePane(sheet = KSAO_Analysis,colSplit = 3,rowSplit = 2) # Freeze Panes
}

# Task Saving Excel File --------------------------------------------------
# Standard Task Output ----------------------------------------------------
if(length(ATT_Tasks_Num) > 0)    {Tasks_Output <- rep(DutyAreas_Names_Acro, Duty_Area_Outline$DA.Tasks)[-ATT_Tasks_Num]}
if(length(ATT_Tasks_Num) == 0)   {Tasks_Output <- rep(DutyAreas_Names_Acro, Duty_Area_Outline$DA.Tasks)}

Final_Task_Frame           <- as.data.frame(cbind(Statements_Tasks$Description,round(Task_NA_Sum,digits=3),round(Task_NA_SD,digits=2),round(Task_IMP_Sum,digits=2),round(Task_IMP_SD,digits=2),round(Task_FREQ_Sum,digits=2),round(Task_FREQ_SD,digits=2),round(Task_COMP_Sum,digits=2),round(Task_COMP_SD,digits=2),round(Task_ESS_Sum,digits=3),Zones,Tasks_Output),stringsAsFactors = FALSE)
colnames(Final_Task_Frame) <- c("Description","APP","APP_SD","IMP","IMP_SD","FREQ","FREQ_SD","COMP","COMP_SD","ESS","Zone.Class","Duty.Area")
#6.5.2018 Changed rounding for NA_SUM and ESS_Sum to 3 digits to not confuse with Zones and rounding.

# Standard Task Output + REQU ---------------------------------------------
#Technically called "Performed upon Promotion"
if(Count_REQU_Task > nrow(Statements_Tasks)){
  Final_Task_Frame           <- as.data.frame(cbind(Statements_Tasks$Description,round(Task_NA_Sum,digits=3),round(Task_NA_SD,digits=2),round(Task_IMP_Sum,digits=2),round(Task_IMP_SD,digits=2),round(Task_FREQ_Sum,digits=2),round(Task_FREQ_SD,digits=2),round(Task_COMP_Sum,digits=2),round(Task_COMP_SD,digits=2),round(Task_REQU_Sum[1:nrow(Statements_Tasks)],digits=2),round(Task_REQU_SD[1:nrow(Statements_Tasks)],digits=2),round(Task_ESS_Sum, digits=3), Zones,Tasks_Output),stringsAsFactors = FALSE)
  colnames(Final_Task_Frame) <- c("Description","APP","APP_SD","IMP","IMP_SD","FREQ","FREQ_SD","COMP","COMP_SD","PUP","PUP_SD","ESS","Zone.Class","Duty.Area")
  Final_Task_Frame[2:(ncol(Final_Task_Frame)-1)] <- sapply(X = Final_Task_Frame[2:(ncol(Final_Task_Frame)-1)],FUN = as.numeric)
  #6.5.2018 Changed rounding for NA_SUM and ESS_Sum to 3 digits to not confuse with Zones and rounding.
  }

row.names(Final_Task_Frame)   <-  paste0("TASK_",Statements_Tasks$Number)
Task_Analysis <- createSheet(JAQ_Workbook, sheetName = "Task_Analysis")

dfColIndex           <- rep(list(ROWS), dim(Final_Task_Frame)[2]) 
names(dfColIndex)    <- seq(1, dim(Final_Task_Frame)[2], by = 1)

# Add a table
addDataFrame(x = Final_Task_Frame, sheet = Task_Analysis, startRow=1, startColumn=1,colStyle = dfColIndex,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)
setColumnWidth(sheet = Task_Analysis,colIndex = 2,colWidth = 90) # Change column width
createFreezePane(sheet = Task_Analysis,colSplit = 3,rowSplit = 2) # Freeze Panes

# Importance, Frequency, Composite Output by Duty Area --------------------
IFC_By_DutyArea <- matrix(ncol = 10,nrow = DutyAreas_Count);rownames(IFC_By_DutyArea) <- DutyAreas_Names_Full; colnames(IFC_By_DutyArea) <- c("APP","APP_SD","IMP","IMP_SD","FREQ","FREQ_SD","COMP","COMP_SD","PUP","PUP_SD")
dat             <- NULL

for (i in 1:DutyAreas_Count){
 dat                       <- Final_Task_Frame[Final_Task_Frame$Duty.Area == DutyAreas_Names_Acro[i],]
 IFC_By_DutyArea[i,]       <- c(mean(dat$APP),mean(dat$APP_SD),mean(dat$IMP),mean(dat$IMP_SD),mean(dat$FREQ),mean(dat$FREQ_SD),mean(dat$COMP),mean(dat$COMP_SD),mean(dat$PUP),mean(dat$PUP_SD))
}

# Comprehensiveness Question ----------------------------------------------
Percent <- as.data.frame(mean(task$Percent,na.rm=TRUE))
colnames(Percent) <- c("Comprehensiveness Rating")
row.names(Percent) <- c("How Comprehensive was this JAQ?")
Comp_Analysis <- createSheet(JAQ_Workbook, sheetName = "Comp_Analysis")
addDataFrame(x = Percent, sheet = Comp_Analysis, startRow=1, startColumn=1,colStyle = dfColIndex,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)
setColumnWidth(sheet = Comp_Analysis,colIndex = 2,colWidth = 30)

# Save XLSX JAQ_Workbook --------------------------------------------------
saveWorkbook(wb = JAQ_Workbook,file = paste0(getwd(),"/","RJAQ/OUTPUT/JAQ_Workbook.xlsx"))

if(LAQ_Presence == TRUE){
# --------da.csv Relative Duty Area Weightings ----------
if(file.exists("~/RJAQ/userData/da.csv") == TRUE) {
DutyAreaRatings        <- read.csv(file="~/RJAQ/userData/da.csv",header=FALSE,sep=",",col.names = DutyAreas_Names_Acro[1:ncol(read.csv(file="~/RJAQ/userData/da.csv"))])
Corrected_DAR          <- ((colMeans(DutyAreaRatings,na.rm=T) / mean(rowSums(DutyAreaRatings,na.rm=T))) * 100)
Ratio_DAR              <- (Corrected_DAR / 100)}

# -------Remark Relative Duty Area Weightings ----------
if("DA1_FirstDigit" %in% names(task)){
  Count_DAs_Linkage  <- as.numeric(ncol(task[,grepl("_FirstDigit",names(task))])*2)
  REL_DA_Weightings  <- task[,which(colnames(task) == "DA1_FirstDigit"):(which(colnames(task) == "DA1_FirstDigit") + Count_DAs_Linkage -1)]
  REL_DA_Weightings[is.na(REL_DA_Weightings)] <- 0
  if((sum(is.na(REL_DA_Weightings)) %% Count_DAs_Linkage == 0) == TRUE) {
    DutyAreaRatings <- NULL #create a raw DF.
    First_DA        <- which(colnames(task) == "DA1_FirstDigit")
    Last_DA         <- (First_DA) + (Count_DAs_Linkage - 1)

    for (i in seq(1,ncol(REL_DA_Weightings),by=2)){
      DA_col            <-  as.numeric(do.call(paste0, REL_DA_Weightings[,c(i,i+1)]))
      DutyAreaRatings   <- data.frame(cbind(DutyAreaRatings,DA_col))
    }
    colnames(DutyAreaRatings) <- DutyAreas_Names_Acro[1:(Count_DAs_Linkage/2)]
    Corrected_DAR             <- ((colMeans(DutyAreaRatings) / mean(rowSums(DutyAreaRatings))) * 100)
    Ratio_DAR                 <- (Corrected_DAR / 100)
    rm("First_DA","Last_DA","DA_col", "Corrected_DAR")
}}

# SurveyMonkey Relative Duty Area Weightings ------------------------------
if (Survey_Monkey == TRUE && DutyAreas_Count == ncol(task[,grepl("DA",names(task))])) {
  SM_DAR        <- paste0("DA_",1:DutyAreas_Count)
  #task[SM_DAR]  <- replace(task[SM_DAR],is.na(task[SM_DAR]),0) #ISSUE. Lots of NAs get converted to 0's and then the Ratio_DAR doesn't work anymore.
  Corrected_DAR <- (colMeans(task[,SM_DAR],na.rm=T)) #Updated 10.25.2017 because Survey Monkey forces 100 percent so no need to tweak.
  Ratio_DAR     <- (Corrected_DAR / 100)
  if(length(Ratio_DAR) == DutyAreas_Count)
    names(Ratio_DAR) <- DutyAreas_Names_Full
}

# SurveyMonkey Relative Duty Area Weightings but DA MisMatch ------------------------------
# RJAQ v2.0.2 - Added to reflect Survey Monkey surveys where DA count does NOT match with task analysis count.
if (Survey_Monkey == TRUE && DutyAreas_Count != ncol(task[,grepl("DA",names(task))])) {
  DutyAreas_Count <- ncol(task[,grepl("DA",names(task))])
  SM_DAR        <- paste0("DA_",1:DutyAreas_Count)
  Corrected_DAR <- (colMeans(task[,SM_DAR],na.rm=T)) #Updated 10.25.2017 because Survey Monkey forces 100 percent so no need to tweak.
  Ratio_DAR     <- (Corrected_DAR / 100)
}

} #End parantheses from LAQ_Presence check.

if(LAQ_Presence == TRUE){
  LAQ_Workbook <- createWorkbook(type="xlsx")
    # Styles for the data table row/column names
    TABLE_ROWNAMES_STYLE_LAQ <- CellStyle(LAQ_Workbook) + Font(LAQ_Workbook, isBold=TRUE) + Alignment(horizontal= "ALIGN_CENTER") + 
                                Border(color="black", position=c("TOP", "BOTTOM","LEFT","RIGHT"), pen=c("BORDER_THIN"))
    TABLE_COLNAMES_STYLE_LAQ <- CellStyle(LAQ_Workbook) + Fill(foregroundColor = "dodgerblue4")+ Font(LAQ_Workbook, isBold=TRUE,name = "Calibri",color = "azure",heightInPoints = 12) +
                                Alignment(wrapText=FALSE, horizontal="ALIGN_CENTER") + Border(color="lightgrey", position=c("TOP", "BOTTOM","LEFT","RIGHT"), 
                                pen=c("BORDER_THIN", "BORDER_THICK","BORDER_THIN","BORDER_THIN")) 
    ROWS_LAQ                 <- CellStyle(LAQ_Workbook) + Font(wb = LAQ_Workbook,name="Calibri",heightInPoints = 12) + Alignment(horizontal = "ALIGN_CENTER",wrapText = TRUE,vertical = "VERTICAL_CENTER") + 
                                Border(color="black",position=c("TOP","LEFT","RIGHT","BOTTOM"), pen=c("BORDER_THIN"))

# Linkage: Parameter Settings ---------------------------------------------
SAAL_IMP_Count  <- ncol(link[,grepl("SAAL_IMP_",names(link))])
JDKL_IMP_Count  <- ncol(link[,grepl("JDKL_IMP_",names(link))])
SAALs_IMP       <- paste0("SAAL_IMP_",1:SAAL_IMP_Count)
JDKLs_IMP       <- paste0("JDKL_IMP_",1:JDKL_IMP_Count)
    
# Linkage Read in LAQ.XLSX -------------------------------------------------
LAQ_Dims                 <- which(is.na(Statements_LAQ$Number))
if(length(LAQ_Dims) == 2){
LAQ_Length_Dim           <- length(Statements_LAQ$Number)
Linkage_2                <- length(Statements_LAQ$Number) - which(is.na(Statements_LAQ$Number))[2]
Linkage_1                <- which(is.na(Statements_LAQ$Number))[2] - which(is.na(Statements_LAQ$Number))[1] -1 
SAAL_Names               <- Statements_LAQ$Description[2:(which(is.na(Statements_LAQ$Number))[2] - 1)]
JDKL_Names               <- Statements_LAQ$Description[(which(is.na(Statements_LAQ$Number))[2])+1:length(Statements_LAQ$Number)][1:Linkage_2]}
if(length(LAQ_Dims) == 1 & SAAL_IMP_Count > 0){
LAQ_Names                <- Statements_LAQ$Description[2:length(Statements_LAQ$Description)]
SAAL_Names               <- Statements_LAQ$Description[2:length(Statements_LAQ$Description)]
LAQ_Length_Dim           <- (length(Statements_LAQ$Number)-1)
Linkage_1                <- length(LAQ_Names) }
if(length(LAQ_Dims) == 1 & JDKL_IMP_Count > 0){
LAQ_Names                <- Statements_LAQ$Description[2:length(Statements_LAQ$Description)]
JDKL_Names               <- Statements_LAQ$Description[2:length(Statements_LAQ$Description)]
LAQ_Length_Dim           <- (length(Statements_LAQ$Number)-1)
Linkage_2                <- length(LAQ_Names) }
}

if(JAQ_Presence == FALSE & LAQ_Presence == TRUE){
  #Must manually retrieve the duty areas from the file.
  DutyAreas_Names_Acro <- paste0("DA_",1:(SAAL_IMP_Count / length(SAAL_Names)))
}

if(LAQ_Presence == TRUE){
# Linkage: SAAL Computations ----------------------------------------------
if(SAAL_IMP_Count > 0){
  SAAL_ALI    <- colMeans(link[,SAALs_IMP],na.rm=T)
  SAAL_Matrix <- matrix(SAAL_ALI,nrow=Linkage_1)
  colnames(SAAL_Matrix)   <- c(as.character(DutyAreas_Names_Acro[1:length(Ratio_DAR)]))
  rownames(SAAL_Matrix)   <- c(as.character(SAAL_Names))
  SAAL_Weighted_Matrix    <- as.data.frame(sapply(1:ncol(SAAL_Matrix),function(x) Ratio_DAR[x] * SAAL_Matrix[,x]))
  colnames(SAAL_Weighted_Matrix) <- c(as.character(DutyAreas_Names_Acro[1:length(Ratio_DAR)]))
  SAAL_Total_Row        <- rowSums(SAAL_Weighted_Matrix)
  SAAL_Total_Row_Z      <- scale(SAAL_Total_Row)
  SAAL_Total_Row_STD    <- (((SAAL_Total_Row_Z) *1) + 3)
  SAAL_Weighted_Matrix  <- cbind(SAAL_Weighted_Matrix,SAAL_Total_Row, SAAL_Total_Row_Z, SAAL_Total_Row_STD)
  SAAL_Total_Col        <- colSums(SAAL_Weighted_Matrix)
  SAAL_Weighted_Matrix  <- rbind(SAAL_Weighted_Matrix,SAAL_Total_Col)  
  colnames(SAAL_Weighted_Matrix) <- c(as.character(DutyAreas_Names_Acro[1:length(Ratio_DAR)]),"SAAL_Total","SAAL_Total_Z","SAAL_Total_STD")
  rownames(SAAL_Weighted_Matrix) <- c(as.character(SAAL_Names), "Total")
  SAAL_Weighted_Matrix <- SAAL_Weighted_Matrix[with(SAAL_Weighted_Matrix,order(-SAAL_Total_Row_STD)),]
  #XLSX Output Stuff
  SAAL_Raw_Weightings  <- createSheet(LAQ_Workbook, sheetName = "SAAL_Raw_Weightings")
  SAAL_Calc_Weightings <- createSheet(LAQ_Workbook,sheetName = "SAAL_Weighted_Matrix")
  #Row Styles for SAAL_Weighted_Matrix
  dfColIndex_SAAL_W           <- rep(list(ROWS_LAQ), dim(SAAL_Weighted_Matrix)[2]) 
  names(dfColIndex_SAAL_W)    <- seq(1, dim(SAAL_Weighted_Matrix)[2], by = 1)
  #Row Styles for SAAL_Matrix (Raw)
  dfColIndex_SAAL_R           <- rep(list(ROWS_LAQ), dim(SAAL_Matrix)[2]) 
  names(dfColIndex_SAAL_R)    <- seq(1, dim(SAAL_Matrix)[2], by = 1)
  #Add a table
  addDataFrame(x = SAAL_Matrix, sheet = SAAL_Raw_Weightings,colStyle = dfColIndex_SAAL_R, startRow=1, startColumn=1,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE_LAQ, rownamesStyle = TABLE_ROWNAMES_STYLE_LAQ)
  addDataFrame(x = SAAL_Weighted_Matrix, sheet = SAAL_Calc_Weightings,colStyle = dfColIndex_SAAL_W, startRow=1, startColumn=1,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE_LAQ, rownamesStyle = TABLE_ROWNAMES_STYLE_LAQ)
  setColumnWidth(sheet = SAAL_Raw_Weightings,colIndex = 1,colWidth = 40)
  setColumnWidth(sheet = SAAL_Calc_Weightings,colIndex = 1,colWidth = 40)
  setColumnWidth(sheet = SAAL_Calc_Weightings, colIndex = (ncol(SAAL_Matrix)+2):(ncol(SAAL_Weighted_Matrix)+1),colWidth = 16)
}

# Linkage: JDKL Computations ----------------------------------------------
if(JDKL_IMP_Count > 0) {
  JDKL_ALI    <- colMeans(link[,JDKLs_IMP],na.rm=T)
  JDKL_Matrix <- matrix(JDKL_ALI,nrow=Linkage_2)
  colnames(JDKL_Matrix)   <- c(as.character(DutyAreas_Names_Acro[1:length(Ratio_DAR)]))
  rownames(JDKL_Matrix)   <- c(as.character(JDKL_Names[1:Linkage_2]))
  JDKL_Weighted_Matrix    <- as.data.frame(sapply(1:ncol(JDKL_Matrix),function(x) Ratio_DAR[x] * JDKL_Matrix[,x]))
  colnames(JDKL_Weighted_Matrix)   <- c(as.character(DutyAreas_Names_Acro[1:length(Ratio_DAR)]))
  
  #--------(SPSS) Standardization (LINK:Knowledge) -------
  JDKL_Total_Row        <- rowSums(JDKL_Weighted_Matrix)
  JDKL_Total_Row_Z      <- scale(JDKL_Total_Row)
  JDKL_Total_Row_STD    <- (((JDKL_Total_Row_Z) *1) + 3)
  JDKL_Weighted_Matrix  <- cbind(JDKL_Weighted_Matrix,JDKL_Total_Row, JDKL_Total_Row_Z, JDKL_Total_Row_STD)
  JDKL_Total_Col        <- colSums(JDKL_Weighted_Matrix)
  JDKL_Weighted_Matrix  <- rbind(JDKL_Weighted_Matrix,JDKL_Total_Col)  
  colnames(JDKL_Weighted_Matrix) <- c(as.character(DutyAreas_Names_Acro[1:length(Ratio_DAR)]),"JDKL_Total","JDKL_Total_Z","JDKL_Total_STD")
  rownames(JDKL_Weighted_Matrix) <- c(as.character(JDKL_Names[1:Linkage_2]), "Total")
  #XLSX Output Stuff
  JDKL_Raw_Weightings  <- createSheet(LAQ_Workbook, sheetName = "JDKL_Raw_Weightings")
  JDKL_Calc_Weightings <- createSheet(LAQ_Workbook,sheetName = "JDKL_Weighted_Matrix")
  JDKL_Weighted_Matrix <- JDKL_Weighted_Matrix[with(JDKL_Weighted_Matrix,order(-JDKL_Total_Row_STD)),]
  dfColIndex_JDKL_W           <- rep(list(ROWS_LAQ), dim(JDKL_Weighted_Matrix)[2]) 
  names(dfColIndex_JDKL_W)    <- seq(1, dim(JDKL_Weighted_Matrix)[2], by = 1)
  #Row Styles for SAAL_Matrix (Raw)
  dfColIndex_JDKL_R           <- rep(list(ROWS_LAQ), dim(JDKL_Matrix)[2]) 
  names(dfColIndex_JDKL_R)    <- seq(1, dim(JDKL_Matrix)[2], by = 1)
  #Add a table
  addDataFrame(x = JDKL_Matrix, sheet = JDKL_Raw_Weightings,colStyle = dfColIndex_JDKL_R, startRow=1, startColumn=1,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE_LAQ, rownamesStyle = TABLE_ROWNAMES_STYLE_LAQ)
  addDataFrame(x = JDKL_Weighted_Matrix, sheet = JDKL_Calc_Weightings,colStyle = dfColIndex_JDKL_W, startRow=1, startColumn=1,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE_LAQ, rownamesStyle = TABLE_ROWNAMES_STYLE_LAQ)
  setColumnWidth(sheet = JDKL_Raw_Weightings,colIndex = 1,colWidth = 40)
  setColumnWidth(sheet = JDKL_Calc_Weightings,colIndex = 1,colWidth = 40)
  setColumnWidth(sheet = JDKL_Calc_Weightings, colIndex = (ncol(JDKL_Matrix)+2):(ncol(JDKL_Weighted_Matrix)+1),colWidth = 16)
  }
}
if(LAQ_Presence == TRUE){
# #Duty Area Weightings in XLSX Format ------------------------------------
DutyArea_Weightings2  <- createSheet(LAQ_Workbook, sheetName = "DutyArea_Weightings")
#Add a table
addDataFrame(x = Ratio_DAR, sheet = DutyArea_Weightings2, startRow=1, startColumn=1,showNA = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE_LAQ, rownamesStyle = TABLE_ROWNAMES_STYLE_LAQ)
setColumnWidth(sheet = DutyArea_Weightings2,colIndex = 1,colWidth = 50)
#Output XLSX LAQ_Workbook
saveWorkbook(wb = LAQ_Workbook,file = paste0(getwd(),"/","RJAQ/OUTPUT/LAQ_Workbook.xlsx")) #FIXTHIS
}

#Clean up variables -------------------------------------------------------------------------------------
if(exists("LAQ_Workbook")){rm(LAQ_Workbook)}
rm(dat)

# Save Variables Out
save.image("~/RJAQ/OUTPUT/Dynamic_Reports/Data/Variables.RData")

# Variable Saving ---------------------------------------------------------
Time_Stamp  <- gsub(":",".",Sys.time())

# Run Significance Testing Analyses ---------------------------------------------------
if(SignificanceTesting == TRUE & JAQ_Presence == TRUE){source(file = "G:/IOSolutions/Projects/current/PROJECT TEMPLATES/Job Analysis/RJAQ/2_SignificanceTesting.R",echo = TRUE)}

# History Folder Compilation ----------------------------------------------
task_dir_SAV   <- grep(pattern= "Tasks", x=list.files("~/RJAQ/userData", pattern = "\\.sav$",ignore.case = TRUE), value=TRUE,ignore.case=TRUE)

if(exists("task")) {Client_Name <- gsub("Tasks|.sav","",ignore.case = TRUE,x = task_dir_SAV)}
Client_Name <- gsub("[[:space:]]", "", Client_Name)

#Save Variables again after SigTesting.
save.image("~/RJAQ/OUTPUT/Dynamic_Reports/Data/Variables.RData")

#Run Dynamic Reporting
#Won't work to get it into DOCX format.
#knit(input = "G:/IOSolutions/Projects/current/PROJECT TEMPLATES/Job Analysis/RJAQ/3_DynamicReporting.Rmd",output = "~/RJAQ/OUTPUT/Dynamic_Reports/Testy.docx")

dir.create(path = paste0("~/RJAQ/History/", Time_Stamp, Client_Name)) #timestamp the directory so it will never overwrite.
dir.create(path = paste0("~/RJAQ/History/", Time_Stamp, Client_Name,"/","userData")) #timestamp the directory so it will never overwrite.
dir.create(path = paste0("~/RJAQ/History/", Time_Stamp, Client_Name,"/","OUTPUT")) #timestamp the directory so it will never overwrite.
dir.create(path = paste0("~/RJAQ/History/", Time_Stamp, Client_Name,"/","OUTPUT/","Reports")) #timestamp the directory so it will never overwrite.

file.copy(from = list.files("~/RJAQ/userData/",full.names = TRUE) ,to = paste0("~/RJAQ/History/",Time_Stamp, Client_Name,"/","userData"))
file.copy(from = list.files("~/RJAQ/OUTPUT/",full.names = TRUE,pattern = "_Workbook") ,to = paste0("~/RJAQ/History/",Time_Stamp, Client_Name,"/","OUTPUT"))
file.copy(from = "~/RJAQ/OUTPUT/Dynamic_Reports/Data/Variables.RData", to = paste0("~/RJAQ/History/",Time_Stamp, Client_Name,"/","OUTPUT/","Reports") )
