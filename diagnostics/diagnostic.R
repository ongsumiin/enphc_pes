library(ProjectTemplate)
load.project()

workbook.name <- list.files("data/", pattern = "[xlsx]")
filename <- list.files("data/", pattern = "[xlsx]", full.names = TRUE)
sheets <- excel_sheets(filename)

for (sheet.name in sheets) {
    
    tryCatch(assign(sheet.name,
                    readxl::read_excel(filename,
                                       sheet = sheet.name),
                    envir = ProjectTemplate:::.TargetEnv),
             error = function(e)
             {
                 warning(paste("The worksheet", sheet.name, "didn't load correctly."))
             })
}


######completion rate######
completed <- capture.output(table(ExitSurvey$DataCompleted), split = TRUE)
completed.prop <- capture.output(round(prop.table(table(ExitSurvey$DataCompleted))*100, 
                                       1), split = TRUE)
cat("Completion Count", completed, "\n", "Completion Rate", 
    completed.prop, file = "diagnostics/PES-Completion Rate.txt", sep = "\n")
#completion by clinic
capture.output(cat("\n", "ALL", "\n"), addmargins(table(ExitSurvey$Site_Name, 
                                                  ExitSurvey$DataCompleted,
                                dnn = c(" ", "Complete Cases"))), 
               file = "diagnostics/PES by clinics.txt", split = TRUE)
#interview completed by clinic
capture.output(cat("\n", "Interview", "\n"), addmargins(table(ExitSurvey$Site_Name, 
                                ExitSurvey$PatientDataCompleted,
                                dnn = c(" ", "Complete Cases"))), 
               file = "diagnostics/PES by clinics.txt", split = TRUE, append = TRUE)

######Data entry by DEP######
#aggregate complete cases by data entry person
creator.y <- aggregate(DataCompleted~CreatedBy, 
                       data = ExitSurvey, 
                       function(x) table(x)["Yes"])
#replace NA with 0
creator.y$DataCompleted[is.na(creator.y$DataCompleted)] <- 0
#calculate total complete cases
creator.y[nrow(creator.y)+1, 1:2] <- c("Total", sum(creator.y$DataCompleted))

#aggregate incomplete cases by data entry person
creator.n <- aggregate(DataCompleted~CreatedBy, 
                       data = ExitSurvey, 
                       function(x) table(x)["No"])
#replace NA with 0
creator.n$DataCompleted[is.na(creator.n$DataCompleted)] <- 0
#calculate total incomplete cases
creator.n[nrow(creator.n)+1, 1:2] <- c("Total", sum(creator.n$DataCompleted))
#rename column
names(creator.n)[2] <- "DataIncomplete"

#merge complete and incomplete data frame by data entry person name
creator <- merge(creator.y, creator.n, by = "CreatedBy", sort = FALSE)
#change class to numeric
creator[2:3] <- apply(creator[2:3], 2, function(x) as.numeric(x))
#calculate total cases by each person
creator$Total <- creator$DataCompleted + creator$DataIncomplete
#sort by total ascending order
creator <- creator[order(creator$Total), ]
#export results to excel in diagnostics folder
w.excel(creator, file = "diagnostics/PES-Data Entry Performance.xlsx")


#######subset complete cases only######
PES <- ExitSurvey[ExitSurvey$DataCompleted == "Yes", ]

#####Age check#####
#calculate patient age
PES$PtAge <- 117 - PES$PtBirthYr
#select cases with age <30
PES.age <- subset(PES, PtAge < 30)
#export if there are cases with age <30
if(nrow(PES.age) > 0) {
    w.excel(data.frame(PES.age), 
            file = "diagnostics/PES-Age less than 30.xlsx")
}else{
    NULL
}


#####Interview Duration#####
#calculate duration to finish interview
PES$duration <- PES$PatientDateCompleted - PES$CreatedDate
#duration that is 3 std score below mean
PES$duration.low <- scale(PES$duration) < -3

#export short duration interview if there are cases
PES.dura <- subset(PES, duration.low == TRUE)
if(nrow(PES.dura) > 0) {
    w.excel(data.frame(PES.dura), 
            file = "diagnostics/PES-Duration outliers.xlsx")
}else{
    NULL
}

#export interview less than 10 min if there are cases
PES.dura1 <- subset(PES, duration < 10)
if(nrow(PES.dura1) > 0) {
    w.excel(data.frame(PES.dura1), 
            file = "diagnostics/PES-Duration Under 10 min.xlsx")
}else{
    NULL
}

#tabulate interview with short duration by data entry personnel 
capture.output(cat("\n", "Short Duration Outliers", "\n"),
               addmargins(table(PES$CreatedBy, PES$duration.low,
                                dnn = c("Data Entry Personnel", 
                                        "Short duration"))),
               file = "diagnostics/PES-duration.txt")
#tabulate interview with duration <10 minits by data entry personnel
capture.output(cat("\n", "Less than 10 minutes", "\n"),
               addmargins(table(PES$CreatedBy, PES$duration < 10,
                                dnn = c("Data Entry Personnel", 
                                        "< 10 minutes"))),
               file = "diagnostics/PES-duration.txt", append = TRUE)





