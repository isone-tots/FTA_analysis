
Sys.setenv(JAVA_HOME="C:\\Program Files\\Java\\openjdk8u232\\jre")

#install packages
if (!require("pacman")) install.packages("pacman")
pacman::p_load( RJDBC, splitstackshape, readxl, lubridate, dplyr, svDialogs, readr, stringr, tcltk, anytime, xlsx)

#clear the environment pane
rm(list=ls())

#Set the time
currDate <- Sys.time()

#set the file location for the output file
outputlocation <- "\\\\iso-ne.com\\shares\\performance_auditing\\Performance Monitoring\\Audit Resource Parameters\\Claim 10 30\\FTA Results\\Blend Results\\"

filelocation <- "\\\\rtsmb\\Claim1030AuditingData_rw\\TMOR\\"

#Make a list of the directory contents
list <- file.info(list.files(filelocation,full.names=TRUE),drop=FALSE)
list <- tibble::rownames_to_column(list, "Path")

#Set names for the list
names(list) <- c("Path", "size", "isdir", "mode", "mtime", "ctime", "atime", "exe")

#sort file modified time to newest to oldest
list <- arrange(list,desc(mtime))

#Get the filename for use in the import
latestfile <- list[1,"Path"]

#Import the data and make a copy dataframe called TMORFull
TMOR <- read_csv(latestfile)
TMORFull <- TMOR

#Get the date we need for the BI query, and format it to character, call it eventDate
TMORDate <- TMOR[1,"TIMESTAMP_RECEIVED"]
TMORDate <- as.POSIXct(TMORDate$TIMESTAMP_RECEIVED, format = "%m/%d/%Y")
TMORDate <- as.character(TMORDate)
eventDate <- TMORDate

#Get the time we need for the BI query, and format it to character, call it TMORTime
TMORTime <- min(as.POSIXct(TMORFull$TIMESTAMP_RECEIVED, format = "%m/%d/%Y %I:%M:%S %p"))
TMORTime <-hour(TMORTime)+1
TMORTime <- sprintf("%02d",TMORTime)


#Query 1 for BI system. BIPROD for warehouse. BIPROD2 for Direct Access
BI_System_1 <-"BIPROD2"

#heres the query, using eventDate and TMORTime variables
BI_Query_1 <- paste0("SELECT \"Operations\".\"Asset Dimension\".\"Asset ID\" as ASSET_ID,
\"Operations\".\"Asset Dimension\".\"Asset Name\" as ASSET_NAME,
\"Operations\".\"Reserve Assigned MW\".\"Reserve Product Type\" as RESERVE_PRODUCT_TYPE,
\"Operations\".\"Time Dimension\".\"Local Day\" as LOCAL_DAY,
\"Operations\".\"Time Dimension\".\"Local Hour End\" as LOCAL_HOUR_END,
\"Operations\".\"Reserve Assigned MW\".\"Assigned MW\" as ASSIGNED_MW
FROM\"Operations\" WHERE ((\"Time Dimension\".\"Local Day\" = date '",eventDate,"') AND (\"Time Dimension\".\"Local Hour End\" = '",TMORTime,"') AND (\"Reserve Assigned MW\".\"Reserve Product Type\" = 'TMOR'))")

# Loading JDBC to connect to Oracle
driver <- JDBC("oracle.bi.jdbc.AnaJdbcDriver","c:/RJDBC/bijdbc.jar", identifier.quote="'")

#Get BI username
My_Username <- Sys.info()["user"]
#Get BI password
My_Password <- rstudioapi::askForPassword(" Please enter your Business Intelligence password: ")

# Connect to BI
Connection_1 <- dbConnect(driver, paste("jdbc:oraclebi://",BI_System_1, ".iso-ne.com:9703/", sep=""), My_Username, My_Password)

#Query Results
Query_Results_1 <- dbGetQuery(Connection_1, BI_Query_1)

# Close connection 1
dbDisconnect(Connection_1)
# # Close connection 2
# dbDisconnect(Connection_2)

#merge the two frames and get the final list
final <- merge(Query_Results_1,TMORFull, by.x = "ASSET_ID", by.y = "RESOURCE_ID")

#Purge User password from environment pane
rm(My_Password)

#generate report only if the finaloutput has generated a list

xlFileName <- paste0(outputlocation, "FTA TMOR Blend Results - ", format(currDate, "%Y-%m-%d %H-%M-%S"),".xlsx")

rowcheck <- nrow(final)
if (rowcheck > 0) {
  write.xlsx(x=final, file=xlFileName, sheetName = "TMOR_Blend_List", col.names=TRUE, row.names = TRUE)
  msgBox <- tkmessageBox(title = "All done!",
                         message = "A file was created and saved to iso-ne.com//shares//performance_auditing//Performance Monitoring//Audit Resource Parameters//Claim 10 30//FTA Results//Blend Results", icon = "info", type = "ok")
} else {
  msgBox <- tkmessageBox(title = "All done!",
                         message = "There were no units that need to be evaluared for FTA. No report was generated.", icon = "info", type = "ok")
}
# If script does not seem to finish, look for the pop-up box - it might be under another window!
#__________________________________________________________________________END

