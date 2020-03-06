
Sys.setenv(JAVA_HOME="C:\\Program Files\\Java\\openjdk8u232\\jre")


#install packages
if (!require("pacman")) install.packages("pacman")
pacman::p_load(readr,xlsx, openxlsx, xlsx, RJDBC, splitstackshape, readxl, lubridate, dplyr, svDialogs, readr, tcltk)


#clear the environment pane
rm(list=ls())

#Set the time
currDate <- Sys.time()
currDate1 <- Sys.Date()

#set the file location for the output file
outputlocation <- "\\\\iso-ne.com\\shares\\performance_auditing\\Performance Monitoring\\Audit Resource Parameters\\Claim 10 30\\FTA Results\\Blend Results"

#Choose file - commented out due to hardcoded path
# myfile <- file.choose()

#Get usernames and passwords

My_Username <- Sys.info()["user"]

PWLocation <- paste0("//iso-ne.com//shares//",My_Username,"//passwords.csv")

passwords <- read.csv(PWLocation)

My_Password <- as.character(passwords[1,2])

#Get CDSPD Case data, dates, apply formatting
myDateFormat <- "%Y/%m/%d"

cdspd <- read.table("\\\\iso-ne.com\\shares\\performance_auditing\\Performance Monitoring\\Audit Resource Parameters\\Claim 10 30\\FTA Results\\CDSPD Cases\\app_filesisogtwapprovedcases.csv", stringsAsFactors = FALSE, header = TRUE ,sep = ",")


cdspd$local_day <- as.Date(cdspd$local_day, format="%m/%d/%Y")

eventHour <- cdspd[1,2]
eventHour <- formatC(eventHour, width=2, format="d", flag="0")

eventDate <- cdspd[1,1]
eventDate <- as.character(eventDate)

currDate1 <- as.character(currDate1)

#Begin BI Stuff
#Query 1 for BI system. BIPROD for warehouse. BIPROD2 for Direct Access
BI_System_1 <-"BIPROD2"

#Query 1 for BI    
BI_Query_1 <- paste0("SELECT \"Operations\".\"Asset Dimension\".\"Asset ID\" as ASSET_ID, 
\"Operations\".\"Asset Dimension\".\"Asset Name\" as ASSET_NAME, 
\"Operations\".\"Reserve Assigned MW\".\"Reserve Product Type\" as RESERVE_PRODUCT_TYPE, 
\"Operations\".\"Time Dimension\".\"Local Day\" as LOCAL_DAY, 
\"Operations\".\"Time Dimension\".\"Local Hour End\" as LOCAL_HOUR_END, 
\"Operations\".\"Reserve Assigned MW\".\"Assigned MW\" as ASSIGNED_MW
FROM\"Operations\" WHERE ((\"Time Dimension\".\"Local Day\" = date '",eventDate,"') AND (\"Time Dimension\".\"Local Hour End\" = '",eventHour,"') AND (\"Reserve Assigned MW\".\"Reserve Product Type\" = 'TMNSR'))")


# Loading JDBC to connect to Oracle
driver <- JDBC("oracle.bi.jdbc.AnaJdbcDriver","\\\\ISO-NE.COM\\shares\\performance_auditing\\Performance Monitoring\\Audit Resource Parameters\\Claim 10 30\\FTA Results\\bijdbc.jar", identifier.quote="'")

# Connect to BI
Connection_1 <- dbConnect(driver, paste("jdbc:oraclebi://",BI_System_1, ".iso-ne.com:9703/", sep=""), My_Username, My_Password)

#Query Results
Query_Results_1 <- dbGetQuery(Connection_1, BI_Query_1)

# Close connection 1
dbDisconnect(Connection_1)
# # Close connection 2
# dbDisconnect(Connection_2)

final <- merge(Query_Results_1,cdspd, by.x = "ASSET_ID", by.y = "asset_id")

#____________________________________________________________Final Ouput Section 

#Purge User password from environment pane
rm(My_Password)
rm(passwords)


#Example filenames:
  #Current Date filename
    #xlFileName <- paste0(outputlocation, "FTA TMNSR Blend Results - ", format(currDate, "%Y-%m-%d"),".csv")
  #Event Date filename
    #xlFileName <- paste0(outputlocation, "FTA TMNSR Blend Results - ",eventDate,".csv")
    #xlFileName <- paste0(outputlocation, "FTA TMNSR Blend Results - Event Date: ",eventDate,"- review date: ",format(currDate, "%m-%d-%Y"),".csv")

#Create a dynamic file name, lets capture event and revew date in filename
xlFileName <- paste0(outputlocation, "\\TMNSR Blend Results - Event Date ", eventDate, " review date ",currDate1,".csv")

#determine how many rows are in our final dataset
rowcheck <- nrow(final)

#if we have rows, then create a file and alert analyst, if not just alert
if (rowcheck > 0) {
  write.csv(x=final, file=xlFileName, row.names = TRUE)
  msgBox <- tkmessageBox(title = "All done!",
                         message = "A CSV file was created and saved to ////iso-ne.com//shares//performance_auditing//Performance Monitoring//Audit Resource Parameters//Claim 10 30//FTA Results//Blend Results", icon = "info", type = "ok")
} else {
  msgBox <- tkmessageBox(title = "All done!",
                         message = "There were no units that need to be evaluated for FTA. No report was generated.", icon = "info", type = "ok")
}

###  IMPORTANT NOTE  ####

# If script does not seem to finish, look for the pop-up box - it might be under another window!
#__________________________________________________________________________END

