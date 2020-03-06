
Sys.setenv(JAVA_HOME="C:\\Program Files\\Java\\openjdk8u232\\jre")

#install packages
if (!require("pacman")) install.packages("pacman")
pacman::p_load(xlsx, xlsx, RJDBC, splitstackshape, readxl, lubridate, dplyr, svDialogs, readr, stringr, tcltk, anytime)

#clear the environment pane
rm(list=ls())

My_Username <- Sys.info()["user"]

PWLocation <- paste0("//iso-ne.com//shares//",My_Username,"//passwords.csv")

passwords <- read.csv(PWLocation)

#Get the correct password for the database
My_Password <- as.character(passwords[4,2])


#Set the time
currDate <- Sys.time()

#set the file location for the output file
outputlocation <- "\\\\iso-ne.com\\shares\\performance_auditing\\Performance Monitoring\\Audit Resource Parameters\\Claim 10 30\\FTA Results\\Blend Results\\"

filelocation <- "\\\\rtsmb\\Claim1030AuditingData_rw\\C1030_Contingency_App\\BLTS\\"

#Make a list of the directory contents
list <- file.info(list.files(filelocation, pattern = "FTA",full.names=TRUE),drop=FALSE)
list <- tibble::rownames_to_column(list, "Path")

#Set names for the list
names(list) <- c("Path", "size", "isdir", "mode", "mtime", "ctime", "atime", "exe")

#sort file modified time to newest to oldest
list <- arrange(list,desc(mtime))

dateFormat <- "%m/%d/%Y %I:%M:%OS %p"
#Get the filename for use in the import
latestfile <- list[1,"Path"]

#latestfile <- "\\\\rtsmb\\Claim1030AuditingData_rw\\C1030_Contingency_App\\BLTS\\BLTS_FTA_06_19_2019.csv"

TMOR <- read_csv(latestfile)
TMORFull <- TMOR
TMORDate <- as.Date(TMORFull$TIMESTAMP_DDP_RECEIVED, format = "%m/%d/%Y")
TMORDate <- as.character(TMORDate, format = "%m/%d/%Y")

eventDate <- TMORDate

TMORTime <- min(as.POSIXct(TMORFull$TIMESTAMP_DDP_RECEIVED, format = "%m/%d/%Y %I:%M:%OS %p"))
TMORTime <-hour(TMORTime)+1
TMORTime <- sprintf("%02d",TMORTime)

 drv <- JDBC("oracle.jdbc.OracleDriver",
           classPath="\\\\iso-ne.com\\shares\\performance_auditing\\Performance Monitoring\\Audit Resource Parameters\\Claim 10 30\\FTA Results\\ojdbc6.jar"," ")

# Set up the connection to the database 

# Edit this next string in " " as required for a different database.
# Find any database tnsname entry: thin:@(XXXXX  use "tnsping yourdatabasename" in cmd.exe window.


CAMSCon <- dbConnect(drv, "jdbc:oracle:thin:@(description=(address=(host=smsmisp.iso-ne.com)(protocol=tcp)(port=1521))(connect_data=(service_name=smsmisp.world)))",
                     My_Username, My_Password)

# Use the connection to send the query

#    Example: This SQL gets passed to DB by the function dbGetQuery,
#    and puts the results in a data frame called 'test table'

# Make there are empty lines above and below the code

official <- dbGetQuery(CAMSCon,
        
              paste0("
                       
                       select gmt.local_day,
       sms_utilities.get_local_end_hour_from_gmt(t.begin_date) as HE,
       t.asset_id,
       o.customer_id,
       t.location_id,
       t.product_type,
       t.assigned_mw,
       m.asset_type 

from SMS_OWNER.LFR_ASSET_ASSIGNMENT_T t,
     sms_owner.asset_t m,
     sms_owner.ownership_share_t o,
     gmt_local_map_t gmt

where t.begin_date = gmt.gmt_begin_date
and m.asset_type = 7
and t.product_type = 'TMNSR'
and   gmt.local_day = to_date('",eventDate,"', 'mm/dd/yyyy') 
and   gmt.local_hour_end = ('",TMORTime,"')
and   t.asset_id = m.asset_id
and   t.asset_id = o.asset_id
and   o.end_date is NULL

order by 1,2,3,4
                       
                       ")
              )
#Disconnect, and purge info from environment pane
dbDisconnect(CAMSCon)
rm(passwords)
rm(My_Password)


final <- merge(official,TMORFull, by.x = "ASSET_ID", by.y = "RESOURCE_ID")

#____________________________________________________________Final Ouput Section 


rm(My_Password)

#generate report only if the finaloutput has generated a list

xlFileName <- paste0(outputlocation, "FTA DRR TMNSR Blend Results - ", format(currDate, "%Y-%m-%d %H-%M-%S"),".xlsx")

rowcheck <- nrow(final)
if (rowcheck > 0) {
  write.xlsx(x=final, file=xlFileName, sheetName = "DRR_TMNSR_Blend_List", col.names=TRUE, row.names = TRUE)
  msgBox <- tkmessageBox(title = "All done!",
                         message = "A file was created and saved to iso-ne.com//shares//performance_auditing//Performance Monitoring//Audit Resource Parameters//Claim 10 30//FTA Results//Blend Results", icon = "info", type = "ok")
} else {
  msgBox <- tkmessageBox(title = "All done!",
                         message = "There were no DRRs that need to be evaluated for FTA. No report was generated.", icon = "info", type = "ok")
}
# If script does not seem to finish, look for the pop-up box - it might be under another window!
#__________________________________________________________________________END

