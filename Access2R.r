# R code for getting Access data into R using RODBC.

#################
## Step 1
#################
# Installing RODBC, Loading RODBC, get RODBC vignette (RODBC documentation)

install.packages("RODBC")
library(RODBC)

vignette("RODBC") # (RODBC documentation)

#################
## Setp 2
#################
# Opening a connection to an Access database

   # Modify the path to the location of the Access database on your computer
WD <- getwd()
db = paste(WD,"/","Test.accdb", sep="")
con2 = odbcConnectAccess2007(db)

   # Gets the names of the tables in the Access database
sqlTables(con2, tableType = "TABLE")$TABLE_NAME


#################
## Setp 3
#################
# importing data

   # Import option 1: Create data frame using the sqlFetch 
testTable = sqlFetch(con2, "Test")
str(testTable)

  # Import option 2: Create data frame using the sqlQuery
qry = "SELECT * FROM Test"
testQueryResult = sqlQuery(con2, qry)
str(testQueryResult)

   # Close connection to Access DB
odbcCloseAll()


