library(xlsx)
library("dplyr")

library("openxlsx")
mydf <- read.xlsx("All Activity Report FY17 Q1 July-September 2016.xlsx", sheet = 1, startRow = 1, colNames = TRUE)

##test for full name
FullName<-paste(mydf[,2],mydf[,3],sep=" ")
mydf2<-cbind(mydf,FullName)
DupCount<-rep(0,nrow(mydf2))
mydf3<-cbind(mydf2,DupCount)

newdata<-mydf3[order(FullName),]
for(i in 2:nrow(newdata)){
	if (newdata[i-1,52]== newdata[i,52]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}

FirstDOB<-paste(newdata[,2],newdata[,4],sep=" ")
newdata<-cbind(newdata,FirstDOB)
newdata<-newdata[order(FirstDOB),]

for(i in 2:nrow(newdata)){
	if (newdata[i-1,54]== newdata[i,54]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}


  

test<-newdata[,13:14]
FirstAdd<-paste(newdata[,2],newdata[,13],newdata[,14],sep=" ")
newdata<-cbind(newdata,FirstAdd)
newdata<-newdata[order(FirstAdd),]

for(i in 2:nrow(newdata)){
	if (newdata[i-1,55]== newdata[i,55]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}

LastDOB<-paste(newdata[,3],newdata[,4],sep=" ")
newdata<-cbind(newdata,LastDOB)
newdata<-newdata[order(LastDOB),]

for(i in 2:nrow(newdata)){
	if (newdata[i-1,56]== newdata[i,56]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}

FirstAddZip<-paste(newdata[,2],newdata[,13],newdata[,15],sep=" ")
newdata<-cbind(newdata,FirstAddZip)
newdata<-newdata[order(FirstAddZip),]

for(i in 2:nrow(newdata)){
	if (newdata[i-1,57]== newdata[i,57]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}

FirstAddZip2<-paste(newdata[,2],newdata[,14],newdata[,15],sep=" ")
newdata<-cbind(newdata,FirstAddZip2)
newdata<-newdata[order(FirstAddZip2),]

for(i in 2:nrow(newdata)){
	if (newdata[i-1,58]== newdata[i,58]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}

DOBZip<-paste(newdata[,4],newdata[,15],sep=" ")
newdata<-cbind(newdata,DOBZip)
newdata<-newdata[order(DOBZip),]

for(i in 2:nrow(newdata)){
	if (newdata[i-1,59]== newdata[i,59]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}

DOBAdd<-paste(newdata[,4],newdata[,13],sep=" ")
newdata<-cbind(newdata,DOBAdd)
newdata<-newdata[order(DOBAdd),]

for(i in 2:nrow(newdata)){
	if (newdata[i-1,60]== newdata[i,60]){
		newdata[i-1,53]=newdata[i-1,53]+1
		newdata[i,53]=newdata[i,53]+1
		}
	}


write.xlsx(x=newdata, file = "dataclean.xlsx",
        sheetName = "Test", row.names = FALSE)

Sys.setenv(R_ZIPCMD= "C:/Rtools/bin/zip")
