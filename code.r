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
