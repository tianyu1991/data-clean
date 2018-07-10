library("openxlsx")
mydf <- read.xlsx("Copy of Master All Activity Report 2015-2017 use this May 2017.xlsx", sheet = 1, startRow = 1, colNames = TRUE)

##test for full name
FullName<-paste(mydf$First.Name,mydf$Last.Name,sep=" ")
mydf2<-cbind(mydf,FullName)
FullName_Count<-rep(0,nrow(mydf2))
mydf3<-cbind(mydf2,FullName_Count)

newdata<-mydf3[order(FullName),]
for(i in 2:nrow(newdata)){
	if (newdata$FullName[i-1]== newdata$FullName[i]){
		newdata$FullName_Count[i-1]=newdata$FullName_Count[i-1]+1
		newdata$FullName_Count[i]=newdata$FullName_Count[i]+1
		}
	}

FirstDOB<-paste(newdata$First.Name,newdata$DOB,sep=" ")
newdata<-cbind(newdata,FirstDOB)
newdata<-newdata[order(FirstDOB),]


FirstDOB_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstDOB_Count)

for(i in 2:nrow(newdata)){
	if (newdata$FirstDOB[i-1]== newdata$FirstDOB[i]){
		newdata$FirstDOB_Count[i-1]=newdata$FirstDOB_Count[i-1]+1
		newdata$FirstDOB_Count[i]=newdata$FirstDOB_Count[i]+1
		}
	}


FirstAdd<-paste(newdata$First.Name,newdata$standardized_address1,newdata$Address.2,sep=" ")
newdata<-cbind(newdata,FirstAdd)
newdata<-newdata[order(FirstAdd),]

FirstAdd_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstAdd_Count)

for(i in 2:nrow(newdata)){
	if (newdata$FirstAdd[i-1]== newdata$FirstAdd[i]){
		newdata$FirstAdd_Count[i-1]=newdata$FirstAdd_Count[i-1]+1
		newdata$FirstAdd_Count[i]=newdata$FirstAdd_Count[i]+1
		}
	}

LastDOB<-paste(newdata$Last.Name,newdata$DOB,sep=" ")
newdata<-cbind(newdata,LastDOB)
newdata<-newdata[order(LastDOB),]

LastDOB_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,LastDOB_Count)

for(i in 2:nrow(newdata)){
	if (newdata$LastDOB[i-1]== newdata$LastDOB[i]){
		newdata$LastDOB_Count[i-1]=newdata$LastDOB_Count[i-1]+1
		newdata$LastDOB_Count[i]=newdata$LastDOB_Count[i]+1
		}
	}

FirstAddZip<-paste(newdata$First.Name,newdata$standardized_address1,newdata$Zip,sep=" ")
newdata<-cbind(newdata,FirstAddZip)
newdata<-newdata[order(FirstAddZip),]

FirstAddZip_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstAddZip_Count)

for(i in 2:nrow(newdata)){
	if (newdata$FirstAddZip[i-1]== newdata$FirstAddZip[i]){
		newdata$FirstAddZip_Count[i-1]=newdata$FirstAddZip_Count[i-1]+1
		newdata$FirstAddZip_Count[i]=newdata$FirstAddZip_Count[i]+1
		}
	}

FirstAddZip2<-paste(newdata$First.Name,newdata$Address.2,newdata$Zip,sep=" ")
newdata<-cbind(newdata,FirstAddZip2)
newdata<-newdata[order(FirstAddZip2),]

FirstAddZip2_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstAddZip2_Count)

for(i in 2:nrow(newdata)){
	if (newdata$FirstAddZip2[i-1]== newdata$FirstAddZip2[i]){
		newdata$FirstAddZip2_Count[i-1]=newdata$FirstAddZip2_Count[i-1]+1
		newdata$FirstAddZip2_Count[i]=newdata$FirstAddZip2_Count[i]+1
		}
	}

DOBZip<-paste(newdata$DOB,newdata$Zip,sep=" ")
newdata<-cbind(newdata,DOBZip)
newdata<-newdata[order(DOBZip),]

DOBZip_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,DOBZip_Count)

for(i in 2:nrow(newdata)){
	if (newdata$DOBZip[i-1]== newdata$DOBZip[i]){
		newdata$DOBZip_Count[i-1]=newdata$DOBZip_Count[i-1]+1
		newdata$DOBZip_Count[i]=newdata$DOBZip_Count[i]+1
		}
	}

DOBAdd<-paste(newdata$DOB,newdata$standardized_address1,sep=" ")
newdata<-cbind(newdata,DOBAdd)
newdata<-newdata[order(DOBAdd),]


DOBAdd_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,DOBAdd_Count)

for(i in 2:nrow(newdata)){
	if (newdata$DOBAdd[i-1]== newdata$DOBAdd[i]){
		newdata$DOBAdd_Count[i-1]=newdata$DOBAdd_Count[i-1]+1
		newdata$DOBAdd_Count[i]=newdata$DOBAdd_Count[i]+1
		}
	}


EnterID_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,EnterID_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,1]== newdata[i,1]){
		newdata$EnterID_Count[i-1]=newdata$EnterID_Count[i-1]+1
		newdata$EnterID_Count[i]=newdata$EnterID_Count[i]+1
		}
	}

Total_Count<-newdata$FullName_Count+newdata$FirstDOB_Count+newdata$FirstAdd_Count+newdata$LastDOB_Count+newdata$FirstAddZip_Count+newdata$FirstAddZip2_Count+newdata$DOBZip_Count+newdata$DOBAdd_Count+newdata$EnterID_Count
newdata2<-cbind(newdata,Total_Count)

Sys.setenv(R_ZIPCMD= "C:/Rtools/bin/zip")

write.xlsx(x=newdata2, file = "dataclean.xlsx",
        sheetName = "Test", row.names = FALSE)
