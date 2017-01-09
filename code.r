mydf <- read.xlsx("2017-01-08 All Activity Report FY17 Q2.xlsx", sheet = 1, startRow = 1, colNames = TRUE)

##test for full name
FullName<-paste(mydf[,2],mydf[,3],sep=" ")
mydf2<-cbind(mydf,FullName)
FullName_Count<-rep(0,nrow(mydf2))
mydf3<-cbind(mydf2,FullName_Count)

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

FirstDOB_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstDOB_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,54]== newdata[i,54]){
		newdata[i-1,55]=newdata[i-1,55]+1
		newdata[i,55]=newdata[i,55]+1
		}
	}



FirstAdd<-paste(newdata[,2],newdata[,13],newdata[,14],sep=" ")
newdata<-cbind(newdata,FirstAdd)
newdata<-newdata[order(FirstAdd),]

FirstAdd_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstAdd_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,56]== newdata[i,56]){
		newdata[i-1,57]=newdata[i-1,57]+1
		newdata[i,57]=newdata[i,57]+1
		}
	}


LastDOB<-paste(newdata[,3],newdata[,4],sep=" ")
newdata<-cbind(newdata,LastDOB)
newdata<-newdata[order(LastDOB),]

LastDOB_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,LastDOB_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,58]== newdata[i,58]){
		newdata[i-1,59]=newdata[i-1,59]+1
		newdata[i,59]=newdata[i,59]+1
		}
	}

FirstAddZip<-paste(newdata[,2],newdata[,13],newdata[,15],sep=" ")
newdata<-cbind(newdata,FirstAddZip)
newdata<-newdata[order(FirstAddZip),]

FirstAddZip_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstAddZip_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,60]== newdata[i,60]){
		newdata[i-1,61]=newdata[i-1,61]+1
		newdata[i,61]=newdata[i,61]+1
		}
	}

FirstAddZip2<-paste(newdata[,2],newdata[,14],newdata[,15],sep=" ")
newdata<-cbind(newdata,FirstAddZip2)
newdata<-newdata[order(FirstAddZip2),]

FirstAddZip2_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,FirstAddZip2_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,62]== newdata[i,62]){
		newdata[i-1,63]=newdata[i-1,63]+1
		newdata[i,63]=newdata[i,63]+1
		}
	}

DOBZip<-paste(newdata[,4],newdata[,15],sep=" ")
newdata<-cbind(newdata,DOBZip)
newdata<-newdata[order(DOBZip),]

DOBZip_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,DOBZip_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,64]== newdata[i,59]){
		newdata[i-1,65]=newdata[i-1,65]+1
		newdata[i,65]=newdata[i,65]+1
		}
	}

DOBAdd<-paste(newdata[,4],newdata[,13],sep=" ")
newdata<-cbind(newdata,DOBAdd)
newdata<-newdata[order(DOBAdd),]


DOBAdd_Count<-rep(0,nrow(mydf2))
newdata<-cbind(newdata,DOBAdd_Count)

for(i in 2:nrow(newdata)){
	if (newdata[i-1,66]== newdata[i,66]){
		newdata[i-1,67]=newdata[i-1,67]+1
		newdata[i,67]=newdata[i,67]+1
		}
	}

Total_Count<-newdata[,67]+newdata[,65]+newdata[,63]+newdata[,61]+newdata[,59]+newdata[,57]+newdata[,55]+newdata[,53]
newdata2<-cbind(newdata,Total_Count)

Sys.setenv(R_ZIPCMD= "C:/Rtools/bin/zip")

write.xlsx(x=newdata2, file = "dataclean.xlsx",
        sheetName = "Test", row.names = FALSE)
