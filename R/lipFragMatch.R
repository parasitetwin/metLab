##Fragment to exact mass matching##
##By Anton Ribbenstedt##

#'Matches a list of points where a mz2-fragment was measured to have a certain intensity above zero to a list of masses and putative identities for them
#'
#'
#'@param fileOutputName The output name of the file you want to create
#'@usage Put in a file name you want on your output file. The script will prompt you to open a document with the same layout as "ExampleFileLipFrag.xlsx". The script will match mz1 m/z & RT against >0 intensity mz2 hits of your chosen fragment. It will then match it against m/z of a list of mass-IDs and produce a full m/z, RT and massID-list as output.
#'@return A list with as many items as lipid types will be returned by the function
#'@export

lipFragMatch<-function(fileOutputName="LipFragMatchResults.xlsx"){
  ####Loading libraries, setting up directories and file names####
  options(java.parameters = "-Xmx8000m")
  options(digits=8)

  library(openxlsx)
  library(installr)
  library(compare)
  #library(dtplyr)
  #library(plyr)
  #library(dplyr)

  #function for counting duplicates
  library(data.table)

  ##The whole procedure for PCs only##
  print("Choose the input file, structured according to \"ExampleFileLipFragMatch.xlsx\":")
  dev.flush()
  fileName<-file.choose()
  numberOfFrags<-length(getSheetNames(fileName))

  for(nLips in 3:(numberOfFrags)){
    cat("nLips: ", nLips, "\n")
    dev.flush()
    AllPCfrags<-NULL
    AllPCfrags<-read.xlsx(fileName, colNames = FALSE, sheet=nLips)
    AllPCfrags[is.na(AllPCfrags)]<-1
    PCfrags<-list()

    Nsamples<-NULL
    Nsamples<-table(grepl("End Trace Points",AllPCfrags[,1]))
    sampleList<-NULL
    sampleList<-grep("End Trace Points", AllPCfrags[,1])
    AllPCfrags<-AllPCfrags[!AllPCfrags[,2]==0,]
    nameList<-NULL
    nameList<-AllPCfrags[(grep("Start Trace",AllPCfrags[,1])+1),]

    ##Putting all the RTs where the fragment was found in each sample into a list
    #For the first sample at first
    AllPCfrags<-AllPCfrags[grep("File:",AllPCfrags[,1],invert=TRUE),]
    PCfrags[[1]]<-AllPCfrags[c(1:match(sampleList[1],rownames(AllPCfrags))-1),]  #
    PCfrags[[1]]<-PCfrags[[1]][grep("Start Trace",PCfrags[[1]][,1],invert=TRUE),]
    PCfrags[[1]]<-PCfrags[[1]][grep("End Trace",PCfrags[[1]][,1],invert=TRUE),]

    #For the rest of the samples in a loop
    for(i in 1:(Nsamples[2]-1)){
      PCfrags[[i+1]]<-AllPCfrags[c((match(sampleList[i],rownames(AllPCfrags))+2):match(sampleList[i+1],rownames(AllPCfrags))-1),]
      PCfrags[[i+1]]<-PCfrags[[i+1]][grep("Start Trace",PCfrags[[i+1]][,1],invert=TRUE),]
      PCfrags[[i+1]]<-PCfrags[[i+1]][grep("End Trace",PCfrags[[i+1]][,1],invert=TRUE),]
      PCfrags[[i+1]][,1]<-as.double(PCfrags[[i+1]][,1])
      #PCfrags[[i+1]]<-PCfrags[[i+1]][PCfrags[[i+1]][,2]>10000,] ##FILTER HERE, WAS CAUSING TOO BIG REMOVALS IN CARNITINES, IF-STATEMENT FOR THE DIFFERENT KIND OF ANALYTES
    }
    PCfrags[[1]][,1]<-as.numeric(PCfrags[[1]][,1])


    #cat("Bug-Check 1\n")
    #dev.flush()

    ####Cleaning up memory and starting matching against masses####
    AllPCfrags<-NULL
    colindex<-1:2
    mzRT<-read.xlsx(fileName,sheet=1)
    mzRT<-mzRT[,c(1:2)]#RT and mz from file
    massID<-read.xlsx(fileName,sheet=2)
    massID<-massID[,c(1:2)]#Only reading m/z and name now, plans to extend to write out annotations later?

    #Connecting every RT with one or more m/z at that same RT, in every sample separately
    nonNACols<-0
    for(l in 1:length(PCfrags)){

      ##Sorting out empty MeOH-files and thus avoiding errors
      if(sum(length(which(round(as.numeric(mzRT[,2]),digits=2)%in%round(as.numeric(PCfrags[[l]][,1]),digits=2))))==0){
        PCfrags[[l]]<-numeric(0)
        next()
      }


      #PCfrags[[l]] <-cbind(PCfrags[[l]][round(as.numeric(mzRT[,2]),digits=2)%in%round(as.numeric(PCfrags[[l]][,1]),digits=2),],NA)
      for(q in 1:100){
        PCfrags[[l]] <-cbind(PCfrags[[l]],NA)  ##HERE IS INTRODUCTION OF ALL COLUMNS, NOT SURE WHY?! <------------------------------------- ???
      }
      PCfrags[[l]]<-PCfrags[[l]][!is.na(PCfrags[[l]][,1]),]
      for(i in 1:length(PCfrags[[l]][,1])){
        if(round(as.numeric(PCfrags[[l]][i,1]),digits=2)%in%round(as.numeric(mzRT[,2]),digits=2)==TRUE){
          PCfrags[[l]][i,3:(length(mzRT[round(as.numeric(mzRT[,2]),digits=2)%in%round(as.numeric(PCfrags[[l]][i,1]),digits=2),1])+2)]<-mzRT[round(as.numeric(mzRT[,2]),digits=2)%in%round(as.numeric(PCfrags[[l]][i,1]),digits=2),1]
        }
      }
      ##Loop to find out which is the first column to be completely NA
      for(i in 1:length(PCfrags[[l]][1,])){
        if(sum(is.na(PCfrags[[l]][,i]))==length(PCfrags[[l]][,i])){
          if(i > nonNACols){
            nonNACols<-i
          }
          break()
        }
      }

      PCfrags[[l]]<-PCfrags[[l]][!is.na(PCfrags[[l]][,3]),]
    }

    #cat("Bug-Check 2\n")
    #dev.flush()

    PCfrags<-PCfrags[lapply(PCfrags,length)>0]
    nonNACols<-nonNACols+1 #Just to easify the processing of data lower in the script

    ##Creating a list of all samples which contains lists of all mass columns for each sample
    allSamps<-list()
    for(i in 1:length(PCfrags)){
      allSamps[[i]]<-list()
      for(n in 1:(nonNACols-3)){
        allSamps[[i]][[n]]<-PCfrags[[i]][,c(1,2,(2+n))]
      }
    }

    #cat("Bug-Check 3\n")
    #dev.flush()

    ###1### Adding names of lipids to masses
    longestCol<-0
    for(l in 1:length(allSamps)){ #Sample
      for(n in 1:length(allSamps[[l]])){ #List for every mass
        for(q in 1:length(allSamps[[l]][[n]][,1])){ #Rows of masses
          if((round(as.numeric(allSamps[[l]][[n]][q,3]),digits=2)%in%round(as.numeric(massID[,1]),digits=2))==TRUE){ #Check if mass on row q matches any masses in mass-ID-list
            k<-which(round(massID[,1], digits=2)==round(allSamps[[l]][[n]][q,3],digits=2))[1]
            for(i in 4:(length(which(round(massID[,1], digits=2)==round(allSamps[[l]][[n]][q,3],digits=2)))+3)){
              allSamps[[l]][[n]][q,i]<-massID[k,2]
              #PCfrags[[l]][q,i]<-massID[k,2]
              k<-k+1
            }
          }
        }
        if(length(allSamps[[l]][[n]][1,])>longestCol){longestCol<-length(allSamps[[l]][[n]][1,])}
      }
    }

    ##Adding all the sublists of masses together into one big list / sample
    fullList<-list()
    for(i in 1:length(allSamps)){
      fullList[[i]]<-data.frame()

      for(n in 1:length(allSamps[[i]])){
        if(n==1){
          fullList[[i]]<-allSamps[[i]][[n]]
          next()
        }
        fullList[[i]]<-rbind.fill(fullList[[i]],allSamps[[i]][[n]])
      }
      if(length(fullList[[i]][1,])<4){next}
      else{
        fullList[[i]]<-fullList[[i]][!is.na(fullList[[i]][,4]),]  ##Remove rows with NAs in column 3
      }
      fullList[[i]][,1]<-round(fullList[[i]][,1],digits=0)
      fullList[[i]]<-fullList[[i]][!duplicated(round(fullList[[i]][,c(1,3)],digits=2)),] #filtering on mz and RT
    }

    ##Checking how many samples a lipid is present in by first building a huge list
    ##and then counting the number of times a lipid appears
    supremeList<-as.data.frame(fullList[[1]])
    for(i in 2:length(fullList)){
      supremeList<-rbind.fill(supremeList,as.data.frame(fullList[[i]]))
    }

    #cat("Bug-Check 4\n")
    #dev.flush()

    ##Ordering all rows based on compound mass and storing number of duplicates in "percentages"
    ##Percentages here is in how many of the samples had a mz2-hit recorded, ___NOT___ how many files it was within in total. Could make such a percentage as well
    percentages<-double()

    DT<-NULL
    DT <- data.table(round(supremeList[,c(1,3)], digits=2))

    #cat("Test0.0: ", DT[,.N, by = names(DT)]$N, "\n")
    percentages<-(DT[,.N, by = names(DT)]$N/length(fullList))*100
    #cat("Test0", length(percentages), "\n")

    #cat("Bug-Check 4.2\n")
    #dev.flush()

    #cat("Test1: ", length(supremeList[!duplicated(round(supremeList[,c(1,3)],digits=2)),]), "\n")

    ##After counting duplicates, removing duplicates. Also binding column of mz2-hit percentages to the list of masses and names
    supremeList = supremeList[!duplicated(round(supremeList[,c(1,3)],digits=2)),]
    #cat("Test2: ", length(supremeList[,1]) , "\n")
    #cat("Test3: ", length(percentages) , "\n")
    supremeList = cbind(as.matrix(percentages),supremeList)

    #cat("Bug-Check 4.3\n")
    #dev.flush()


    ##Remove all the lipids which were never matched against the list
    supremeList<-supremeList[!is.na(supremeList[,5]),]
    supremeList<-supremeList[order(supremeList[,4]),]

    #cat("Bug-Check 4.4\n")
    #dev.flush()

    #####
    ##Area for checking presence in samples based on mz1 intensity and perhaps getting RSD-values while at it
    #####

    mzRT<-NULL
    mzRT<-read.xlsx(fileName,sheet=1) #Readin the document with areas from all sample
    #Just keeping the lipids which were matched against mz2

    #Transforming RT and mz in mzRT in order to make a merge-DF with all areas preserved
    mzRT[,2]<-round(mzRT[,2],digits=0)
    mzRT[,1]<-round(mzRT[,1],digits=2)
    mzRT<-mzRT[order(mzRT[,1]),]
    mzRT<-mzRT[!duplicated(round(mzRT[,c(1,2)],digits=2)),] #Removing duplicates of RT&mz combined

    #cat("Bug-Check 5\n")
    #dev.flush()


    #Merging based on RT and mz, sorting out RT&mzs from mzRT which are not in supremeList
    matching<-merge(mzRT, round(supremeList[,c(4,2)],digits=2), by.x=colnames(mzRT[,c(1,2)]), by.y=colnames(supremeList[,c(4,2)]))


    #Removing lines from supremelist which are not matching
    colnames(supremeList)[2:4]<-c("RT","Int","mz")
    colnames(matching)[1:2]<-c("mz","RT")           ##Formating names of cols to be able to use anti_join properly
    excludedMasses<-anti_join(round(supremeList[,1:4],digits=2),matching,by=c("RT","mz"))
    supremeList[,4]<-round(supremeList[,4],digits=2)
    supremeListFiltered<-inner_join(supremeList,matching)


    #Create stats for the mzRT data frame
    mzRTStats<-data.frame()
    for(i in 1:length(matching[,1])){
      mzRTStats[i,1] <- round(sd(matching[i,-c(1,2)])/mean(as.numeric(unlist(matching[i,-c(1,2)])))*100, digits=1)
    }

    supremeListFiltered<-cbind(mzRTStats[,1],supremeListFiltered)
    colnames(supremeListFiltered)<-c("CV(%),mz1-area","Mz2 hits in files (%)","RT","Int mz2","m/z 2 dec rounded","Lipid name 1","Etc.")

    #####
    ##Automatically calculating PC-length according to our terminology
    #####

    #cat("Bug-Check 6\n")
    #dev.flush()


    ######
    ###Creating excel-workbook, add worksheets for every sample
    ######
    if((nLips-2)==1){
      wb<-createWorkbook()
    }

    fragName<-getSheetNames(fileName)[nLips]
    addWorksheet(wb,paste(fragName))
    writeData(wb,sheet=paste(fragName),supremeListFiltered,colNames=TRUE,rowNames=FALSE)
    #writeData(wb,sheet="Test",matching)
  }

  if(substring(fileOutputName,nchar(fileOutputName)-4,nchar(fileOutputName))==".xlsx"){
    saveWorkbook(wb,fileOutputName)
  } else {saveWorkbook(wb,paste(fileOutputName,".xlsx",sep=""))}
}
