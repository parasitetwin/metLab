##Compare masses of lipidomics method and Orbitrap, makes comparisons easier and faster##
##By Anton Ribbenstedt##

#'Takes a lipFragMatch-object and compares the masses to the lipidomics lipid-reference file and puts matching names in a new column in the lipFragMatch-object
#'
#'
#'@param fragMatchObject An object which has been produced through the function "lipFragMatch". If empty function is called you will be asked to choose an ".xlsx" file through explorer
#'@usage Give the function a fragMatchObject or choose a location. Will also be prompted to choose your reference file used for "deconLipids" and "fullDecon". The function will compare the masses rounded up no decimals and put the names of the lipidomics lipids in a new column of the lipFragObject.
#'@return A data frame containing the fragMatchObject with a new first column containing the reference name for the lipidomics method. Will also save and overwrite the excel-document provided, if any.
#'@export

compLipMeth<-function(fragMatchObject=NULL){
  library(openxlsx)
  library(Rdisop)
  library(MASS)

  if(is.null(fragMatchObject)==TRUE){
    print("Choose the fragment matching save file you want to use:")
    flush.console()
    fragMatchFile<-file.choose()
    fragMatchObject<-read.xlsx(fragMatchFile)
  }

  print("Choose the lipidomics reference file:")
  flush.console()
  fileName<-file.choose()                #Get doc name to open it


  nLips<-length(getSheetNames(fileName)) #Number of sheets in doc
  fullFrame<-data.frame()

  ##Adding all the lipids present in the lipid reference document together in one big 2-col data frame
  for(i in 1:nLips){
    tempSheet<-read.xlsx(fileName,sheet=i,colNames=FALSE)
    fullFrame<-rbind(fullFrame,tempSheet)
  }

  ##Adding an empty column in reference data frame to fill up with exact masses
  fullFrame<-cbind(fullFrame,NA)

  ##Getting exact masses for all molecules and stuffing them into third column
  for (i in 1:length(fullFrame[,1])){
    if(is.na(as.numeric(substr(fullFrame[i,3],regexpr("H",fullFrame[i,3])+1,regexpr("H",fullFrame[i,3])+2))-1)==TRUE){
      realHnumb<-as.numeric(substr(fullFrame[i,3],regexpr("H",fullFrame[i,3])+1,regexpr("H",fullFrame[i,3])+1))-1
    }else{
      realHnumb<-as.numeric(substr(fullFrame[i,3],regexpr("H",fullFrame[i,3])+1,regexpr("H",fullFrame[i,3])+2))-1
    }
    fullFrame[i,3]<-sub(as.character(realHnumb+1),as.character(realHnumb),fullFrame[i,3])
    fullFrame[i,4]<-getMolecule(fullFrame[i,3])$exactmass
  }

  ##Comparing the fragment matched masses against the masses from the lipidomics reference document
  ##Adding a column in the lipid fragment match object with

  ###1### Adding names of lipids to masses in first column of RT-Mass match, will have to repeat for every extra match and add to worksheets
  #longestCol<-0
  fragMatchObject<-cbind(NA,fragMatchObject)

  for(n in 1:length(fragMatchObject[,1])){ #List for every mass
    if((round(as.numeric(fragMatchObject[n,6]),digits=2)%in%round(as.numeric(fullFrame[,4]),digits=2))==TRUE){ #Check if mass on row n matches any masses in lipomics object column
      k<-which(round(fullFrame[,4], digits=0)==round(fragMatchObject[n,6],digits=0))[1]
      fragMatchObject[n,1]<-fullFrame[k,1]
    }
  }

  #colnames(fragMatchObject)[1]<-c("Lipidomics m/z match")


  ##Creating a workbook and saving the fragMatchObject with the new column with matching matches for lipids
  wb<-createWorkbook()
  addWorksheet(wb, "LipidMatchResults")
  writeData(wb,sheet=1,fragMatchObject,colNames=TRUE,rowNames=FALSE)
  saveWorkbook(wb,paste(fragMatchFile), overwrite=TRUE)

  return(fragMatchObject)
}
