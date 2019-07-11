##Deconv full lipid set and save file##

#'Deconvolutes all lipid types present and saves them as .xslx
#'
#'Takes a .xlsx file outputed from reorgXcalArea, applies deconvLipids on all different lipid types and saves an .xlsx-document.
#'@param fileOutputName The output name of the file you want to create
#'@usage Put in the file names you want and the function will use deconvLipids() to create an .xlsx document with all of them deconvoluted inside
#'@return A list with as many items as lipid types will be returned by the function
#'@depends Rdisop MASS openxlsx
#'@export




fullDecon<-function(fileOutputName){
  require(openxlsx)
  cat("#############################################################\nChoose the reference file with lipids and chemical formulas:\n#############################################################\n\n")
  flush.console()
  refFileName<-file.choose()

  cat("#################################################################\nChoose the file containing areas from samples and lipid species:\n#################################################################\n\n")
  flush.console()
  areaFileName<-file.choose()

  wb<-createWorkbook()
  lipidTypes<-as.character(getSheetNames(refFileName))
  nSheets<-length(lipidTypes)
  deconList<-list()

  for(i in 1:nSheets){
    deconList[[i]]<-deconLipids(i, refFileName, areaFileName)
    addWorksheet(wb,paste(lipidTypes[i]))
    writeData(wb,paste(lipidTypes[i]),deconList[[i]])
  }

  if(substring(fileOutputName,nchar(fileOutputName)-4,nchar(fileOutputName))==".xlsx"){
    saveWorkbook(wb,fileOutputName)
  } else {saveWorkbook(wb,paste(fileOutputName,".xlsx",sep=""))}

  return(deconList[[i]])
}
