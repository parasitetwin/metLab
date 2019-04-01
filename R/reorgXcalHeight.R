#Importing .csv output-files from Xcalibur quantifications and deconvoluting Lipids for lipidomics method#
#Anton Ribbenstedt, May 2016#
#' Reorganizing XCalibur-files
#'
#' Takes an output file from XCalibur (resaved as .xlsx), grabs the peak heights and reorganizes it into a one worksheet column and row ".xlsx" file
#' @param saveName String with name of the reorganized output-file saved to your homefolder ending with ".xlsx"; default = "ReorgResults.xslx"
#' @param saveFile Boolean parameter, deciding whether you want to save a file to your homefolder with the reorganized data or not; default = FALSE
#' @param shortReport Boolean parameter, governing whether the input file is in long or short report format
#' @usage Function will ask for an input file which will be organized
#' @return The data reorganized into columns and rows will be returned as a data frame (and saved to homefolder if chosen)
#' @export

reorgXcalHeight<-function(saveFile=TRUE, saveName="ReorgResults.xlsx", shortReport=TRUE){
  ####Loading packages needed for import and analysis of data####
  if(is.na(saveName)==TRUE){saveName="ReorgResults.xlsx"}
  if(is.na(saveFile)==TRUE){saveFile=FALSE}
  require(openxlsx)
  #require(dplyr)
  #require(reshape)
  #options(java.parameters = "-Xmx4g" )
  fileName<-file.choose()

  ####Setting up starting parameters for short and long reports####
  if(shortReport==TRUE){
    startNum<-4
    cols<-c(5)
  } else {
    startNum <- 3
    cols<-c(15)
  }

  ####xlsx and xlconnect operations before requiring openxlsx####
  wbList<-list()
  wholeFile<-read.xlsx(fileName, 1)


  ##Getting analyte names from document##
  sheetNames<- getSheetNames(fileName)
  n_comp<- length(sheetNames)-2


  #Loop counting number of samples and making a list of sample names
  nSamples<-0
  sampleNames<-list()

  for(i in startNum:1000){                     #####3:1000 if LongReport
    if(wholeFile[i,1]=="Created By:")
      break

    nSamples<-nSamples+1
    sampleNames[nSamples]<-wholeFile[i,1]
  }

  expDF<-data.frame(matrix(NA,nrow=(length(sampleNames)+1),ncol=(n_comp+1)))

  ####Importing any specified column from all compounds in the Xcalibur export document####
  ## To change what data you want to extract simply look in the document in which column the data you want is present, and put it into the # in "cols=c(#)" below ##

  xlsxFile<-system.file(fileName, package = "openxlsx")
  sampleLength<-length(sampleNames)+6

  for(i in 1:n_comp){
    wbList[[i]]<-read.xlsx(fileName, i, rows=c(6:sampleLength), cols=cols, colNames=FALSE, rowNames=FALSE) ###cols=c(15) if LongReport (This is for Area)
  }

  names(wbList[[i]])<-c("Area")

  ####Restructuring data into rows of samples and columns of compounds####
  for(i in 2:(length(sampleNames)+1)){
    expDF[i,1]<-sampleNames[i-1]
  }
  expDF[1,-1]<-sheetNames[1:n_comp]

  #Loop for areas and export
  for(i in 1:(n_comp)){
    expDF[-1,i+1]<-wbList[[i]]
  }

  ## Replacing peaks Not found (NF) with 0 for future mathematical applications ##
  expDF[expDF=="NF"]<-0

  ##Exporting data to the filename specified and the sheet-name specified##
  if(saveFile==TRUE){
    if(substring(saveName,nchar(saveName)-4,nchar(saveName))==".xlsx" || substring(saveName,nchar(saveName)-5,nchar(saveName))==".XLSX"){
      write.xlsx(expDF,saveName,col.names=FALSE)
    }else{write.xlsx(expDF,paste(saveName,".xlsx",sep=""),col.names=FALSE)}
  }
  return(expDF)
}
