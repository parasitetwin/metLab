####Deconvolution of lipidomics data-script####
#By Anton Ribbenstedt, April 2016#

#'Deconvoluting flow-injected lipid m/z isotopes
#'
#'Takes a dataset with samples and lipids in rows and columns and deconvolutes isotopic overlap using a reference document with chemical formulas for each lipid name.
#'@param nSheet Which sheet number to grab chemical formulas for the lipids from
#'@param fileRef Reference to a ".xlsx"-file with reference chemical formulas and names of compounds (same used in XCalibur method), if NULL you will be prompted to choose a file
#'@param fileArea Reference to a ".xlsx"-file which contains samples with areas to be deconvoluted, if NULL you will be prompted to choose a file
#'@usage First choose the reference file with your lipid fragments (see example doc. "Lipids.xlsx"), then choose the file with the areas from samples and lipids (see example doc. "DeconvExample.xlsx")
#'@return The corrected data-set will be returned from the function as a data frame
#'@import Rdisop MASS openxlsx
#'@depends Rdisop MASS openxlsx
#'@export


deconLipids<-function(nSheet,fileRef=NA,fileArea=NA){
  require(Rdisop)
  require(MASS)
  require(openxlsx)
  
  ## Starts of with the deconvolution of just carnitines, since a different fragment is being monitored than for the other lipids ##
  ## Input list of Col1: Names of analytes to be deconvoluted & Col2: Chemical formula of product ion of interest (Can require some thought to figure out which one to use)
  
  if(is.na(fileRef)==TRUE){
    paste("Choose the reference file containing col1: Analyte name  col2: Chemical formula of the product ions of interest:")
    fileName1<-file.choose()
  } else {
    fileName1<-fileRef
  }
  
  if(is.na(fileArea)==TRUE){
    paste("Choose the file containing samples in rows and compounds in columns with areas: ")
    fileName2<-file.choose()
  } else {
    fileName2<-fileArea
  }
  
  
  linn<-list()
  wholeFile<-readWorkbook(fileName1, sheet=nSheet)
  linn$name<-as.character(wholeFile[,1])
  linn$formula<-as.character(wholeFile[,2])
  
  Isotopes<-list()
  
  
  #Getting exact mass and sorting molecules based on this
  for (i in 1:length(linn$formula)){
    linn$mass[i]<-getMolecule(linn$formula[i])$exactmass
  }
  
  linn<-as.data.frame(linn)
  linn<-linn[order(linn$mass),]
  row.names(linn) <- 1:nrow(linn)
  linn$formula<-as.character(linn$formula)
  
  
  #Reading intensities from "intensities.txt"
  intensities<-readWorkbook(fileName2, sheet=nSheet)
  sampleNames<-as.character(intensities[,1])
  intensities<-t(intensities[,-1])
  intensities<-intensities[match(linn$name,rownames(intensities)),]
  #intensities<-intensities[,-1]
  class(intensities)<-"numeric"
  rownames(intensities)<-1:length(intensities[,1])
  
  #Getting low res mass spec isotopic spectrums
  for (i in 1:length(linn$formula)){
    Isotopes[[i]]<-as.data.frame(getMolecule(linn$formula[i],z=1,maxisotopes=10)$isotopes)
  }
  
  #Writing .csv file with all isotopic distributions
  #lapply(Isotopes, function(x) write.table( data.frame(x), 'isotopedist.csv'  , append= T, sep=',', col.names=TRUE))
  
  #Loop that will correct higher mass lipids for isotopic intensity contribution
  for(i in 1:(length(linn$formula)-1)){
    
    currentIsotopes<-Isotopes[[i]]
    
    #Counter to avoid trying vs already corrected isotopes
    n=0
    k=0
    #Looping through 10 isotopes
    for(j in 1:9){
      
      #Looping through checks vs higher massed compounds for each isotope
      for(k in 1:9){
        
        #If the last compound has already been checked, break the loop to avoid error
        if((i+k+n) > length(linn$mass)){
          break
        }
        
        #If the compound is as big as the isotope, correct intensities and break the loop
        if(round(linn$mass[i+k])==round(currentIsotopes[1,(1+j)])){
          cat("Isotope: ",currentIsotopes[1,1]," ?ndrar: ", linn$mass[i+k] ,", i+k: ", (i+k), ", 1+j: ", (1+j), ", i: ", i, "\n")
          #Loop to correct intensities for all samples
          for(samp in 1:(length(intensities[1,]))){
            
            #Checking if correction would result in a negative number before making it
            if((intensities[(i+k),samp]-((intensities[(i),samp]/currentIsotopes[2,1])*currentIsotopes[2,(1+j)]))>0){
              intensities[(i+k),samp]<-intensities[(i+k),samp]-((intensities[(i),samp]/currentIsotopes[2,1])*currentIsotopes[2,(1+j)])
            }
            
            #If statement which defines "detection"-limit and sets all NA:s to 0
            #if(intensities[(i+k+n),samp] < 10000 | is.na(intensities[(i+k+n),samp])==TRUE){
            #  intensities[(i+k+n),samp]<-0
            #}
            
          }
          
          n<-n+1
        }
      }
    }
  }
  
  #Writing corrected intensities to excel-file
  intensities<-cbind(as.character(linn$name),intensities)
  colnames(intensities)<-c("",sampleNames)
  
  
  #write.xlsx(intensities, addWorksheet(fileOutput, sheetName=paste(getSheetNames(fileName1)[nSheet])))
  
  return(intensities)
}
