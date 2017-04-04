options(java.parameters = "-Xmx1024m")
library(ggplot2)
library(gdata)
library(lattice)
library(gridExtra)
#library(XLConnect)
library(dplyr)
library(openxlsx)

#Path is the path where 'data' folder resides
Path <- "/Users/pranaywankhede/Documents/Statistical/Statistical Assignment"

#Path is the folder wherein the zip files are located
DataPath <- "/Users/pranaywankhede/Documents/Statistical/Statistical Assignment/Data"

#Ouput pdf files are stored on following path
outputPath <- "/Users/pranaywankhede/Documents/Statistical/Statistical Assignment/Output"

RESPath <-"/Users/pranaywankhede/Documents/Statistical/Statistical Assignment/Output/RES"

#create Index File and assign column names
exceldataframe <- data.frame(matrix(ncol = 8, nrow = 0), stringsAsFactors=F)
colnames(exceldataframe) <- c( "Subject", "Session", "thermal data", "pp", "HR", "BR", "peda","performance(res)")
# wb = loadWorkbook("MYDataSet.xlsx", create = TRUE)
# createSheet(wb, name = "DataSheet")

statisticsFrame <- data.frame(matrix(ncol = 6, nrow = 5), stringsAsFactors=F)
colnames(statisticsFrame) <- c("thermal data", "pp", "HR", "BR", "peda","performance(res)")
row.names(statisticsFrame) <- c("Valid", "Invalid", "Missing", "Not Applicable", "Sum")

#Load Index File for cross check
indexFile <- read.xls("/Users/pranaywankhede/Documents/Statistical/Statistical Assignment/Data/Dataset-Table-Index.xlsx")
indexmatrix = as.matrix(indexFile)
indexdataframe <<- data.frame(indexmatrix)


calculateStatistics <- function() {
  statscols = colnames(statisticsFrame)
  for (i in 1:length(statscols)) {
    colname = as.character(statscols[i]);
    print(exceldataframe)
    colstatslist <- aggregate(data.frame(count = exceldataframe[[colname]]), list(value = exceldataframe[[colname]]), length)
    print(colstatslist)
    colNo <- which(colnames(statisticsFrame) == colname)
    # regexcol <- paste("^",colname,"$",sep="")
    # colindex <- grep(regexcol, colnames(df))
    #print(paste("length", length(colstatslist), sep=""))
    if (nrow(colstatslist) > 0) {
      for (j in 1:nrow(colstatslist)) {
        if (colstatslist$value[j] == 1) {
          print(colstatslist$count[j])
          statisticsFrame[1, colNo] <<- colstatslist$count[j]
        } else if (colstatslist[j]$value == -1) {
          statisticsFrame[2, colNo] <<- colstatslist$count[j]
        } else if (colstatslist[j]$value == 0) {
          statisticsFrame[3, colNo] <<- colstatslist$count[j]
        }
        else if (colstatslist[j]$value == "") {
          statisticsFrame[4, colNo] <<- colstatslist$count[j]
        }
      }
    }
  }
}

# jgc <- function()
# {
#   .jcall("java/lang/System", method = "gc")
# }

plotTimeSeriesRES <- function(zipnames, dataFrame, session, channel) {
  print(session)
  dataframelength <- length(dataFrame)
  plotlist = list()
  counter <- 0
  xlabel <- xlab("")
  
  plot_valid_speed <- ggplot()
  plot_valid_accelaration <- ggplot()
  plot_valid_braking <- ggplot()
  plots_steering_raw <- ggplot()
  plot_lane_position <- ggplot()
  
  
  
  for(i in 1: dataframelength) {
    xlcFreeMemory()
    excelname <- print(paste(zipnames[i], session, channel,".xlsx", sep = ""))
    
    resFrame <- data.frame(matrix(ncol = 8, nrow = 0), stringsAsFactors=F)
    print(paste("row", nrow(resFrame)))
    colnames(resFrame) <- c("Frame", "Time", "Speed", "Acceleration", "Braking", "NR Speed", "NR Acceleration", 
                            "NRBraking")
    isCleaned = FALSE
    
    zipname <- as.character(zipnames[i])
    print(zipname)
    unzip(zipname)
    fullpathwithoutextension = tools::file_path_sans_ext(zipnames[i])
    subject = as.character(fullpathwithoutextension)
    print(paste("Subject", fullpathwithoutextension, sep = " "))
    print(paste("Session", session, sep = " "))
    
    path1 <- as.character(dataFrame[i])
    dataset <- read.xls(path1, skip = 1)
    rawdata <- data.frame(dataset)
    
    #FrameNoframe <- data.frame(dataset$Frame)
    Timeframe <- data.frame(dataset$Time)
    rawspeedframe <- data.frame(dataset$Speed)
    rawaccelarationframe <- data.frame(dataset$Acceleration)
    rawbrakingframe <- data.frame(dataset$Braking)
    rawsteeringframe <- data.frame(dataset$Steering)
    rawLanePosition <- data.frame(dataset$Lane.Position)
    
    validspeedframe <- data.frame(matrix(ncol = 2, nrow = 0), stringsAsFactors=F)
    colnames(validspeedframe) <- c("Time", "Speed")
    validaccelarationframe <- data.frame(matrix(ncol = 2, nrow = 0), stringsAsFactors=F)
    colnames(validaccelarationframe) <- c("Time", "Acceleration")
    validbrakingframe <- data.frame(matrix(ncol = 2, nrow = 0), stringsAsFactors=F)
    colnames(validbrakingframe) <- c("Time", "Braking")
    validsteeringframe <- data.frame(matrix(ncol = 2, nrow = 0), stringsAsFactors=F)
    colnames(validsteeringframe) <- c("Time", "Steering")
    validlanepositionframe <- data.frame(matrix(ncol = 2, nrow = 0), stringsAsFactors=F)
    colnames(validlanepositionframe) <- c("Time", "Lane.Position")
    
    
    dataset_rows <- nrow(rawdata)
    for (row in 1:dataset_rows) {
      #resFrame[row, 1] <- FrameNoframe[row,1]
      resFrame[row, 2] <- Timeframe[row,1]
      rows <- nrow(rawspeedframe)
      #for (row in 1:rows) {
      if(!is.null(rawspeedframe[row,1])) { 
        print("Speed")
        resFrame[row, 3] <- rawspeedframe[row, 1]
        if(is.na(rawspeedframe[row,1])) {
          print("na in speed")
        }
        else if ((rawspeedframe[row, 1] > -0.1 & rawspeedframe[row, 1] < 0.1 )) {
          validspeedframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], 0)
          resFrame[row, 6] <- 0
          isCleaned = TRUE
        } else if (rawspeedframe[row, 1] < -0.1) {
          validspeedframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], NA)
          resFrame[row, 6] <- NA
          isCleaned = TRUE
        } else {
          validspeedframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], rawspeedframe[row, 1])
          resFrame[row, 6] <- rawspeedframe[row, 1]
        }
        
      }#else {
      #validspeedframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], rawspeedframe[row, 1])
      #}
      #}
      
      
      #Acceleration:
      #Timeframe <- data.frame(dataset$Time)
      rows <- nrow(rawaccelarationframe)
      
      # for (row in 1:rows) {
      if(!is.null(rawaccelarationframe[row,1])){
        print("Acceleration")
        resFrame[row, 4] <- rawaccelarationframe[row,1]
        if(is.na(rawaccelarationframe[row,1])){
          print("na in acceleration")
        }
        else if (rawaccelarationframe[row,1] < 0) {
          validaccelarationframe[nrow(validaccelarationframe)+1, ] <- c(Timeframe[row,1], NA)
          resFrame[row, 7] <- NA
          #rawaccelarationframe[row, 7] <- NA
          isCleaned = TRUE
        } else {
          validaccelarationframe[nrow(validaccelarationframe)+1, ] <- c(Timeframe[row,1], 
                                                                        rawaccelarationframe[row,1])
          resFrame[row, 7] <- rawaccelarationframe[row,1]
        }
      }#else {
      #validaccelarationframe[nrow(validaccelarationframe)+1, ] <- c(Timeframe[row,1], rawaccelarationframe[row,1])
      
      #}
      #}
      
      #Braking:
      rows <- nrow(rawbrakingframe)
      #for (row in 1:rows) {
      if(!is.null(rawbrakingframe[row,1])){
        print("Braking")
        resFrame[row, 5] <- rawbrakingframe[row,1]
        if(is.na(rawbrakingframe[row,1])){
          print("na in braking") 
        }
        else if (rawbrakingframe[row,1] > 300) {
          validbrakingframe[nrow(validbrakingframe)+1, ] <- c(Timeframe[row,1], 300)
          resFrame[row, 8] <- 300
          isCleaned = TRUE
        }else {
          validbrakingframe[nrow(validbrakingframe)+1, ] <- c(Timeframe[row,1], rawbrakingframe[row,1])
          resFrame[row, 8] <- rawbrakingframe[row,1]
        }
      }#else {
      #validbrakingframe[nrow(validbrakingframe)+1, ] <- c(Timeframe[row,1], validbrakingframe[row,1])
      
      #}
      #}
      
    } #main for  
    
    #speed
    valid_speed_data <- geom_line(data = validspeedframe, aes(x = Time,y= Speed))
    plot_valid_speed <-  plot_valid_speed + valid_speed_data
    
    #Acceleration:
    valid_accelaration_data <- geom_line(data = validaccelarationframe, aes(x = Time,y = Acceleration))
    plot_valid_accelaration <-  plot_valid_accelaration + valid_accelaration_data
    
    #breaking:
    valid_braking_data <- geom_line(data = validbrakingframe, aes(x = Time,y = Braking))
    plot_valid_braking <-  plot_valid_braking + valid_braking_data
    
    #steering:
    #rawsteeringframe <- data.frame(dataset[,2], dataset[,6])
    geomData <- geom_line(data = dataset, aes(x = dataset$Time ,y= dataset$Steering))
    plots_steering_raw <-  plots_steering_raw + geomData
    
    #Lane position
    geomData <- geom_line(data = dataset, aes(x = dataset$Time ,y= dataset$Lane.Position))
    plot_lane_position <-  plot_lane_position + geomData
    
    excelrowNo <- which(exceldataframe[,1] == subject & exceldataframe[,2] == session)
    excelcolNo <- which(colnames(exceldataframe) == channel)
    if (isCleaned) {
      ttg <- paste("File Should be there:-", session, channel, sep = " ")
      print(ttg)
      oldpath <- getwd()
      setwd(RESPath)
      openwb = createWorkbook(excelname)
      addWorksheet(openwb, "Sheet1")
      writeData(openwb, sheet = 1, resFrame)
      saveWorkbook(openwb, excelname, overwrite = TRUE)
      setwd(oldpath)
      exceldataframe[excelrowNo, excelcolNo] <<- -1
    } else {
      exceldataframe[excelrowNo, excelcolNo] <<- 1
    }
    
    unlink(subject, recursive = TRUE, force = TRUE)
  }
  
  if(grepl(session,"FD")){
    xlabel <- xlab("Time[s]")
  }
  theme1 <- theme(plot.title = element_text(size = 9, face = "bold"),text = element_text(size=9,face = 'bold')) 
  
  plotlist[[1]] <- plot_valid_speed +  ggtitle(paste('Speed[Km/h] for',session))+xlabel+ylab("") + theme1
  plotlist[[2]] <- plot_valid_accelaration +  ggtitle(paste('Accelaration [°] for',session))+xlabel+ylab("")+ theme1
  plotlist[[3]] <- plot_valid_braking + ggtitle(paste('Braking [N] for',session))+xlabel+ylab("")+ theme1
  plotlist[[4]] <- plots_steering_raw + ggtitle(paste('Steering[rad] for',session))+xlabel+ylab("")+ theme1
  plotlist[[5]] <- plot_lane_position + ggtitle(paste('Lane Position[m] for',session))+xlabel+ylab("")+ theme1
  
  return (plotlist)  
}

plotTimeSeriesPeda <-function(zipnames, dataFrame,session,channel){
  l <- length(dataFrame)
  #print(l)
  plots_raw <- ggplot()
  plots_valid <- ggplot()
  mainlist<- list()
  counter <- 0
  setwd(DataPath)
  print(getwd())
  for(i in 1:l) {
    
    zipname <- as.character(zipnames[i])
    unzip(zipname)
    fullpathwithoutextension = tools::file_path_sans_ext(zipnames[i])
    print(paste("Subject", fullpathwithoutextension, sep = " "))
    print(paste("Session", session, sep = " "))
    path1 <- as.character(dataFrame[i])
    subject = as.character(fullpathwithoutextension)
    dataset <- read.xls(path1,skip = 1)
    
    geomData <-geom_line(data = dataset, aes(x = Time,y = Palm.EDA))
    #Raw data:
    plots_raw <-  plots_raw + geomData
    
    #valid data
    vaidframe <-data.frame()
    channelFrame <- data.frame(dataset$Palm.EDA)
    rows <- nrow(channelFrame)
    flag <- TRUE
    
    for (r in 1:rows) {
      if ((channelFrame[r, 1] < 10 | channelFrame[r, 1] > 4700 )){
        flag <- FALSE
      }
    }#for
    
    #All values within range:
    #print(exceldataframe)
    excelrowNo <- which(exceldataframe[,1] == subject & exceldataframe[,2] == session)
    excelcolNo <- which(colnames(exceldataframe) == channel)
    if(flag == TRUE) {
      exceldataframe[excelrowNo, excelcolNo] <<- 1
      #print(exceldataframe)
      counter <- counter+1
      plots_valid <-  plots_valid + geomData
    } else {
      exceldataframe[excelrowNo, excelcolNo] <<- -1
      #print(exceldataframe)
    }
    unlink(subject, recursive = TRUE, force = TRUE)
  }#outer for
  
  
  tit_raw <- paste("Palm EDA [k-ohm] raw signal sets for session=",session,", n=",l,sep = "")
  tit_valid <- paste("Palm EDA [k-ohm] valid signal sets for ssession=",session,", n=",counter,sep = "")
  
  
  mainlist[[1]] <- plots_raw +  ylab('EDA [k-ohm]') +  xlab("Time[s]")+ 
    ggtitle(tit_raw) +
    theme(plot.title = element_text(size = 9, face = "bold"))+
    scale_y_continuous(limits = c(0,NA))
  
  #annotate(geom="text", label="Scatter plot",position="right")
  
  
  
  mainlist[[2]] <- plots_valid + ylab('EDA [k-ohm]') + xlab("Time[s]")+ 
    ggtitle(tit_valid)+
    theme(plot.title = element_text(size = 9, face = "bold"))+
    scale_y_continuous(limits = c(4,NA))
  
  return (mainlist)
  
}



plotTimeSeriesPP <-function(zipnames, dataFrame,session,channel){
  
  l <- length(dataFrame)
  mainlist<- list()
  plots_valid <- ggplot()
  
  for(i in 1: l){
    zipname <- as.character(zipnames[i])
    unzip(zipname)
    fullpathwithoutextension = tools::file_path_sans_ext(zipnames[i])
    path1 <- as.character(dataFrame[i])
    dataset <- read.xls(path1,skip = 1)
    
    geomData <-geom_line(data = dataset, aes(x = dataset[,2],y= dataset[,4]))
    #Raw data:
    plots_valid <-  plots_valid + geomData
    unlink(fullpathwithoutextension, recursive = TRUE, force = TRUE)
    excelrowNo <- which(exceldataframe[,1] == fullpathwithoutextension & exceldataframe[,2] == session)
    excelcolNo <- which(colnames(exceldataframe) == channel)
    exceldataframe[excelrowNo, excelcolNo] <<- 1
  }#outer for
  
  tit_valid <- paste("Perinasal EDA valid signal sets for session=",session,", n=",l,sep = "")
  
  mainlist[[1]] <- plots_valid + xlab("Time[s]") +ylab("0C2") + ggtitle(tit_valid)+
    theme(plot.title = element_text(size = 10, face = "bold"))+
    scale_y_continuous(limits = c(0,0.04))
  
  return(mainlist)
  
}


plotTimeSeriesBR <-function(zipnames, dataFrame,session,channel){
  
  l <- length(dataFrame)
  #print(l)
  plots_raw <- ggplot()
  plots_valid <- ggplot()
  counter <- 0
  mainlist<- list()
  for(i in 1: l){
    zipname <- as.character(zipnames[i])
    unzip(zipname)
    fullpathwithoutextension = tools::file_path_sans_ext(zipnames[i])
    path1 <- as.character(dataFrame[i])
    dataset <- read.xls(path1,skip = 1)
    
    geomData <-geom_line(data = dataset, aes(x = Time,y= Breathing.Rate))
    #Raw data:
    plots_raw <-  plots_raw + geomData
    
    #valid data
    vaidframe <-data.frame()
    channelFrame <- data.frame(dataset$Breathing.Rate)
    Timeframe <- data.frame(dataset$Time)
    rows <- nrow(channelFrame)
    flag <- TRUE
    
    for (r in 1:rows){
      if ((channelFrame[r, 1] < 4 | channelFrame[r, 1] > 70 )){
        
        flag <- FALSE
      }
    }#for
    
    excelrowNo <- which(exceldataframe[,1] == fullpathwithoutextension & exceldataframe[,2] == session)
    excelcolNo <- which(colnames(exceldataframe) == channel)
    #All values within range:
    if(flag == TRUE) {
      counter <- counter+1
      plots_valid <-  plots_valid + geomData
      exceldataframe[excelrowNo, excelcolNo] <<- 1
    } else {
      exceldataframe[excelrowNo, excelcolNo] <<- -1
    }
    unlink(fullpathwithoutextension, recursive = TRUE, force = TRUE)
  }#outer for
  
  
  tit_raw <- paste("BR [bpm] raw signal sets for session=",session,", n=",l,sep = "")
  tit_valid <- paste("BR [bpm] valid signal sets for ssession=",session,", n=",counter,sep = "")
  
  
  mainlist[[1]] <- plots_raw +  ylab('BR [bpm]') +  xlab("Time[s]") +
    ggtitle(tit_raw) +
    theme(plot.title = element_text(size = 10, face = "bold"))+
    scale_y_continuous(limits = c(0,NA))
  
  
  mainlist[[2]] <- plots_valid + ylab('BR [bpm]') + xlab("Time[s]") +
    ggtitle(tit_valid)+
    theme(plot.title = element_text(size = 10, face = "bold"))+
    scale_y_continuous(limits = c(4,NA))
  
  return (mainlist)
  
}

plotDataPP <- function(templist,imgName){
  
  #Plot is saved to pdf file on output path
  
  imgName<- paste(outputPath,imgName,sep = "/")
  imgName <- paste(imgName,".pdf",sep = '')
  n =  length(templist)
  pdf(file=imgName,width = 10,height = 15)
  
  do.call("grid.arrange", c(templist,nrow = n, ncol = 1))
  dev.off()
}

plotData <- function(templist,imgName){
  
  #Plot is saved to pdf file on output path
  
  imgName<- paste(outputPath,imgName,sep = "/")
  imgName <- paste(imgName,".pdf",sep = '')
  
  n =  length(templist)
  pdf(file=imgName,width = 10,height = 15)
  n =  length(templist)/2
  #print("before")
  do.call("grid.arrange", c(templist,nrow = n, ncol = 2))
  
  #grid.arrange(arrangeGrob(templist,ncol=n, nrow=1), main = "Here your title should be inserted",nrow=1)
  dev.off()
}

plotTimeSeriesHR <-function(zipnames, dataFrame, session, channel){
  
  l <- length(dataFrame)
  plots_raw <- ggplot()
  plots_valid <- ggplot()
  mainlist<- list()
  counter <- 0
  
  for(i in 1: l){
    unzip(as.character(zipnames[i]))
    filenamewithoutext = tools::file_path_sans_ext(as.character(zipnames[i]))
    path1 <- as.character(dataFrame[i])
    dataset <- read.xls(path1,skip = 1)
    
    geomData <-geom_line(data = dataset, aes(x = Time,y = Heart.Rate))
    
    #Raw data:
    plots_raw <-  plots_raw + geomData
    #valid data
    vaidframe <-data.frame()
    
    channelFrame <- data.frame(dataset$Heart.Rate)
    rows <- nrow(channelFrame)
    flag <- TRUE
    
    for (r in 1:rows){
      if ((channelFrame[r, 1] < 40 | channelFrame[r, 1] > 140 )){
        flag <- FALSE
      }
    }#for
    
    excelrowNo <- which(exceldataframe[,1] == filenamewithoutext & exceldataframe[,2] == session)
    excelcolNo <- which(colnames(exceldataframe) == channel)
    #All values within range:
    if(flag == TRUE) {
      counter <- counter+1
      plots_valid <-  plots_valid + geomData
      exceldataframe[excelrowNo, excelcolNo] <<- 1
    } else {
      exceldataframe[excelrowNo, excelcolNo] <<- -1
    }
    
    # if (cleaned == TRUE) {
    #         ans <- which(exceldataframe[,1] == filenamewithoutext & exceldataframe[,2] == session)
    #         print(ans)
    #         colNo <- which(colnames(exceldataframe) == "HR")
    #         print(colNo)
    #         exceldataframe[ans, colNo] <- -1
    #         cleaned = FALSE
    #       } else {
    #         # exceldataframe <- within(exceldataframe, HR[Subject == fullpathwithoutextension &
    #         #                                               Session == sessionName] <- "1")
    #         ans <- which(exceldataframe[,1] == filenamewithoutext & exceldataframe[,2] == session)
    #         print(ans)
    #         colNo <- which(colnames(exceldataframe) == "HR")
    #         print(colNo)
    #         exceldataframe[ans, colNo] <- "1"
    #         print(exceldataframe)
    #       }
    
    unlink(filenamewithoutext, recursive = TRUE, force = TRUE) 
  }#outer for
  
  
  tit_raw <- paste("Heart Rate [bpm] raw signal sets for session=",session,", n=",l,sep = "")
  tit_valid <- paste("Heart Rate [bpm] valid signal sets for session=",session,", n=",counter,sep = "")
  
  if(grepl(session,"FD")){
    plots_valid <- plots_valid + xlab("Time[s]")
    plots_raw <- plots_raw + xlab("Time[s]")
    
  }else{
    plots_valid <- plots_valid + xlab("")
    plots_raw <- plots_raw + xlab("")
  }
  
  mainlist[[1]] <- plots_raw +  ylab('HR [bpm]') +
    ggtitle(tit_raw) +
    theme(plot.title = element_text(size = 11, face = "bold"),text = element_text(size=11,face = 'bold'))+
    ylim(0,NA)
  
  
  mainlist[[2]] <- plots_valid + ggtitle(tit_valid)+ylab("")+
    theme(plot.title = element_text(size = 11, face = "bold"),text = element_text(size=11,face = 'bold'))+
    ylim(40,140)
  
  return (mainlist)
  
}

#----------------------Plot for RES------------------------------------------------------------
#------------------------------------------

plotDataRES <- function(templist,imgName){
  print("in plot res")
  
  #Plot is saved to pdf file on output path
  
  imgName<- paste(outputPath,imgName,sep = "/")
  imgName <- paste(imgName,".pdf",sep = '')
  
  n =  length(templist)
  #print(n)
  pdf(file=imgName,width = 12,height = 15)
  n =  length(templist)/5
  do.call("grid.arrange", c(templist,nrow = n, ncol = 5))
  dev.off()
}



# plotHRdata <- function(exceldataframe,zipnames, dataFrame, session, channel) {
#   dataframelength <- length(dataFrame)
#   plots_raw <- ggplot()
#   plots_valid <- ggplot()
#   plotlist = list()
#   cleaned = FALSE
#   
#   for(i in 1: 1) {
#     unzip(as.character(zipnames[i]))
#     filenamewithoutext = tools::file_path_sans_ext(as.character(zipnames[i]))
#     path1 <- as.character(dataFrame[i])
#     dataset <- read.xls(path1, skip = 1)
#     rawdata <- data.frame(dataset)
#     sessionName <- session
#     geomData <- geom_line(data = rawdata, aes(x = Time ,y= Heart.Rate))
#     plots_raw <-  plots_raw + geomData
#     
#     #Valid Data
#     vaidframe <- data.frame( "Time" = integer(0), "Heart.Rate" = integer(0))
#     #vaidframe <- data.frame()
#     heartrateFrame <- data.frame(dataset$Heart.Rate)
#     Timeframe <- data.frame(dataset$Time)
#     rows <- nrow(heartrateFrame)
#     flag = TRUE
#     for (row in 1:rows) {
#       if ((heartrateFrame[row, 1] < 40 | heartrateFrame[row, 1] > 120 )) {
#         flag <- FALSE
#         cleaned = TRUE
#       }
#       
#       if (flag == TRUE) {
#         vaidframe[nrow(vaidframe)+1, ] <- c(Timeframe[row,1], heartrateFrame[row,1])
#       } else {
#         flag = TRUE
#       }
#     }
#     
#     if (cleaned == TRUE) {
#       # exceldataframe <- within(exceldataframe, HR[Subject == fullpathwithoutextension &
#       #                                               Session == sessionName] <- "-1")
#       #print(exceldataframe)
#       ans <- which(exceldataframe[,1] == filenamewithoutext & exceldataframe[,2] == session)
#       print(ans)
#       colNo <- which(colnames(exceldataframe) == "HR")
#       print(colNo)
#       exceldataframe[ans, colNo] <- -1
#       cleaned = FALSE
#     } else {
#       # exceldataframe <- within(exceldataframe, HR[Subject == fullpathwithoutextension &
#       #                                               Session == sessionName] <- "1")
#       ans <- which(exceldataframe[,1] == filenamewithoutext & exceldataframe[,2] == session)
#       print(ans)
#       colNo <- which(colnames(exceldataframe) == "HR")
#       print(colNo)
#       exceldataframe[ans, colNo] <- "1"
#       print(exceldataframe)
#     }
#     
#     valid_geom_data <- geom_line(data = vaidframe, aes(x = Time,y= Heart.Rate))
#     plots_valid <-  plots_valid + valid_geom_data
#     
#     unlink(fullpathwithoutextension, recursive = TRUE, force = TRUE)
#   }
#   
#   plotlist[[1]] <- plots_raw + xlab("Time") +  ylab('HR [bpm]') +
#     ggtitle("Heart Rate [bpm] raw signal sets")
#   
#   plotlist[[2]] <- plots_valid + xlab("Time") +  ylab('HR [bpm]') +
#     ggtitle("Heart Rate [bpm] valid signal sets")
#   #multiplot(plots_raw, plots_valid, cols=2)
#   #assign('exceldataframe',exceldataframe,envir=.GlobalEnv)
#   return (list(plotlist, exceldataframe))
#   #
# }

# plotRESdata <- function(zipnames, dataFrame, session, channel) {
#   dataframelength <- length(dataFrame)
#   plots_raw <- ggplot()
#   plots_valid <- ggplot()
#   plotlist = list()  
#   
#   for(i in 1: 1) {
#     unzip(as.character(zipnames[i]))
#     fullpathwithoutextension = tools::file_path_sans_ext(as.character(zipnames[i]))
#     path1 <- as.character(dataFrame[i])
#     dataset <- read.xls(path1, skip = 1)
#     rawdata <- data.frame(dataset)
#     
#     Timeframe <- data.frame(dataset$Time)
#     rawspeedframe <- data.frame(rawdata$Time, rawdata$Speed)
#     rawaccelarationframe <- data.frame(rawdata$Time, rawdata$Acceleration)
#     rawbrakingframe <- data.frame(rawdata$Time, rawdata$Braking)
#     rawsteeringframe <- data.frame(rawdata$Time, rawdata$Steering)
#     rawLanePosition <- data.frame(rawdata$Time, rawdata$Lane.Position)
#     
#     validspeedframe <- data.frame("Time" = integer(0), "Speed" = integer(0))
#     validaccelarationframe <- data.frame("Time" = integer(0), "Acceleration" = integer(0))
#     validbrakingframe <- data.frame("Time" = integer(0), "Braking" = integer(0))
#     validsteeringframe <- data.frame("Time" = integer(0), "Steering" = integer(0))
#     validlanepositionframe <- data.frame("Time" = integer(0), "Lane Position" = integer(0))
#     
#     
#     
#     rows <- nrow(rawspeedframe)
#     plot_valid_speed <- ggplot()
#     for (row in 1:rows) {
#       if ((rawspeedframe[row, 1] >= -0.1 | rawspeedframe[row, 1] <= 0.1 )) {
#         validspeedframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], 0)
#       } else if (rawspeedframe[row, 1] < 0.1) {
#         validspeedframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], NA)
#       } else {
#         validspeedframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], rawspeedframe[row, 1])
#       }
#     }
#     valid_speed_data <- geom_line(data = validspeedframe, aes(x = Time,y= Speed))
#     plot_valid_speed <-  plot_valid_speed + valid_speed_data
#     
#     rows <- nrow(rawaccelarationframe)
#     plot_valid_accelaration <- ggplot()
#     for (row in 1:rows) {
#       if (rawaccelarationframe[row,1] > 0) {
#         validaccelarationframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], NA)
#       } else {
#         validaccelarationframe[nrow(validspeedframe)+1, ] <- c(Timeframe[row,1], rawaccelarationframe[row,1])
#       }
#     }
#     valid_accelaration_data <- geom_line(data = validaccelarationframe, aes(x = Time,y = Acceleration))
#     plot_valid_accelaration <-  plot_valid_accelaration + valid_accelaration_data
#     
#     rows <- nrow(rawbrakingframe)
#     plot_valid_braking <- ggplot()
#     for (row in 1:rows) {
#       if (rawbrakingframe[row,1] > 300) {
#         validbrakingframe[nrow(validbrakingframe)+1, ] <- c(Timeframe[row,1], 300)
#       } else {
#         validbrakingframe[nrow(validbrakingframe)+1, ] <- c(Timeframe[row,1], validbrakingframe[row,1])
#       }
#     }
#     valid_braking_data <- geom_line(data = validaccelarationframe, aes(x = Time,y = Acceleration))
#     plot_valid_braking <-  plot_valid_braking + valid_braking_data
#     
#     plots_steering_raw <- ggplot()
#     geomData <- geom_line(data = rawsteeringframe, aes(x = Time ,y= Steering))
#     plots_steering_raw <-  plots_steering_raw + geomData
#     
#     plot_lane_position <- ggplot()
#     geomData <- geom_line(data = rawLanePosition, aes(x = Time ,y= Lane.Position))
#     plot_lane_position <-  plot_lane_position + geomData
#     
#     rows <- nrow(rawLanePosition)
#     plot_valid_lane_position <- ggplot()
#     for (row in 1:rows) {
#       if (!(rawLanePosition[row,1] < 0)) {
#         validlanepositionframe[nrow(validlanepositionframe)+1, ] <- c(Timeframe[row,1], validlanepositionframe[row,1])
#       }
#     }
#     validlanepositionframe <- geom_line(data = validlanepositionframe, aes(x = Time,y = Lane.Position))
#     plot_valid_braking <-  plot_valid_braking + validlanepositionframe
#   }
#   plotlist[[1]] <- plot_valid_speed +  ylab('Speed[KM/H]')
#   plotlist[[2]] <- plot_valid_accelaration +  ylab('Accelaration ["]')
#   plotlist[[3]] <- plot_valid_braking + ylab('Braking ["]')
#   plotlist[[4]] <- plots_steering_raw + ylab('Steering[rad]')
#   plotlist[[5]] <- plot_lane_position + ylab('Lane Position[m]')
#   
#   return (plotlist)  
# }



main <-function(){
  setwd(Path)
  curDir<-getwd()
  #debugprint(curDir)
  #channels=c("BR","pp","peda")
  channels=c("pp","HR","BR","peda","res")
  sessions=c("BL","PD","RD","ND","CD","ED","MD","FD")#FDN/FDL"
  sessions_list <- list()
  #sessions=c("RD")#FDN/FDL"
  frames_lst <- list()
  #Check if current working directory has "data" folder which contains all zip folders:
  templist <- list()
  templist2 <- list()
  file.names = list.files(pattern = "*.zip", recursive = F)
  
  if(!is.na(match(TRUE, c(grepl("Data", dir()))))){
    setwd("data")
    dataPath <- getwd()
    
    file.names = list.files(pattern = "*.zip", recursive = F)
    rowcount = 1
    filelength <- length(file.names)
    for (i in 1:1) {
      unzip(file.names[i])
      fullpathwithoutextension = tools::file_path_sans_ext(file.names[i])
      # folderlist <- list.dirs(path = fullpathwithoutextension, recursive = FALSE, 
      #                         full.names = FALSE)
      for (j in 1:length(sessions)) {
        exceldataframe[rowcount,1] <<- fullpathwithoutextension
        #foldername <- sub(".* ", "", folderlist[j])
        # if (foldername == "FDN" | foldername == "FDL")
        #   foldername <- "FD"
        exceldataframe[rowcount,2] <<-  sessions[j]
        exceldataframe[rowcount,3] <<- 1
        rowcount = rowcount + 1
      }
      
      unlink(fullpathwithoutextension, recursive = TRUE, force = TRUE)
    }
    
    #Step 1: Read which channel to plot
    channelLength = length(channels)
    for(c in 1:channelLength){
      zipnames <- list()
      actualChannelName = channels[c]
      #step 2: Read sessions
      sessionsLength = length(sessions)
      for(s in 1: sessionsLength ) {
        sessionName = sessions[s]
        setwd(DataPath)
        #Step 3: Read zip folders
        zipFolder <- dir(pattern = "zip")
        len <- length(zipFolder)    #change this
        
        for(i in 1:1) {
          
          setwd(DataPath)
          unzipF <- gsub(pattern = "\\.zip$", "", zipFolder[i])
          tempDirs <- paste(dataPath,unzipF,sep = '/')
          #print(tempDirs)
          if(! file.exists(tempDirs)){
            #print("file does not exist")
            unzip(zipFolder[i],overwrite = TRUE)
          }
          actualFileName = tools::file_path_sans_ext(zipFolder[i])
          setwd(unzipF)
          
          #Step 4:
          #List of folders in a unzip file: BL-FDN for HR
          folderlist <- list.dirs(path = ".", full.names = FALSE, recursive = FALSE)
          lenSub <- length(folderlist)
          for(j in 1:lenSub) {
            isPresent = FALSE;
            #Step 5: Check if current session
            if(grepl(sessions[s],folderlist[j])){
              sessionPath <- paste(".",folderlist[j],sep = "/")
              #Get file names:
              insideFolderList <- list.files(path = sessionPath, full.names = FALSE, recursive = FALSE)
              lenFileList <- length(insideFolderList)
              #print(insideFolderList)
              #step 6: current channel
              #print(insideFolderList)
              for(f in 1:lenFileList) {
                #Check for needed data channel:
                if(grepl(channels[c],insideFolderList[f]) & grepl("^T",insideFolderList[f])) {
                  isPresent = TRUE
                  
                  filePath <- paste(sessionPath,insideFolderList[f],sep = "/")
                  #print("***************")
                  #print(filePath)
                  frame_len <- length(frames_lst)+1
                  zipnames[frame_len] = file.names[i]
                  #set full path:
                  temp_path <-  as.character(paste(getwd(), gsub(pattern = "^.","",filePath),sep = ""))
                  #print(temp_path)
                  
                  frames_lst[frame_len] <- temp_path
                  break
                }
              }
              if (isPresent == FALSE) {
                print(paste("Subject", file.names[i], sep = " "))
                print(paste("Session", sessionName, sep = " "))
                rowNo <- which(indexdataframe[,1] == actualFileName & indexdataframe[,2] == sessionName)
                colNo <- which(colnames(indexdataframe) == actualChannelName)
                excelrowNo <- which(exceldataframe[,1] == actualFileName & exceldataframe[,2] == sessionName)
                excelcolNo <- which(colnames(exceldataframe) == actualChannelName)
                value <<- indexdataframe[rowNo, colNo]
                print(value)
                if (length(value) == 0) {
                  rowNo <- which(indexdataframe[,1] == actualFileName & indexdataframe[,2] == "FDN")
                  colNo <- which(colnames(indexdataframe) == actualChannelName)
                  value = indexdataframe[rowNo, colNo]
                  if (length(value) == 0) {
                    rowNo <- which(indexdataframe[,1] == actualFileName & indexdataframe[,2] == "FDL")
                    colNo <- which(colnames(indexdataframe) == actualChannelName)
                    value <<- indexdataframe[rowNo, colNo]
                  }
                  #print(value)
                }
                if (length(value) == 0) {
                  exceldataframe[excelrowNo, excelcolNo] <<- "NA"
                }
                else if (is.na(value)) {
                  #print(value)
                  exceldataframe[excelrowNo, excelcolNo] <<- "NA"
                }
                else if (value == 0) {
                  print(value)
                  exceldataframe[excelrowNo, excelcolNo] <<- 0
                }
              }
            }
          }
          setwd(DataPath)
          dataPath <- getwd()
          fullpathwithoutextension = tools::file_path_sans_ext(file.names[i])
          unlink(fullpathwithoutextension, recursive = TRUE, force = TRUE)
        }
        
        if(length(frames_lst) != 0) {
          if(grepl(channels[c],"BR")){
            templist <- plotTimeSeriesBR(zipnames, frames_lst, sessions[s], channels[c])
          }else if(grepl(channels[c],"pp")){
            templist <- plotTimeSeriesPP(zipnames,frames_lst,sessions[s],channels[c])
          }else if(grepl(channels[c],"peda")){
            templist <- plotTimeSeriesPeda(zipnames,frames_lst,sessions[s],channels[c])
          }else if(grepl(channels[c],"HR")){
            templist <- plotTimeSeriesHR(zipnames,frames_lst,sessions[s],channels[c])
          }else if(grepl(channels[c],"res")){
            templist <- plotTimeSeriesRES(zipnames,frames_lst,sessions[s],channels[c])
          }
        }
        
        frames_lst <- list()  
        i = 0
        n =  length(templist2)
        while (i != length(templist)) {
          n = n + 1
          i = i + 1
          templist2[n] = templist[i]
        }
        templist <- list()
      }#session
      
      if(grepl(channels[c],"pp")){
        plotDataPP(templist2,channels[c])
      }else if(grepl(channels[c],"res")){
        plotDataRES(templist2,channels[c])
      }else{
        plotData(templist2,channels[c])
      }
      
      templist2 <- list()
      
      # if(grepl(channels[c],"pp")){
      #   plotDataPP(templist2,channels[c])
      # }else{
      #   plotData(templist2,channels[c])
      # }
      # calculateStatistics();
      # sessions_list<-list()
      # print(exceldataframe)
      # setwd("/Users/pranaywankhede/Documents/Statistical/Statistical Assignment/Data")
      # dataPath <- getwd()
      excelname = print(paste(actualChannelName,".xlsx", sep = ""))
      wb = createWorkbook(excelname)
      addWorksheet(wb, "Sheet1")
      addWorksheet(wb, "Sheet2")
      setwd(Path)
      print(getwd())
      writeData(wb, sheet = "Sheet1", exceldataframe)
      #calculateStatistics()
      #writeData(wb, sheet = "Sheet2", statisticsFrame)
      saveWorkbook(wb, file = excelname, overwrite = TRUE)
    }
  } else
  {
    print("data folder not found")
  }
  
  calculateStatistics()
  writeData(wb, sheet = "Sheet2", statisticsFrame)
  saveWorkbook(wb, file = excelname, overwrite = TRUE)
  
  # calculateStatistics()
  # setwd(Path)
  # print(getwd())
  # print(exceldataframe)
  # writeData(wb, sheet = 1, exceldataframe)
  # saveWorkbook(wb, "MydataSheet.xlsx", overwrite = TRUE)
  # statisticsFrame <- cbind(Values = rownames(statisticsFrame), statisticsFrame)
  # writeWorksheet(wb, exceldataframe, sheet = "DataSheet", startRow = 1, startCol = 1)
  # print(statisticsFrame)
  # appendWorksheet(wb,statisticsFrame, sheet = "DataSheet", header = TRUE, rownames = TRUE)
  # saveWorkbook(wb)
  #ddply(myvec,~exceldataframe$pp,summarise,number_of_valid=length(unique(order_no)))
  
}#function main

#call function:
main()