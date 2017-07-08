#Retrieving passed arguments

myArgs <- commandArgs(trailingOnly = TRUE)
slash <- myArgs[1]
data_folder <- myArgs[2]
config_path <- myArgs[3]
r_scripts <- myArgs[4]
r_libs <- myArgs[5]
log_file <- myArgs[6]

setwd(data_folder)

#Loading libraries

package_list <- c("Rcpp","openxlsx")
	for( i in 1:length(package_list))
	{
		if (!require(package_list[i],character.only = TRUE,lib.loc=r_libs))
    			{
		      	install.packages(package_list[i],dep=TRUE,repos="http://cran.us.r-project.org",lib=r_libs)
			        if(!require(package_list[i],character.only = TRUE,lib.loc=r_libs)) stop("Package not found")
			}
	}

#Loading configuration file

config_data <- readLines(config_path)

run_analysis_1 <- config_data[7]=="[True]"

analysis_1_1_string <- strsplit(substring(config_data[2],2,nchar(config_data[2])-1),", ")[[1]]
param_1_1 <- vector(length = length(analysis_1_1_string))
for (i in 1:length(analysis_1_1_string))
    {
        param_1_1[i] <- analysis_1_1_string[i] == "True"
    }

param_1_2 <- strsplit(substring(config_data[3],3,nchar(config_data[3])-2),"', '")[[1]]

param_1_3 <- strsplit(substring(config_data[4],2,nchar(config_data[4])-2),", '")[[1]]
param_1_3[1] <- param_1_3[1] == "True"

param_1_4 <- strsplit(substring(config_data[5],2,nchar(config_data[5])-2),", '")[[1]]
param_1_4[1] <- param_1_4[1] == "True"

analysis_1_5_string <- strsplit(substring(config_data[6],2,nchar(config_data[6])-1), ", ")[[1]]
param_1_5 <- vector(length = length(analysis_1_5_string))
for (i in 1:length(analysis_1_5_string))
    {
        param_1_5[i] <- analysis_1_5_string[i] == "True"
    }

params_1<- list(run_analysis_1,param_1_1,param_1_2,param_1_3,param_1_4,param_1_5)

#Creating Workbook

workbook <- createWorkbook()

#Analysis 1

if(params_1[[1]]==TRUE)
    {
        seguros <- read.csv2(paste0("data",slash,"Ses_seguros.csv"))

    #Subsetting py Selected Params

        subset_params_1 <- c(TRUE,TRUE,TRUE,TRUE, params_1[[2]][c(1,3,4,6,8,9,14,5,2,7,10,11,12,13,16,17,18,15)])
        seguros <- seguros[,subset_params_1]

    #Subsetting by Initial Date

        if(nchar(params_1[[3]][1])==10)
            {
                init_date_1 <- as.Date(params_1[[3]][1], format="%d/ %m/ %Y")
                if(!is.na(init_date_1))
                    {
                        init_date_1 <- format(init_date_1,"%Y%m")
                        if(nchar(init_date_1)==6)
                            {
                                seguros <- seguros[seguros$damesano >= as.numeric(init_date_1),]
                            }
                    }
            }

    #Subsetting by Final Date
        if(nchar(params_1[[3]][2])==10)
            {
                final_date_1 <- as.Date(params_1[[3]][2], format="%d/ %m/ %Y")
                if(!is.na(final_date_1))
                    {
                        final_date_1 <- format(final_date_1,"%Y%m")
                        if(nchar(final_date_1)==6)
                            {
                                seguros <- seguros[seguros$damesano <= as.numeric(final_date_1),]
                            }
                    }
            }

    #Subsetting by Company
        if(params_1[[4]][1]=="FALSE")
            {
                coenti_1 <- as.numeric(strsplit(params_1[[4]][2],",")[[1]])
                seguros <- seguros[seguros$coenti %in% coenti_1,]
            }

    #Subsetting by Ramos
        if(params_1[[5]][1]=="FALSE")
            {
                ramos_1 <- as.numeric(strsplit(params_1[[5]][2],",")[[1]])
                seguros <- seguros[seguros$coramo %in% ramos_1,]
            }
    #Aggreggating by time interval
        
#        test <- paste0("0",as.character(ceiling(as.numeric(format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%m"))/6)))
        #Aggregate by Year/Month
        if(params_1[[6]][1] == TRUE)
            {
                seguros<-aggregate(seguros[,5:ncol(seguros)],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(seguros$damesano)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/12
                    )
                    )
                    )
                                            ),
                                           coenti=seguros$coenti,
                                           coramo=seguros$coramo,
                                           cogrupo=seguros$cogrupo
                                           ),FUN=sum,na.rm=TRUE)
            }
        #Aggregate by Year/Trimester
        else if(params_1[[6]][2] == TRUE)
            {
                seguros<-aggregate(seguros[,5:ncol(seguros)],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(seguros$damesano)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/3
                    )
                    )
                    )
                                            ),
                                           coenti=seguros$coenti,
                                           coramo=seguros$coramo,
                                           cogrupo=seguros$cogrupo
                                           ),FUN=sum,na.rm=TRUE)
            }
        #Aggregate by Year/Semester
        else if(params_1[[6]][3] == TRUE)
            {
                seguros<-aggregate(seguros[,5:ncol(seguros)],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(seguros$damesano)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/6
                    )
                    )
                    )
                                            ),
                                           coenti=seguros$coenti,
                                           coramo=seguros$coramo,
                                           cogrupo=seguros$cogrupo
                                           ),FUN=sum,na.rm=TRUE)
            }
        #Aggregate by Year
        else if(params_1[[6]][4] == TRUE)
            {
                seguros<-aggregate(seguros[,5:ncol(seguros)],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                "01"),
                                           coenti=seguros$coenti,
                                           coramo=seguros$coramo,
                                           cogrupo=seguros$cogrupo
                                           ),FUN=sum,na.rm=TRUE)
            }

    #Writing file

        write.csv2(seguros,paste0("proc_data",slash,"premios_sinistros_seg.csv"))
        addWorksheet(workbook,"premios_sinistros_seg")
        writeDataTable(workbook,"premios_sinistros_seg",format(seguros,decimal.mark=","))
    }

saveWorkbook(workbook,paste0("SES-Susep-",Sys.Date(),".xlsx"),overwrite=TRUE)


#nrow(seguros)
#seguros[,5:ncol(seguros)]
#seguros[,5:ncol(seguros)]<-as.numeric(sub(",",".",as.character(seguros[,5:ncol(seguros)]),fixed=TRUE))
#seguros[,5:ncol(seguros)]
#seguros<-aggregate(seguros[,5:ncol(seguros)],by=list(yearsec=paste0(format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%Y"),"01"),coenti=seguros$coenti,coramo=seguros$coramo,cogrupo=seguros$cogrupo),FUN=sum,na.rm=TRUE)
#nrow(seguros)
#format(seguros,decimal.mark=",")
#head(seguros)
#as.numeric(format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%m"))
#test
