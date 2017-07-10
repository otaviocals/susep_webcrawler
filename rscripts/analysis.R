##############################
#    Retrieving arguments    #
##############################

myArgs <- commandArgs(trailingOnly = TRUE)
slash <- myArgs[1]
data_folder <- myArgs[2]
config_path <- myArgs[3]
r_scripts <- myArgs[4]
r_libs <- myArgs[5]
log_file <- myArgs[6]

setwd(data_folder)

#Loading libraries

package_list <- c("Rcpp","openxlsx","stringi")
	for( i in 1:length(package_list))
	{
		if (!require(package_list[i],character.only = TRUE,lib.loc=r_libs))
    			{
		      	install.packages(package_list[i],dep=TRUE,repos="http://cran.us.r-project.org",lib=r_libs)
			        if(!require(package_list[i],character.only = TRUE,lib.loc=r_libs)) stop("Package not found")
			}
	}


#########################
#    Loading configs    #
#########################

config_data <- readLines(config_path)

#Analysis 1 config

run_analysis_1 <- config_data[7]=="[True]"

#Param_1_1
analysis_1_1_string <- strsplit(substring(config_data[2],2,nchar(config_data[2])-1),", ")[[1]]
param_1_1 <- vector(length = length(analysis_1_1_string))
for (i in 1:length(analysis_1_1_string))
    {
        param_1_1[i] <- analysis_1_1_string[i] == "True"
    }

#Param_1_2
param_1_2 <- strsplit(substring(config_data[3],3,nchar(config_data[3])-2),"', '")[[1]]

#Param_1_3
param_1_3 <- strsplit(substring(config_data[4],2,nchar(config_data[4])-2),", '")[[1]]
param_1_3[1] <- param_1_3[1] == "True"

#Param_1_4
param_1_4 <- strsplit(substring(config_data[5],2,nchar(config_data[5])-2),", '")[[1]]
param_1_4[1] <- param_1_4[1] == "True"

#Param_1_5
analysis_1_5_string <- strsplit(substring(config_data[6],2,nchar(config_data[6])-1), ", ")[[1]]
param_1_5 <- vector(length = length(analysis_1_5_string))
for (i in 1:length(analysis_1_5_string))
    {
        param_1_5[i] <- analysis_1_5_string[i] == "True"
    }

#Final params_1
params_1<- list(run_analysis_1,param_1_1,param_1_2,param_1_3,param_1_4,param_1_5)



#Analysis 2 config

run_analysis_2 <- config_data[15]=="[True]"

#Param_2_1
analysis_2_1_string <- strsplit(substring(config_data[9],2,nchar(config_data[9])-1),", ")[[1]]
param_2_1 <- vector(length = length(analysis_2_1_string))
for (i in 1:length(analysis_2_1_string))
    {
        param_2_1[i] <- analysis_2_1_string[i] == "True"
    }

#Param_2_2
param_2_2 <- strsplit(substring(config_data[10],3,nchar(config_data[10])-2),"', '")[[1]]

#Param_2_3
param_2_3 <- strsplit(substring(config_data[11],2,nchar(config_data[11])-2),", '")[[1]]
param_2_3[1] <- param_2_3[1] == "True"

#Param_2_4
param_2_4 <- strsplit(substring(config_data[12],2,nchar(config_data[12])-2),", '")[[1]]
param_2_4[1] <- param_2_4[1] == "True"

#Param_2_5
param_2_5 <- strsplit(substring(config_data[13],2,nchar(config_data[13])-2),", '")[[1]]
param_2_5[1] <- param_2_5[1] == "True"

#Param_2_6
analysis_2_6_string <- strsplit(substring(config_data[14],2,nchar(config_data[14])-1), ", ")[[1]]
param_2_6 <- vector(length = length(analysis_2_6_string))
for (i in 1:length(analysis_2_6_string))
    {
        param_2_6[i] <- analysis_2_6_string[i] == "True"
    }

params_2<- list(run_analysis_2,param_2_1,param_2_2,param_2_3,param_2_4,param_2_5,param_2_6)

#Analysis 3 config

run_analysis_3 <- config_data[17]=="[True]"
text_data <- config_data[18:length(config_data)]

params_3<- list(run_analysis_3,text_data)

#Creating Workbook

workbook <- createWorkbook()

####################
#    Analysis 1    #
####################

if(params_1[[1]]==TRUE)
    {
        seguros <- read.csv2(paste0("data",slash,"Ses_seguros.csv"))

    #Subsetting py Selected Params
        subset_params_1 <- c(TRUE,TRUE,TRUE,TRUE, params_1[[2]][c(1,3,4,6,8,9,14,5,2,7,10,11,12,13,16,17,18)])
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
                    )/1
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

####################
#    Analysis 2    #
####################

if(params_2[[1]]==TRUE)
    {
        lista_cmpid <- list(c(12032,6475,6807), c(11278,11275), c(6808,6476), c(6809,12028,12078,12340,12349,12364,6477,12204,12210,12246,12251), c(6810,6478), c(6844,7202,7139), c(7231,6816,7147,6481,6811,1613), c(6817,6470,6812,1612,4034), c(6818,6582,6813,4035), c(6819,6583,6814,4036), c(11247,11271,5988,6001,6014,6026,12689,12223,12261,1779), c(6198,6252,6820,6815,3235), c(11267,11322,7267))

        lista_labels <- list(c("Premios_de_Resseguro"), c("Premios_de_Resseguro_RVNE"), c("Operacoes_com_Exterior"), c("Retrocessoes_Aceitas"), c("Premios_Cedidos_de_Retrocessao"), c("Variacoes_das_Provisoes_Tecnicas_de_Premios"), c("Sinistros"), c("Recuperacao_de_Sinistros"), c("Salvados"), c("Ressarcimentos"), c("Salvados_e_Ressarcimentos"), c("Variacoes_das_Provisoes_de_Sinistros_Ocorridos_mas_nao_Avisados"), c("Variacoes_das_Provisoes_IBNER"))

        cmpid_filter <- c()
        labels_filter <- c()
        for(i in 1:length(lista_cmpid))
            {
                #cmpid_filter <- c(cmpid_filter,lista_cmpid[[]])
                if(params_2[[2]][i]== TRUE)
                    {
                        cmpid_filter <- c(cmpid_filter,lista_cmpid[[i]])
                        labels_filter <- c(labels_filter, rep(lista_labels[[i]],length(lista_cmpid[[i]]) ) )
                        test <- rep(lista_labels,length(lista_cmpid[i]))
                    }
            }
        filter_table <- data.frame(cmpid=cmpid_filter,label=labels_filter)

        mov_grupos <- read.csv2(paste0("data",slash,"ses_valoresresmovgrupos.csv"))
        mov_ramos <- read.csv2(paste0("data",slash,"SES_ValoresMovRamos.csv"))

        #Subsetting by CMPID
    
        mov_grupos <- mov_grupos[mov_grupos$CMPID %in% filter_table[,1], ]
        mov_ramos <- mov_ramos[mov_ramos$cmpid %in% filter_table[,1], ]

    #Subsetting by Initial Date

        if(nchar(params_2[[3]][1])==10)
            {
                init_date_2 <- as.Date(params_2[[3]][1], format="%d/ %m/ %Y")
                if(!is.na(init_date_2))
                    {
                        init_date_2 <- format(init_date_2,"%Y%m")
                        if(nchar(init_date_2)==6)
                            {
                                mov_grupos <- mov_grupos[mov_grupos$DAMESANO >= as.numeric(init_date_2),]
                                mov_ramos <- mov_ramos[mov_ramos$damesano >= as.numeric(init_date_2),]
                            }
                    }
            }

    #Subsetting by Final Date
        if(nchar(params_2[[3]][2])==10)
            {
                final_date_2 <- as.Date(params_2[[3]][2], format="%d/ %m/ %Y")
                if(!is.na(final_date_2))
                    {
                        final_date_2 <- format(final_date_2,"%Y%m")
                        if(nchar(final_date_2)==6)
                            {
                                mov_grupos <- mov_grupos[mov_grupos$DAMESANO <= as.numeric(final_date_2),]
                                mov_ramos <- mov_ramos[mov_ramos$damesano <= as.numeric(final_date_2),]
                            }
                    }
            }

    #Subsetting by Company
        if(params_2[[4]][1]=="FALSE")
            {
                coenti_2 <- as.numeric(strsplit(params_2[[4]][2],",")[[1]])
                mov_grupos <- mov_grupos[mov_grupos$COENTI %in% coenti_2,]
                mov_ramos <- mov_ramos[mov_ramos$coenti %in% coenti_2,]
            }

    #Subsetting by Ramos
        if(params_2[[5]][1]=="FALSE")
            {
                ramos_2 <- as.numeric(strsplit(params_2[[5]][2],",")[[1]])
                mov_ramos <- mov_ramos[mov_ramos$ramcodigo %in% ramos_2,]
            }

    #Subsetting by Grupos
        if(params_2[[6]][1]=="FALSE")
            {
                grupos_2 <- as.numeric(strsplit(params_2[[6]][2],",")[[1]])
                mov_grupos <- mov_grupos[mov_grupos$GRACODIGO %in% grupos_2,]
                mov_ramos <- mov_ramos[mov_ramos$gracodigo %in% grupos_2,]
            }

    #Aggreggating by time interval
        
        #Aggregate by Year/Month
        if(params_2[[7]][1] == TRUE)
            {
                mov_grupos<-aggregate(mov_grupos[,5],
                                   by=list(
                                           YEARSEC=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_grupos$DAMESANO),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_grupos$DAMESANO)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/1
                    )
                    )
                    )
                                            ),
                                           COENTI=mov_grupos$COENTI,
                                           CMPID=mov_grupos$CMPID,
                                           GRACODIGO=mov_grupos$GRACODIGO
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_grupos)[5]<-"VALOR"

                mov_ramos<-aggregate(mov_ramos[,6],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_ramos$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_ramos$damesano)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/1
                    )
                    )
                    )
                                            ),
                                           coenti=mov_ramos$coenti,
                                           cmpid=mov_ramos$cmpid,
                                           ramcodigo=mov_ramos$ramcodigo,
                                           gracodigo=mov_ramos$gracodigo,
                                           seq=mov_ramos$seq,
                                           quadro=mov_ramos$quadro
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_ramos)[8]<-"valor"
                mov_ramos <- mov_ramos[,c(1,2,3,4,5,8,6,7)]
            }
        #Aggregate by Year/Trimester
        else if(params_2[[7]][2] == TRUE)
            {
                mov_grupos<-aggregate(mov_grupos[,5],
                                   by=list(
                                           YEARSEC=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_grupos$DAMESANO),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_grupos$DAMESANO)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/3
                    )
                    )
                    )
                                            ),
                                           COENTI=mov_grupos$COENTI,
                                           CMPID=mov_grupos$CMPID,
                                           GRACODIGO=mov_grupos$GRACODIGO
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_grupos)[5]<-"VALOR"

                mov_ramos<-aggregate(mov_ramos[,6],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_ramos$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_ramos$damesano)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/3
                    )
                    )
                    )
                                            ),
                                           coenti=mov_ramos$coenti,
                                           cmpid=mov_ramos$cmpid,
                                           ramcodigo=mov_ramos$ramcodigo,
                                           gracodigo=mov_ramos$gracodigo,
                                           seq=mov_ramos$seq,
                                           quadro=mov_ramos$quadro
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_ramos)[8]<-"valor"
                mov_ramos <- mov_ramos[,c(1,2,3,4,5,8,6,7)]
            }
        #Aggregate by Year/Semester
        else if(params_2[[7]][3] == TRUE)
            {
                mov_grupos<-aggregate(mov_grupos[,5],
                                   by=list(
                                           YEARSEC=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_grupos$DAMESANO),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_grupos$DAMESANO)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/6
                    )
                    )
                    )
                                            ),
                                           COENTI=mov_grupos$COENTI,
                                           CMPID=mov_grupos$CMPID,
                                           GRACODIGO=mov_grupos$GRACODIGO
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_grupos)[5]<-"VALOR"

                mov_ramos<-aggregate(mov_ramos[,6],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_ramos$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_ramos$damesano)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/6
                    )
                    )
                    )
                                            ),
                                           coenti=mov_ramos$coenti,
                                           cmpid=mov_ramos$cmpid,
                                           ramcodigo=mov_ramos$ramcodigo,
                                           gracodigo=mov_ramos$gracodigo,
                                           seq=mov_ramos$seq,
                                           quadro=mov_ramos$quadro
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_ramos)[8]<-"valor"
                mov_ramos <- mov_ramos[,c(1,2,3,4,5,8,6,7)]
            }
        #Aggregate by Year
        else if(params_2[[7]][4] == TRUE)
            {
                mov_grupos<-aggregate(mov_grupos[,5],
                                   by=list(
                                           YEARSEC=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_grupos$DAMESANO),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_grupos$DAMESANO)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/12
                    )
                    )
                    )
                                            ),
                                           COENTI=mov_grupos$COENTI,
                                           CMPID=mov_grupos$CMPID,
                                           GRACODIGO=mov_grupos$GRACODIGO
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_grupos)[5]<-"VALOR"

                mov_ramos<-aggregate(mov_ramos[,6],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(mov_ramos$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                paste0("0",
                as.character(
                ceiling(
                as.numeric(
                format(
                as.Date(
                paste0(
                as.character(mov_ramos$damesano)
                    ,"01")
                    ,format="%Y%m%d")
                    ,"%m")
                    )/12
                    )
                    )
                    )
                                            ),
                                           coenti=mov_ramos$coenti,
                                           cmpid=mov_ramos$cmpid,
                                           ramcodigo=mov_ramos$ramcodigo,
                                           gracodigo=mov_ramos$gracodigo,
                                           seq=mov_ramos$seq,
                                           quadro=mov_ramos$quadro
                                           ),FUN=sum,na.rm=TRUE)

                colnames(mov_ramos)[8]<-"valor"
                mov_ramos <- mov_ramos[,c(1,2,3,4,5,8,6,7)]
            }

    #Writing file

        write.csv2(mov_grupos,paste0("proc_data",slash,"resseg_grupos.csv"))
        addWorksheet(workbook,"resseg_grupos")
        writeDataTable(workbook,"resseg_grupos",format(mov_grupos,decimal.mark=","))

        write.csv2(mov_grupos,paste0("proc_data",slash,"resseg_ramos.csv"))
        addWorksheet(workbook,"resseg_ramos")
        writeDataTable(workbook,"resseg_ramos",format(mov_ramos,decimal.mark=","))

    }

####################
#    Analysis 3    #
####################

if(params_3[[1]]==TRUE)
    {
    #Loading and cleaning Ses_campos.csv

        campos <- read.csv2(paste0("data",slash,"Ses_campos.csv"),encoding="latin1",strip.white=TRUE)
        campos$noitem <- sapply(campos$noitem, function(x) stri_trans_general(x,"Latin-ASCII"))

    #Getting filtered cmpids
        campos_filter <- campos[campos$noitem %in% params_3[[2]], 1]

    #Filtering data

        mov_grupos <- read.csv2(paste0("data",slash,"ses_valoresresmovgrupos.csv"))
        mov_ramos <- read.csv2(paste0("data",slash,"SES_ValoresMovRamos.csv"))

        mov_grupos <- mov_grupos[mov_grupos$CMPID %in% campos_filter, ]
        mov_ramos <- mov_ramos[mov_ramos$cmpid %in% campos_filter, ]

    #Writing file

        write.csv2(mov_grupos,paste0("proc_data",slash,"dem_cont_grupos.csv"))
        addWorksheet(workbook,"dem_cont_grupos")
        writeDataTable(workbook,"dem_cont_grupos",format(mov_grupos,decimal.mark=","))

        write.csv2(mov_grupos,paste0("proc_data",slash,"dem_cont_ramos.csv"))
        addWorksheet(workbook,"dem_cont_ramos")
        writeDataTable(workbook,"dem_cont_ramos",format(mov_ramos,decimal.mark=","))

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
#cmpid_filter
#labels_filter
#head(mov_grupos)
#head(mov_ramos)
#mov_grupos$COENTI
#params_2[[7]]
#params_3
#head(campos)
#head(normalized)
#head(normalized_ascii)
#stri_trans_list()
#head(campos_filter)
