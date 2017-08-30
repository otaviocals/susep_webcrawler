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

package_list <- c("Rcpp","openxlsx","stringi",
                  "labeling","ggplot2","digest",
                  "reshape2","stringr","grid","gridExtra")
	suppressWarnings(suppressMessages(
    for( i in 1:length(package_list))
	{
		if (!require(package_list[i],character.only = TRUE,lib.loc=r_libs))
    			{
		      	install.packages(package_list[i],dep=TRUE,repos="http://cran.us.r-project.org",lib=r_libs)
			        if(!require(package_list[i],character.only = TRUE,lib.loc=r_libs)) stop("Package not found")
			}
	}))



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

#Param_3_1
analysis_3_1_string <- strsplit(substring(config_data[18],2,nchar(config_data[18])-1),", ")[[1]]
param_3_1 <- vector(length = length(analysis_3_1_string))
for (i in 1:length(analysis_3_1_string))
    {
        param_3_1[i] <- analysis_3_1_string[i] == "True"
    }

#Param_3_2
param_3_2 <- strsplit(substring(config_data[19],3,nchar(config_data[19])-2),"', '")[[1]]

#Param_3_3
param_3_3 <- strsplit(substring(config_data[20],2,nchar(config_data[20])-2),", '")[[1]]
param_3_3[1] <- param_3_3[1] == "True"

#Param_3_4
analysis_3_4_string <- strsplit(substring(config_data[21],2,nchar(config_data[21])-1), ", ")[[1]]
param_3_4 <- vector(length = length(analysis_3_4_string))
for (i in 1:length(analysis_3_4_string))
    {
        param_3_4[i] <- analysis_3_4_string[i] == "True"
    }

params_3<- list(run_analysis_3,param_3_1,param_3_2,param_3_3,param_3_4)

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
        if(params_1[[6]][1] == TRUE && nrow(seguros) > 0)
            {
                seguros<-aggregate(seguros[,5:ncol(seguros)],
                                   by=list(
                                           yearsec=paste0(
                #Year
                format(as.Date(paste0(as.character(seguros$damesano),"01"),format="%Y%m%d"),"%Y"),
                #Section
                str_pad(
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
                    ,2,pad="0")
                                            ),
                                           coenti=seguros$coenti,
                                           coramo=seguros$coramo,
                                           cogrupo=seguros$cogrupo
                                           ),FUN=sum,na.rm=TRUE)
            }
        #Aggregate by Year/Trimester
        else if(params_1[[6]][2] == TRUE && nrow(seguros) > 0)
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
        else if(params_1[[6]][3] == TRUE && nrow(seguros) > 0)
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
        else if(params_1[[6]][4] == TRUE && nrow(seguros) > 0)
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

    #Post-processing & Writing file
        seguros <- aggregate(seguros[,5:ncol(seguros)],
                             by=list(
                                     yearsec=seguros$yearsec,
                                     coenti=seguros$coenti
                                     )
                             ,FUN=sum,na.rm=TRUE)


        write.csv2(seguros,paste0("proc_data",slash,"premios_sinistros_seg.csv"))
        addWorksheet(workbook,"premios_sinistros_seg")

    #Merging All Companies
        if(params_1[[4]][1]=="TRUE")
            {
                seguros <- aggregate(seguros[,3:ncol(seguros)],
                                     by=list(yearsec=seguros$yearsec)
                                     ,FUN=sum,na.rm=TRUE)
                if(!params_1[[6]][1])
                    {
                        seguros <- seguros[1:(nrow(seguros)-1),]
                    }

                insert_line <- 1
                numCol <- ceiling((ncol(seguros)-1)/3)

                writeData(workbook,
                          "premios_sinistros_seg",
                          "Todas as Empresas",
                          startRow=insert_line)


                insert_line <- insert_line+1

            #Transposing Data
                seguros_t <- seguros[,-1]
                rownames(seguros_t) <- seguros[,1]
                seguros_t <- as.data.frame(t(seguros_t))

                writeDataTable(workbook,
                               "premios_sinistros_seg",
                               format(
                                seguros_t,
                                decimal.mark=","),
                               rowNames=TRUE,
                               startRow=insert_line)

                insert_line <- insert_line+nrow(seguros_t)+1

                writeData(workbook,
                          "premios_sinistros_seg",
                          "Ver Tabela 1",
                          startRow=insert_line)
                insert_line <- insert_line + 2

            #Removing old files
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^premios_sinistros_seg_",
                                 full.names=TRUE))
            #Plotting data
                plot_data <- melt(seguros,id="yearsec")
                period <- c("Mensal",
                                   "Trimestral",
                                   "Semestral",
                                   "Anual")[params_1[[6]]]

                x_index <- rep(F,nrow(seguros))
                x_length <- length(x_index)
                if(ceiling(x_length/3)!=x_length/3)
                    {
                        first_third <- ceiling(x_length/3)
                    }
                else
                    {
                        first_third <- (x_length/3)+1
                    }
                if(floor(x_length*2/3)!=x_length*2/3)
                    {
                        second_third <- floor(x_length*2/3)+1
                    }
                else
                    {
                        second_third <- x_length*2/3
                    }
                x_index[c(1,
                        first_third,
                        second_third,
                        length(x_index))
                        ]<-T

                p1 <- ggplot(data=plot_data,
                             aes(x=yearsec,y=value,colour=variable,group=variable))+
                                geom_line()+
                                scale_x_discrete(
                                    breaks=plot_data$yearsec[x_index]
                                    )+
                                facet_wrap(~variable, scales="free_y",ncol=3)+
                                ggtitle(paste0(
                                    "Tabela 1: Todas as Empresas\nPeriodicidade ",
                                    period,
                                    " \nDe ",
                                    seguros[1,1],
                                    " a ",
                                    seguros[nrow(seguros),1])
                                )+
                                ylab("Valor")+
                                xlab("Periodo")
                if(nrow(seguros)<15)
                    {
                        p1 <- p1 + geom_point()
                    }
                ggsave("premios_sinistros_seg_total.jpeg",
                       plot=p1,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol,
                       units="cm"
                       )
                insertImage(workbook,
                            "premios_sinistros_seg",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "premios_sinistros_seg_total.jpeg"),
                            width=28,
                            height=5*numCol,
                            units="cm",
                            startRow=insert_line
                            )

            }


    #Plotting selected companies
        else
            {
                insert_line<- 1
                coenti_1 <- as.numeric(strsplit(params_1[[4]][2],",")[[1]])

            #Getting companies names
                cias <- read.csv2(paste0("data",slash,"Ses_cias.csv"),
                                  encoding="latin1")
            #Removing old data
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                    pattern="^premios_sinistros_seg_",
                                    full.names=TRUE))

            #Making plots
                plots<- list()
                numCol <- ceiling((ncol(seguros)-1)/3)

                for(j in 1:length(coenti_1))
                    {
                        seguros_sub <- seguros[seguros$coenti==coenti_1[j],-2]
                        if(!params_1[[6]][1])
                            {
                                seguros_sub <- seguros_sub[1:(nrow(seguros_sub)-1),]
                            }

                        cia <- as.character(cias[cias$Coenti==coenti_1[j],2])
                        cia <- stri_trans_general(cia,"Latin-ASCII")
                        writeData(workbook,
                                  "premios_sinistros_seg",
                                  cia,
                                  startRow=insert_line)
                        insert_line <- insert_line+1

                        seguros_sub_t <- seguros_sub[,-1]
                        rownames(seguros_sub_t) <- seguros_sub[,1]
                        seguros_sub_t <- as.data.frame(t(seguros_sub_t))
                        writeDataTable(workbook,
                                       "premios_sinistros_seg",
                                       format(
                                        seguros_sub_t,
                                        decimal.mark=","),
                                       rowNames=TRUE,
                                       startRow=insert_line)
                        insert_line <- insert_line + nrow(seguros_sub_t) +1

                        writeData(workbook,
                                  "premios_sinistros_seg",
                                  paste0("Ver Tabela ",j),
                                  startRow=insert_line)
                        insert_line <- insert_line + 2

                        plot_data <- melt(seguros_sub,id="yearsec")
                        period <- c("Mensal",
                                           "Trimestral",
                                           "Semestral",
                                           "Anual")[params_1[[6]]]

                        x_index <- rep(F,nrow(seguros_sub))
                        x_length <- length(x_index)
                        if(ceiling(x_length/3)!=x_length/3)
                            {
                                first_third <- ceiling(x_length/3)
                            }
                        else
                            {
                                first_third <- (x_length/3)+1
                            }
                        if(floor(x_length*2/3)!=x_length*2/3)
                            {
                                second_third <- floor(x_length*2/3)+1
                            }
                        else
                            {
                                second_third <- x_length*2/3
                            }
                        x_index[c(1,
                                  first_third,
                                  second_third,
                                  length(x_index))
                                ]<-T
                        p1 <- ggplot(data=plot_data,
                                     aes(x=yearsec,
                                         y=value,
                                         colour=variable,
                                         group=variable))+
                                        geom_line()+                                
                                        scale_x_discrete(
                                            breaks=plot_data$yearsec[x_index]
                                            )+
                                        facet_wrap(~variable, scales="free_y",ncol=3)+
                                        ggtitle(paste0(
                                            "Tabela ",
                                            j,
                                            ": ",
                                            cia,
                                            "\nPeriodicidade ",
                                            period,
                                            " \nDe ",
                                            seguros_sub[1,1],
                                            " a ",
                                            seguros_sub[nrow(seguros_sub),1])
                                        )+
                                        ylab("Valor")+
                                        xlab("Periodo")
                        if(nrow(seguros_sub)<15)
                            {
                                p1 <- p1 + geom_point()
                            }
                        ggsave(paste0("premios_sinistros_seg_",j,".jpeg"),
                               plot=p1,
                               path=paste0("proc_data",slash,"plots"),
                               width=28,
                               height=5*numCol,
                               units="cm"
                               )

                        plots[[j]]<- p1

                    }

            #Merging plots and printing
                plot_total <- arrangeGrob(grobs=plots,ncol=1)
                ggsave("premios_sinistros_seg_total.jpeg",
                       plot=plot_total,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol*length(plots),
                       limitsize=FALSE,
                       units="cm"
                       )

                insertImage(workbook,
                            "premios_sinistros_seg",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "premios_sinistros_seg_total.jpeg"
                                   ),
                            width=28,
                            height=5*numCol*length(plots),
                            units="cm",
                            startRow=insert_line
                            )
            }

    }

####################
#    Analysis 2    #
####################

if(params_2[[1]]==TRUE)
    {
        lista_cmpid <- list(
                            c(12032,6475,6807), 
                            c(11278,11275), 
                            c(6808,6476), 
                            c(6809,12028,12078,12340,12349,12364,6477,12204,12210,12246,12251), 
                            c(6810,6478), 
                            c(6844,7202,7139, 12079), 
                            c(7231,6816,7147,6481,6811,1613, 12242, 12181, 12269, 12270, 12271, 12232, 13111, 12351, 12375, 12376, 12379), 
                            c(6817,6470,6812,1612,4034, 12273, 12235, 12025, 12194, 12195, 12202, 12203, 12239, 12240, 12241), 
                            c(6818,6582,6813,4035), 
                            c(6819,6583,6814,4036), 
                            c(11247,11271,5988,6001,6014,6026,12689,12223,12261,1779, 12689, 12261, 12207, 12248, 12223), 
                            c(6198,6252,6820,6815,3235), 
                            c(11267,11322,7267))

        lista_labels <- list(
                             c("Premios_de_Resseguro"), 
                             c("Premios_de_Resseguro_RVNE"), 
                             c("Operacoes_com_Exterior"), 
                             c("Retrocessoes_Aceitas"), 
                             c("Premios_Cedidos_de_Retrocessao"), 
                             c("Variacoes_das_Provisoes_Tecnicas_de_Premios"), 
                             c("Sinistros"), 
                             c("Recuperacao_de_Sinistros"), 
                             c("Salvados"), 
                             c("Ressarcimentos"), 
                             c("Salvados_e_Ressarcimentos"),  
             c("Variacoes_das_Provisoes_de_Sinistros_Ocorridos_mas_nao_Avisados"), 
                             c("Variacoes_das_Provisoes_IBNER"))

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
        #print(filter_table)

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
				if(nrow(mov_grupos) > 0)
					{
						mov_grupos<-aggregate(mov_grupos[,5],
										by=list(
												YEARSEC=paste0(
						#Year
						format(as.Date(paste0(as.character(mov_grupos$DAMESANO),"01"),format="%Y%m%d"),"%Y"),
						#Section
                        str_pad(
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
                            ,2,pad="0")
													),
												COENTI=mov_grupos$COENTI,
												CMPID=mov_grupos$CMPID,
												GRACODIGO=mov_grupos$GRACODIGO
												),FUN=sum,na.rm=TRUE)

						colnames(mov_grupos)[5]<-"VALOR"
					}
					
				if(nrow(mov_ramos) > 0)
					{
						mov_ramos<-aggregate(mov_ramos[,6],
										by=list(
												yearsec=paste0(
						#Year
						format(as.Date(paste0(as.character(mov_ramos$damesano),"01"),format="%Y%m%d"),"%Y"),
						#Section
                        str_pad(
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
                            ,2,pad="0")
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
            }
        #Aggregate by Year/Trimester
        else if(params_2[[7]][2] == TRUE)
            {
                if(nrow(mov_grupos) > 0)
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
					}
					
				if(nrow(mov_ramos) > 0)
					{
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
            }
        #Aggregate by Year/Semester
        else if(params_2[[7]][3] == TRUE)
            {
                if(nrow(mov_grupos) > 0)
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
					}
					
				if(nrow(mov_ramos) > 0)
					{
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
            }
        #Aggregate by Year
        else if(params_2[[7]][4] == TRUE)
            {
                if(nrow(mov_grupos) > 0)
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
					}
					
				if(nrow(mov_ramos) > 0)
					{
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
            }

    #Post-Processing

        mov_grupos <- mov_grupos[,-4]
        mov_ramos <- mov_ramos[,c(-4,-5,-7,-8)]

        colnames(mov_grupos) <- tolower(colnames(mov_grupos))

        mov_grupos <- rbind(mov_grupos,mov_ramos)

        rep_cmpids <- rep(seq_along(lista_cmpid),sapply(lista_cmpid,length))
        mov_grupos[,3] <- rep_cmpids[
                                match(mov_grupos[,3],unlist(lista_cmpid))]

        #print(head(mov_grupos))
        #print(head(mov_ramos))

        grupos <- read.csv2(paste0("data",slash,"ses_gruposramos.csv"),
                            encoding="latin1")
        grupos <- grupos[order(grupos[,3]),]

        mov_grupos <- aggregate(mov_grupos[,4],
                                by=list(
                                     yearsec=mov_grupos$yearsec,
                                     coenti=mov_grupos$coenti,
                                     cmpid=mov_grupos$cmpid
                                     )
                             ,FUN=sum,na.rm=TRUE)
        mov_grupos_melt <- mov_grupos
        mov_grupos <- dcast(data=mov_grupos,
                              formula= yearsec + coenti ~ cmpid,
                              fun.aggregate=sum,
                              value.var="x")
        mov_grupos_names <- colnames(mov_grupos)[3:ncol(mov_grupos)]
        #grupos_sel <- stri_trans_general(as.character(
        #                 grupos[grupos$GRACODIGO %in% mov_grupos_names,2]),
        #                "Latin-ASCII")
        topic_sel <- lista_labels[
                        as.numeric(colnames(mov_grupos)[3:ncol(mov_grupos)])]
        #print(head(mov_grupos))
        #grupos_sel <- unlist(strsplit(grupos_sel,"- "))
        #grupos_sel <- grupos_sel[c(2*(1:(length(grupos_sel)/2)))]
        colnames(mov_grupos) <- c(colnames(mov_grupos)[1:2],topic_sel)
        write.csv2(mov_grupos,paste0("proc_data",slash,"resseg.csv"))

        #mov_ramos <- aggregate(mov_ramos[,4],
        #                        by=list(
        #                             yearsec=mov_ramos$yearsec,
        #                             coenti=mov_ramos$coenti,
        #                             #ramcodigo=mov_ramos$ramcodigo,
        #                             cmpid=mov_ramos$cmpid
        #                             )
        #                     ,FUN=sum,na.rm=TRUE)
        #mov_ramos_melt <- mov_ramos
        #mov_ramos <- dcast(data=mov_ramos,
        #                      formula= yearsec + coenti ~ cmpid,
        #                      fun.aggregate=sum,
        #                      value.var="x")
        #mov_ramos_names <- colnames(mov_ramos)[3:ncol(mov_ramos)]
        #ramos_sel <- stri_trans_general(as.character(
        #                 grupos[grupos$GRACODIGO %in% mov_ramos_names,2]),
        #                "Latin-ASCII")
        #print(head(mov_ramos))
        #ramos_sel <- unlist(strsplit(ramos_sel,"- "))
        #ramos_sel <- ramos_sel[c(2*(1:(length(ramos_sel)/2)))]
        #colnames(mov_ramos) <- c(colnames(mov_ramos)[1:2],ramos_sel)
        #write.csv2(mov_ramos,paste0("proc_data",slash,"resseg_ramos.csv"))

        addWorksheet(workbook,"resseguros")
        #addWorksheet(workbook,"resseg_ramos")

        insert_line <- 1

    #Merging All Companies
        if(params_2[[4]][1]=="TRUE")
            {

                mov_grupos <- aggregate(mov_grupos[,3:ncol(mov_grupos)],
                                by=list(yearsec=mov_grupos$yearsec),
                                FUN=sum,na.rm=TRUE)

                #mov_ramos <- aggregate(mov_ramos[,3:ncol(mov_ramos)],
                #                by=list(yearsec=mov_ramos$yearsec),
                #                FUN=sum,na.rm=TRUE)

                writeData(workbook,
                          "resseguros",
                          "Todas as Empresas",
                          startRow=insert_line)

                #writeData(workbook,
                #          "resseg_ramos",
                #          "Todas as Empresas",
                #          startRow=insert_line)

                insert_line <- insert_line + 1
                numCol_grupos <- ceiling((ncol(mov_grupos)-1)/3)
                #numCol_ramos <- ceiling((ncol(mov_ramos)-1)/3)

            #Transpose Data Frame
                mov_grupos_t <- mov_grupos[,-1]
                #mov_ramos_t <- mov_ramos[,-1]
                rownames(mov_grupos_t) <- mov_grupos[,1]
                #rownames(mov_ramos_t) <- mov_ramos[,1]
                mov_grupos_t <- as.data.frame(t(mov_grupos_t))
                #mov_ramos_t <- as.data.frame(t(mov_ramos_t))

                writeDataTable(workbook,
                               "resseguros",
                               format(
                                mov_grupos_t,
                                decimal.mark=","),
                               rowNames=TRUE,
                               startRow=insert_line
                               )

                #writeDataTable(workbook,
                #               "resseg_ramos",
                #               format(
                #                mov_ramos_t,
                #                decimal.mark=","),
                #               rowNames=TRUE,
                #               startRow=insert_line
                #               )

                insert_line_grupos <- insert_line+nrow(mov_grupos_t)+1
                #insert_line_ramos <- insert_line+nrow(mov_ramos_t)+1

                writeData(workbook,
                          "resseguros",
                          "Ver Tabela 1",
                          startRow=insert_line_grupos)
                #writeData(workbook,
                #          "resseg_ramos",
                #          "Ver Tabela 1",
                #          startRow=insert_line_ramos)

                insert_line_grupos <- insert_line_grupos + 2
                #insert_line_ramos <- insert_line_ramos + 2

            #Removing old files
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^ressuros_",
                                 full.names=TRUE))
                #file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                #                 pattern="^resseg_ramos_",
                #                 full.names=TRUE))
            #Plotting data
                period <- c("Mensal",
                                   "Trimestral",
                                   "Semestral",
                                   "Anual")[params_2[[7]]]

                plot_data_grupos <- melt(mov_grupos,id="yearsec")
                #plot_data_ramos <- melt(mov_ramos,id="yearsec")

                #Plot Indexes
                x_index_grupos <- rep(F,nrow(mov_grupos))
                x_length_grupos <- length(x_index_grupos)

                #x_index_ramos <- rep(F,nrow(mov_ramos))
                #x_length_ramos <- length(x_index_ramos)

                if(ceiling(x_length_grupos/3)!=x_length_grupos/3)
                    {
                        first_third_grupos <- ceiling(x_length_grupos/3)
                        second_third_grupos <- floor(x_length_grupos*2/3)+1
                    }
                else
                    {
                        first_third_grupos <- (x_length_grupos/3)+1
                        second_third_grupos <- x_length_grupos*2/3
                    }

                #if(ceiling(x_length_ramos/3)!=x_length_ramos/3)
                #    {
                #        first_third_ramos <- ceiling(x_length_ramos/3)
                #        second_third_ramos <- floor(x_length_ramos*2/3)+1
                #    }
                #else
                #    {
                #        first_third_ramos <- (x_length_ramos/3)+1
                #        second_third_ramos <- x_length_ramos*2/3
                #    }

                x_index_grupos[c(1,
                        first_third_grupos,
                        second_third_grupos,
                        length(x_index_grupos))
                        ]<-T

                #x_index_ramos[c(1,
                #        first_third_ramos,
                #        second_third_ramos,
                #        length(x_index_ramos))
                #        ]<-T

                p1_grupos <- ggplot(data=plot_data_grupos,
                             aes(x=yearsec,y=value,colour=variable,group=variable))+
                                geom_line()+
                                scale_x_discrete(
                                    breaks=plot_data_grupos$yearsec[x_index_grupos]
                                    )+
                                facet_wrap(~variable, scales="free_y",ncol=3)+
                                ggtitle(paste0(
                                    "Tabela 1: Todas as Empresas\nPeriodicidade ",
                                    period,
                                    " \nDe ",
                                    mov_grupos[1,1],
                                    " a ",
                                    mov_grupos[nrow(mov_grupos),1])
                                )+
                                ylab("Valor")+
                                xlab("Periodo")
                if(nrow(mov_grupos)<15)
                    {
                        p1_grupos <- p1_grupos + geom_point()
                    }

                #p1_ramos <- ggplot(data=plot_data_ramos,
                #             aes(x=yearsec,y=value,colour=variable,group=variable))+
                #                geom_line()+
                #                scale_x_discrete(
                #                    breaks=plot_data_ramos$yearsec[x_index_ramos]
                #                    )+
                #                facet_wrap(~variable, scales="free_y",ncol=3)+
                #                ggtitle(paste0(
                #                    "Tabela 1: Todas as Empresas\nPeriodicidade ",
                #                    period,
                #                    " \nDe ",
                #                    mov_ramos[1,1],
                #                    " a ",
                #                    mov_ramos[nrow(mov_ramos),1])
                #                )+
                #                ylab("Valor")+
                #                xlab("Periodo")
                #if(nrow(mov_ramos)<15)
                #    {
                #        p1_ramos <- p1_ramos + geom_point()
                #    }

                ggsave("resseguros_total.jpeg",
                       plot=p1_grupos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_grupos,
                       units="cm"
                       )
                #ggsave("resseg_ramos_total.jpeg",
                #       plot=p1_ramos,
                #       path=paste0("proc_data",slash,"plots"),
                #       width=28,
                #       height=5*numCol_ramos,
                #       units="cm"
                #       )

                insertImage(workbook,
                            "resseguros",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "resseguros_total.jpeg"),
                            width=28,
                            height=5*numCol_grupos,
                            units="cm",
                            startRow=insert_line_grupos
                            )
                #insertImage(workbook,
                #            "resseg_ramos",
                #            paste0("proc_data",
                #                   slash,
                #                   "plots",
                #                   slash,
                #                   "resseg_ramos_total.jpeg"),
                #            width=28,
                #            height=5*numCol_ramos,
                #            units="cm",
                #            startRow=insert_line_ramos
                #            )


            }


    #Plotting selected companies
        else
            {
                insert_line_grupos <- 1
                #insert_line_ramos <- 1
                coenti_2 <- as.numeric(strsplit(params_2[[4]][2],",")[[1]])

            #Getting companies names
                cias <- read.csv2(paste0("data",slash,"Ses_cias.csv"),
                                  encoding="latin1")
            #Removing old data
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^resseguros_",
                                 full.names=TRUE))
                #file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                #                 pattern="^resseg_ramos_",
                #                 full.names=TRUE))

            #Making plots
                plots_grupos<- list()
                #plots_ramos<- list()

                tabela_grupos <- 1
                #tabela_ramos <- 1

                for(j in 1:length(coenti_2))
                    {
                        mov_grupos_sub <- mov_grupos[
                                        mov_grupos$coenti==coenti_2[j],-2]

                        if(nrow(mov_grupos_sub)==0)
                            {
                                next
                            }

                        numCol_grupos <- ceiling((ncol(mov_grupos_sub)-1)/3)

                        if(!params_2[[7]][1])
                            {
                                mov_grupos_sub <- mov_grupos_sub[1:
                                            (nrow(mov_grupos_sub)-1),]
                            }

                        cia <- as.character(cias[cias$Coenti==coenti_2[j],2])
                        cia <- stri_trans_general(cia,"Latin-ASCII")

                        writeData(workbook,
                                  "resseguros",
                                  cia,
                                  startRow=insert_line_grupos)

                        insert_line_grupos <- insert_line_grupos+1

                    #Transpose Data Frame
                        mov_grupos_sub_t <- mov_grupos_sub[,-1]
                        rownames(mov_grupos_sub_t) <- mov_grupos_sub[,1]
                        mov_grupos_sub_t <- as.data.frame(t(mov_grupos_sub_t))

                        writeDataTable(workbook,
                                       "resseguros",
                                       format(
                                        mov_grupos_sub_t,
                                        decimal.mark=","),
                                       rowNames=TRUE,
                                       startRow=insert_line_grupos
                                       )

                        insert_line_grupos <- insert_line_grupos +
                                                nrow(mov_grupos_sub_t)+1

                        writeData(workbook,
                                  "resseguros",
                                  paste0("Ver tabela ",tabela_grupos),
                                  startRow=insert_line_grupos)

                        insert_line_grupos <- insert_line_grupos+2

                    #Plotting data
                        period <- c("Mensal",
                                           "Trimestral",
                                           "Semestral",
                                           "Anual")[params_2[[7]]]

                        plot_data_grupos <- melt(mov_grupos_sub,id="yearsec")

                        #Plot Indexes
                        x_index_grupos <- rep(F,nrow(mov_grupos_sub))
                        x_length_grupos <- length(x_index_grupos)


                        if(ceiling(x_length_grupos/3)!=x_length_grupos/3)
                            {
                                first_third_grupos <- ceiling(x_length_grupos/3)
                                second_third_grupos <- floor(x_length_grupos*2/3)+1
                            }
                        else
                            {
                                first_third_grupos <- (x_length_grupos/3)+1
                                second_third_grupos <- x_length_grupos*2/3
                            }

                        x_index_grupos[c(1,
                                first_third_grupos,
                                second_third_grupos,
                                length(x_index_grupos))
                                ]<-T

                        p1_grupos <- ggplot(data=plot_data_grupos,
                                     aes(x=yearsec,
                                         y=value,
                                         colour=variable,
                                         group=variable))+
                                        geom_line()+
                                        scale_x_discrete(
                                            breaks=plot_data_grupos$yearsec[
                                                                x_index_grupos])+
                                        facet_wrap(~variable, 
                                                   scales="free_y",
                                                   ncol=3)+
                                        ggtitle(paste0(
                                            "Tabela ",
                                            tabela_grupos,
                                            ": ",
                                            cia,
                                            " \nPeriodicidade ",
                                            period,
                                            " \nDe ",
                                            mov_grupos_sub[1,1],
                                            " a ",
                                            mov_grupos_sub[nrow(mov_grupos_sub),1])
                                        )+
                                        ylab("Valor")+
                                        xlab("Periodo")
                        if(nrow(mov_grupos_sub)<15)
                            {
                                p1_grupos <- p1_grupos + geom_point()
                            }

                        ggsave(paste0("resseguros_",tabela_grupos,".jpeg"),
                               plot=p1_grupos,
                               path=paste0("proc_data",slash,"plots"),
                               width=28,
                               height=5*numCol_grupos,
                               units="cm"
                               )

                        plots_grupos[[tabela_grupos]]<- p1_grupos

                        tabela_grupos <- tabela_grupos + 1

                    }

            #Merging plots and printing
                plot_total_grupos <- arrangeGrob(grobs=plots_grupos,ncol=1)
                ggsave("resseguros_total.jpeg",
                       plot=plot_total_grupos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_grupos*length(plots_grupos),
                       limitsize=FALSE,
                       units="cm"
                       )

                insertImage(workbook,
                            "resseguros",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "resseguros_total.jpeg"
                                   ),
                            width=28,
                            height=5*numCol_grupos*length(plots_grupos),
                            units="cm",
                            startRow=insert_line_grupos
                            )

                #for(j in 1:length(coenti_2))
                #    {
                #        mov_ramos_sub <- mov_ramos[
                #                        mov_ramos$coenti==coenti_2[j],-2]

                #        if(nrow(mov_ramos_sub)==0)
                #            {
                #                next
                #            }

                #        numCol_ramos <- ceiling((ncol(mov_ramos_sub)-1)/3)

                #        if(!params_2[[7]][1])
                #            {
                #                mov_ramos_sub <- mov_ramos_sub[1:
                #                            (nrow(mov_ramos_sub)-1),]
                #            }

                #        cia <- as.character(cias[cias$Coenti==coenti_2[j],2])
                #        cia <- stri_trans_general(cia,"Latin-ASCII")

                #        writeData(workbook,
                #                  "resseg_ramos",
                #                  cia,
                #                  startRow=insert_line_ramos)

                #        insert_line_ramos <- insert_line_ramos+1

                #    #Transpose Data Frame
                #        mov_ramos_sub_t <- mov_ramos_sub[,-1]
                #        rownames(mov_ramos_sub_t) <- mov_ramos_sub[,1]
                #        mov_ramos_sub_t <- as.data.frame(t(mov_ramos_sub_t))

                #        writeDataTable(workbook,
                #                       "resseg_ramos",
                #                       format(
                #                        mov_ramos_sub_t,
                #                        decimal.mark=","),
                #                       rowNames=TRUE,
                #                       startRow=insert_line_ramos
                #                       )

                #        insert_line_ramos <- insert_line_ramos + 
                #                                nrow(mov_ramos_sub_t)+1

                #        writeData(workbook,
                #                  "resseg_ramos",
                #                  paste0("Ver tabela ",tabela_ramos),
                #                  startRow=insert_line_ramos)

                #        insert_line_ramos <- insert_line_ramos+2

                #    #Plotting data
                #        period <- c("Mensal",
                #                           "Trimestral",
                #                           "Semestral",
                #                           "Anual")[params_2[[7]]]

                #        plot_data_ramos <- melt(mov_ramos_sub,id="yearsec")

                #        #Plot Indexes
                #        x_index_ramos <- rep(F,nrow(mov_ramos_sub))
                #        x_length_ramos <- length(x_index_ramos)

                #        if(ceiling(x_length_ramos/3)!=x_length_ramos/3)
                #            {
                #                first_third_ramos <- ceiling(x_length_ramos/3)
                #                second_third_ramos <- floor(x_length_ramos*2/3)+1
                #            }
                #        else
                #            {
                #                first_third_ramos <- (x_length_ramos/3)+1
                #                second_third_ramos <- x_length_ramos*2/3
                #            }
        
                #        x_index_ramos[c(1,
                #                first_third_ramos,
                #                second_third_ramos,
                #                length(x_index_ramos))
                #                ]<-T

                #        p1_ramos <- ggplot(data=plot_data_ramos,
                #                     aes(x=yearsec,
                #                         y=value,
                #                         colour=variable,
                #                         group=variable))+
                #                        geom_line()+
                #                        scale_x_discrete(
                #                            breaks=plot_data_ramos$yearsec[
                #                                                x_index_ramos])+
                #                        facet_wrap(~variable, 
                #                                   scales="free_y",
                #                                   ncol=3)+
                #                        ggtitle(paste0(
                #                            "Tabela ",
                #                            tabela_ramos,
                #                            ": ",
                #                            cia,
                #                            " \nPeriodicidade ",
                #                            period,
                #                            " \nDe ",
                #                            mov_ramos_sub[1,1],
                #                            " a ",
                #                            mov_ramos_sub[nrow(mov_ramos_sub),1])
                #                        )+
                #                        ylab("Valor")+
                #                        xlab("Periodo")
                #        if(nrow(mov_ramos_sub)<15)
                #            {
                #                p1_ramos <- p1_ramos + geom_point()
                #            }

                #        ggsave(paste0("resseg_ramos_",tabela_ramos,".jpeg"),
                #               plot=p1_ramos,
                #               path=paste0("proc_data",slash,"plots"),
                #               width=28,
                #               height=5*numCol_ramos,
                #               units="cm"
                #               )

                #        plots_ramos[[tabela_ramos]]<- p1_ramos

                #        tabela_ramos <- tabela_ramos + 1

                #    }

            #Mer#ging plots and printing
                #plot_total_ramos <- arrangeGrob(grobs=plots_ramos,ncol=1)
                #ggsave("resseg_ramos_total.jpeg",
                #       plot=plot_total_ramos,
                #       path=paste0("proc_data",slash,"plots"),
                #       width=28,
                #       height=5*numCol_ramos*length(plots_ramos),
                #       limitsize=FALSE,
                #       units="cm"
                #       )

                #insertImage(workbook,
                #            "resseg_ramos",
                #            paste0("proc_data",
                #                   slash,
                #                   "plots",
                #                   slash,
                #                   "resseg_ramos_total.jpeg"
                #                   ),
                #            width=28,
                #            height=5*numCol_ramos*length(plots_ramos),
                #            units="cm",
                #            startRow=insert_line_ramos
                #            )
            }


    }

####################
#    Analysis 3    #
####################

if(params_3[[1]]==TRUE)
    {
    #Loading and cleaning Ses_campos.csv

        campos <- read.csv2(paste0("data",slash,"Ses_campos.csv"),encoding="latin1",strip.white=TRUE)
        campos$noitem <- sapply(campos$noitem, function(x) stri_trans_general(x,"Latin-ASCII"))
        campos$noitem <- sapply(campos$noitem, function(x) stri_trim(x))

    #Getting campos order

        #ativos <- read.csv2("ativos.csv",header=FALSE,encoding="latin1",strip.white=TRUE)
        #ativos <- sapply(ativos[,1],function(x) stri_trans_general(x,"Latin-ASCII"))
        #ativos <- make.unique(ativos,sep="_")
        #ativos <- data.frame(noitem=ativos)
        #campos_ativos <- campos[campos$nuquad=="22A",1:2]
        #campos_ativos$noitem <- make.unique(campos_ativos$noitem,sep="_")
        #ativos_order <- merge(ativos,campos_ativos,by="noitem",sort=FALSE)
        ##print(ativos_order$noitem)
        ##print(as.numeric(strsplit(,split=", ")[[1]]))

        #passivos <- read.csv2("passivos.csv",header=FALSE,encoding="latin1",strip.white=TRUE)
        #passivos <- sapply(passivos[,1],function(x) stri_trans_general(x,"Latin-ASCII"))
        #passivos <- make.unique(passivos,sep="_")
        #passivos <- data.frame(noitem=passivos)
        #campos_passivos <- campos[campos$nuquad=="22P",1:2]
        #campos_passivos$noitem <- make.unique(campos_passivos$noitem,sep="_")
        #passivos_order <- merge(passivos,campos_passivos,by="noitem",sort=FALSE)
        ##print(passivos_order$noitem)
        ##print(toString(passivos_order[,2]))

        #dre <- read.csv2("dre.csv",header=FALSE,encoding="latin1",strip.white=TRUE)
        #dre <- sapply(dre[,1],function(x) stri_trans_general(x,"Latin-ASCII"))
        #dre <- make.unique(dre,sep="_")
        #dre <- data.frame(noitem=dre)
        #campos_dre <- campos[campos$nuquad=="23",1:2]
        #campos_dre$noitem <- make.unique(campos_dre$noitem,sep="_")
        #dre_order <- merge(dre,campos_dre,by="noitem",sort=FALSE)
        #print(dre_order$noitem)
        #print(toString(dre_order[,2]))

    #Loading SES_Balanco.csv

        balanco <- read.csv2(paste0(
            "data",slash,"SES_Balanco.csv"),encoding="latin1",strip.white=TRUE)

    #Subsetting by Initial Date

        if(nchar(params_3[[3]][1])==10)
            {
                init_date_3 <- as.Date(params_3[[3]][1], format="%d/ %m/ %Y")
                if(!is.na(init_date_3))
                    {
                        init_date_3 <- format(init_date_3,"%Y%m")
                        if(nchar(init_date_3)==6)
                            {
                                balanco <- balanco[
                                    balanco$damesano >= as.numeric(init_date_3),]
                            }
                    }
            }

    #Subsetting by Final Date

        if(nchar(params_3[[3]][2])==10)
            {
                final_date_3 <- as.Date(params_3[[3]][2], format="%d/ %m/ %Y")
                if(!is.na(final_date_3))
                    {
                        final_date_3 <- format(final_date_3,"%Y%m")
                        if(nchar(final_date_3)==6)
                            {
                                balanco <- balanco[
                                    balanco$damesano <= as.numeric(final_date_3),]
                            }
                    }
            }

    #Subsetting by Company
        if(params_3[[4]][1]=="FALSE")
            {
                coenti_3 <- as.numeric(strsplit(params_3[[4]][2],",")[[1]])
                balanco <- balanco[balanco$coenti %in% coenti_3,]
            }

    #Aggreggating by time interval
       
        balanco <- balanco[,c(2,1,3,6,4,5)]

        #Aggregate by Year/Month
        if(params_3[[5]][1] == TRUE)
            {
				if(nrow(balanco) > 0)
					{
						balanco<-aggregate(balanco[,5],
										by=list(
												yearsec=paste0(
						#Year
						format(as.Date(paste0(as.character(balanco$damesano),"01"),format="%Y%m%d"),"%Y"),
						#Section
                        str_pad(
						as.character(
						ceiling(
						as.numeric(
						format(
						as.Date(
						paste0(
						as.character(balanco$damesano)
							,"01")
							,format="%Y%m%d")
							,"%m")
							)/1
							)
							)
                            ,2,pad="0")
													),
                                                cmpid=balanco$cmpid,
												coenti=balanco$coenti,
												quadro=balanco$quadro
												),FUN=sum,na.rm=TRUE)

						colnames(balanco)[5]<-"valor"
					}

            }
        #Aggregate by Year/Trimester
        else if(params_3[[5]][2] == TRUE)
            {
				if(nrow(balanco) > 0)
					{
						balanco<-aggregate(balanco[,5],
										by=list(
												yearsec=paste0(
						#Year
						format(as.Date(paste0(as.character(balanco$damesano),"01"),format="%Y%m%d"),"%Y"),
						#Section
						paste0("0",
						as.character(
						ceiling(
						as.numeric(
						format(
						as.Date(
						paste0(
						as.character(balanco$damesano)
							,"01")
							,format="%Y%m%d")
							,"%m")
							)/3
							)
							)
							)
													),
                                                cmpid=balanco$cmpid,
												coenti=balanco$coenti,
												quadro=balanco$quadro
												),FUN=sum,na.rm=TRUE)

						colnames(balanco)[5]<-"valor"
					}

            }
        #Aggregate by Year/Semester
        else if(params_3[[5]][3] == TRUE)
            {
				if(nrow(balanco) > 0)
					{
						balanco<-aggregate(balanco[,5],
										by=list(
												yearsec=paste0(
						#Year
						format(as.Date(paste0(as.character(balanco$damesano),"01"),format="%Y%m%d"),"%Y"),
						#Section
						paste0("0",
						as.character(
						ceiling(
						as.numeric(
						format(
						as.Date(
						paste0(
						as.character(balanco$damesano)
							,"01")
							,format="%Y%m%d")
							,"%m")
							)/6
							)
							)
							)
													),
                                                cmpid=balanco$cmpid,
												coenti=balanco$coenti,
												quadro=balanco$quadro
												),FUN=sum,na.rm=TRUE)

						colnames(balanco)[5]<-"valor"
					}

            }
        #Aggregate by Year
        else if(params_3[[5]][4] == TRUE)
            {
				if(nrow(balanco) > 0)
					{

						balanco<-aggregate(balanco[,5],
										by=list(
												yearsec=paste0(
						#Year
						format(as.Date(paste0(as.character(balanco$damesano),"01"),format="%Y%m%d"),"%Y"),
						#Section
						paste0("0",
						as.character(
						ceiling(
						as.numeric(
						format(
						as.Date(
						paste0(
						as.character(balanco$damesano)
							,"01")
							,format="%Y%m%d")
							,"%m")
							)/12
							)
							)
							)
													),
                                                cmpid=balanco$cmpid,
												coenti=balanco$coenti,
												quadro=balanco$quadro
												),FUN=sum,na.rm=TRUE)

						colnames(balanco)[5]<-"valor"
                        #print(head(balanco))
                        #balanco <- balanco[,-4]
					}

            }

    #Filtering Quadros

        filtros <- c("22A","22P","23")
        filtros_sel <- filtros[params_3[[2]]]

        campos_list <- list()
        balanco_list <- list()

    #Subsetting by Quadros

        for(i in 1:length(filtros_sel))
            {
                campos_list[[i]] <- campos[campos$nuquad==filtros_sel[i],1:2]
                colnames(campos_list[[i]])[1]<-"cmpid"

                balanco_list[[i]] <- balanco[balanco$quadro==filtros_sel[i],]

                balanco_list[[i]] <- balanco_list[[i]][
                                balanco_list[[i]]$cmpid %in% campos_list[[i]]$cmpid,]
                balanco_list[[i]] <- merge(
                                    balanco_list[[i]],campos_list[[i]],by="cmpid",sort=FALSE)

                balanco_list[[i]] <- balanco_list[[i]][
                        order(balanco_list[[i]]$yearsec,balanco_list[[i]]$cmpid),]
                balanco_list[[i]] <- balanco_list[[i]][,c(1,2,6,3,5)]

            #Processing Data

                if(filtros_sel[i]=="22A")
                {
                    contab_name="ativos"
                    string_cmpid = "1479, 356, 1484, 11136, 1480, 6152, 1496, 5497, 11137, 345, 11273, 3162, 363, 1488, 11285, 11286, 11287, 1487, 11290, 7155, 11292, 359, 7154, 3165, 3166, 5498, 3172, 11138, 6155, 6156, 11139, 11142, 13124, 11300, 11301, 11302, 11144, 7223, 12515, 11146, 11147, 12516, 11284, 13125, 11140, 1491, 1489, 6789, 343, 6787, 6158, 6450, 11151, 11348, 336, 7225, 11141, 11153, 11154, 11155, 362, 5586, 5587, 361, 11149, 3170, 3171, 11152, 5503, 11158, 11159, 11160, 11982, 11983, 11984, 340, 11303, 11988, 11989, 353, 11307, 5917, 331, 332, 6165, 1038, 5500, 11164, 358, 11156, 3175, 3188, 3190, 7143, 7233, 7237, 3189, 7146, 13123, 3163, 11157, 3179, 3180, 5501, 3181, 11165, 3168, 6168, 11166, 11169, 13129, 11304, 11305, 11306, 11171, 7235, 12517, 11173, 11174, 12518, 11289, 13130, 11167, 1036, 1035, 6791, 11150, 6788, 6170, 6451, 11177, 11349, 350, 13128, 11168, 11179, 11180, 11181, 3192, 3193, 11178, 3185, 3186, 11183, 351, 11185, 11186, 11187, 11992, 11993, 11994, 11985, 1485, 11998, 11999, 11990, 11986, 323, 6452, 6453, 6454, 6455, 6466, 6467, 11192, 11193, 327, 11191, 11194, 11308, 11195, 11309, 11310, 1503, 1507, 1499, 6474, 5591, 11196, 11197, 5594, 7149, 5920, 7152, 7153, 11184, 1502, 7266, 1498, 5124, 5714, 5126, 5127, 11198, 5716, 5717, 5129"
                }
                else if(filtros_sel[i]=="22P")
                {
                    contab_name="passivos"
                    string_cmpid <- "1040, 3277, 5511, 3207, 5894, 1547, 5895, 5509, 6073, 3213, 3214, 3215, 5896, 5515, 6088, 3219, 5510, 5514, 6090, 11311, 6074, 6091, 11312, 11313, 1550, 3224, 11199, 3226, 11201, 11202, 11203, 3227, 3228, 11205, 11206, 11207, 11208, 6100, 3251, 3252, 3259, 11226, 3268, 11227, 11228, 11229, 11230, 5660, 5720, 7092, 5662, 5908, 3205, 6087, 5909, 3209, 5512, 3212, 3283, 3284, 5910, 1551, 3218, 5513, 6089, 3286, 3222, 3223, 3291, 11209, 11210, 11220, 11221, 6125, 3306, 3307, 3312, 11213, 3319, 11214, 11215, 11216, 11217, 6142, 5907, 7093, 6145, 1555, 6148, 3403, 1563, 1565, 6149, 6150, 7265, 1561, 5517, 6151, 1554, 3343, 3337, 3345, 6285, 3347, 5118, 5718, 5120, 5121, 11274, 5719, 6143, 5721, 5123"
                }
                else if(filtros_sel[i]=="23")
                {
                    contab_name="dre"
                    string_cmpid <- "6183, 7138, 11342, 6185, 6188, 11981, 6189, 6187, 6190, 7220, 7200, 7218, 7201, 7219, 7139, 7075, 12695, 12675, 7081, 12676, 12677, 4027, 11232, 7215, 7216, 6196, 11233, 11234, 11235, 11236, 11314, 6582, 6583, 7195, 12678, 12679, 7221, 7196, 6246, 7198, 11237, 11316, 11317, 11318, 11319, 11320, 11257, 11238, 11239, 11240, 11251, 11241, 11242, 12680, 12681, 11244, 11245, 11323, 11324, 11247, 13191, 7183, 7184, 6486, 6487, 7185, 6488, 6489, 7186, 6238, 6256, 6259, 6462, 7189, 12682, 11248, 7192, 7193, 7217, 5722, 11249, 11325, 11326, 11327, 11328, 11250, 11321, 11252, 11253, 11254, 11255, 6202, 4059, 6309, 6310, 11329, 11330, 4062, 12683, 12684, 12685, 11332, 11333, 11256, 11334, 11335, 11336, 11337, 6261, 11338, 11339, 6312, 6313, 4069, 6591, 6592, 6593, 6317, 6595, 6596, 6597, 6321, 4070, 6571, 6598, 6599, 4074, 6325, 6326, 6327, 6328, 7262, 7263, 6573, 11258, 4079, 4081, 4080, 4083, 518, 535, 517"
                }

                lista_ressegs <- "38741, 30074, 34819, 36099, 37052, 38253, 31623, 38873, 33294, 39764, 37729, 31551, 38270, 30201, 34665, 32875"

                #print(cmpid_order)
                print_line <- 1

                addWorksheet(workbook,paste0(contab_name,"_contabeis"))

                if(params_3[[4]][1]=="TRUE")
                {
                    writeData(workbook,paste0(contab_name,"_contabeis"),
                              "Todas as Empresas",startRow=print_line)

                    print_line <- print_line+1

                    balanco_list[[i]] <- aggregate(balanco_list[[i]][,5],by=list(
                                            yearsec=balanco_list[[i]]$yearsec,
                                            cmpid=balanco_list[[i]]$cmpid),
                                                   FUN=sum,na.rm=TRUE)

                    colnames(balanco_list[[i]])[3] <- "valor"


                    balanco_list_curr_ex <- dcast(data=balanco_list[[i]],
                              formula= cmpid ~ yearsec,
                              fun.aggregate=sum,
                              value.var="valor")

                    cmpid_order <- data.frame(
                        cmpid=as.numeric(strsplit(string_cmpid,split=", ")[[1]]))

                    balanco_list_curr_ex <- merge(cmpid_order,balanco_list_curr_ex,
                                               by="cmpid",sort=FALSE)


                    new_row_names_ex <- campos[
                            match(balanco_list_curr_ex[,1], campos$nuitem),2]
                    colnames(balanco_list_curr_ex)[1] <- "Campos"

                    balanco_list_curr_ex[,1] <- new_row_names_ex
                    #print(head(balanco_list_curr_ex))


                    writeDataTable(workbook,paste0(contab_name,"_contabeis"),
                                       format(balanco_list_curr_ex,decimal.mark=","),
                                       startRow=print_line)
                }
                else
                {
                    coenti_3 <- as.numeric(strsplit(params_3[[4]][2],",")[[1]])

                #Getting companies names
                    cias <- read.csv2(paste0("data",slash,"Ses_cias.csv"),
                                      encoding="latin1")

                    for(j in 1:length(coenti_3))
                        {
                            cia <- as.character(cias[cias$Coenti==coenti_3[j],2])
                            cia <- stri_trans_general(cia,"Latin-ASCII")


                            ressegs <- as.numeric(strsplit(lista_ressegs,", ")[[1]])
                            #print(ressegs)
                            #print(coenti_3[j])
                            if(filtros_sel[i]=="23" && coenti_3[j] %in% ressegs)
                                {
                                    string_cmpid <- "11259, 6475, 11340, 11341, 11342, 6476, 6477, 6564, 6565, 11232, 6481, 12686, 6582, 6583, 7238, 7240, 12687, 12688, 11237, 11334, 11344, 6568, 11262, 11239, 11240, 11247, 11253, 11267, 12690, 12691, 11244, 11269, 11346, 11324, 11270, 13191, 4069, 6591, 6592, 6593, 6594, 6595, 6596, 6597, 4070, 6571, 6598, 6599, 4074, 6600, 6601, 7260, 7261, 7262, 7263, 6573, 11258, 4079, 4081, 4080, 4083, 518, 535, 6581"
                                }
                            else if (filtros_sel[i]=="23" && !(coenti_3[j] %in%  ressegs))
                                {
                                    string_cmpid <- "6183, 7138, 11342, 6185, 6188, 11981, 6189, 6187, 6190, 7220, 7200, 7218, 7201, 7219, 7139, 7075, 12695, 12675, 7081, 12676, 12677, 4027, 11232, 7215, 7216, 6196, 11233, 11234, 11235, 11236, 11314, 6582, 6583, 7195, 12678, 12679, 7221, 7196, 6246, 7198, 11237, 11316, 11317, 11318, 11319, 11320, 11257, 11238, 11239, 11240, 11251, 11241, 11242, 12680, 12681, 11244, 11245, 11323, 11324, 11247, 13191, 7183, 7184, 6486, 6487, 7185, 6488, 6489, 7186, 6238, 6256, 6259, 6462, 7189, 12682, 11248, 7192, 7193, 7217, 5722, 11249, 11325, 11326, 11327, 11328, 11250, 11321, 11252, 11253, 11254, 11255, 6202, 4059, 6309, 6310, 11329, 11330, 4062, 12683, 12684, 12685, 11332, 11333, 11256, 11334, 11335, 11336, 11337, 6261, 11338, 11339, 6312, 6313, 4069, 6591, 6592, 6593, 6317, 6595, 6596, 6597, 6321, 4070, 6571, 6598, 6599, 4074, 6325, 6326, 6327, 6328, 7262, 7263, 6573, 11258, 4079, 4081, 4080, 4083, 518, 535, 517"
                                }

                            writeData(workbook,paste0(contab_name,"_contabeis"),
                              cia,startRow=print_line)

                            print_line <- print_line+1
                            
                            balanco_total <- balanco_list[[i]]
                            balanco_total <- balanco_total[
                                balanco_total$coenti %in% coenti_3[j],]

                            if(nrow(balanco_total)==0)
                                {
                                    next
                                }

                            balanco_total <- aggregate(balanco_total[,5],
                                                    by=list(
                                                    yearsec=balanco_total$yearsec,
                                                    cmpid=balanco_total$cmpid),
                                                           FUN=sum,na.rm=TRUE)

                            colnames(balanco_total)[3] <- "valor"
#
                            balanco_list_curr_ex <- dcast(data=balanco_list[[i]],
                                      formula= cmpid ~ yearsec,
                                      fun.aggregate=sum,
                                      value.var="valor")

                            cmpid_order <- data.frame(
                                cmpid=as.numeric(
                                    strsplit(string_cmpid,split=", ")[[1]]))

                            balanco_list_curr_ex <- merge(cmpid_order,
                                balanco_list_curr_ex,by="cmpid",sort=FALSE)


                            new_row_names_ex <- campos[
                                    match(balanco_list_curr_ex[,1], campos$nuitem),2]
                            colnames(balanco_list_curr_ex)[1] <- "Campos"

                            balanco_list_curr_ex[,1] <- new_row_names_ex
                            #balanco_list_curr <- dcast(data=balanco_total,
                            #          formula= yearsec ~ cmpid,
                            #          fun.aggregate=sum,
                            #          value.var="valor")

                            #balanco_list_curr_t <- balanco_list_curr[,-1]

                            #new_col_names <- balanco_list_curr[,1]

                            #new_row_names <- campos_list[[i]][
                            #            campos_list[[i]]$cmpid %in% colnames(
                            #            balanco_list_curr_t),2]

                            #balanco_list_curr_t <- as.data.frame(
                            #                        t(balanco_list_curr_t))

                            #colnames(balanco_list_curr_t) <- new_col_names
                            #balanco_list_curr_t <- cbind(new_row_names,
                            #                             balanco_list_curr_t)
                            #colnames(balanco_list_curr_t)[1] <- "Campos"

                            writeDataTable(workbook,paste0(contab_name,"_contabeis"),
                                        format(balanco_list_curr_ex,decimal.mark=","),
                                        startRow=print_line)

                            print_line <- print_line+nrow(balanco_list_curr_ex)+2

                        }
                }

            }

    }

saveWorkbook(workbook,paste0("SES-Susep-",Sys.Date(),".xlsx"),overwrite=TRUE)
