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
param_3_1 <- strsplit(substring(config_data[18],3,nchar(config_data[18])-2),"', '")[[1]]

#Param_3_2
param_3_2 <- strsplit(substring(config_data[19],2,nchar(config_data[19])-2),", '")[[1]]
param_3_2[1] <- param_3_2[1] == "True"

#Param_3_3
param_3_3 <- strsplit(substring(config_data[20],2,nchar(config_data[20])-2),", '")[[1]]
param_3_3[1] <- param_3_3[1] == "True"

#Param_3_4
param_3_4 <- strsplit(substring(config_data[21],2,nchar(config_data[21])-2),", '")[[1]]
param_2_4[1] <- param_3_4[1] == "True"

#Param_3_5
analysis_3_5_string <- strsplit(substring(config_data[22],2,nchar(config_data[22])-1), ", ")[[1]]
param_3_5 <- vector(length = length(analysis_3_5_string))
for (i in 1:length(analysis_3_5_string))
    {
        param_3_5[i] <- analysis_3_5_string[i] == "True"
    }

text_data <- config_data[23:length(config_data)]

params_3<- list(run_analysis_3,param_3_1,param_3_2,param_3_3,param_3_4,param_3_5,text_data)

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

        grupos <- read.csv2(paste0("data",slash,"ses_gruposramos.csv"),
                            encoding="latin1")
        grupos <- grupos[order(grupos[,3]),]

        mov_grupos <- aggregate(mov_grupos[,5],
                                by=list(
                                     yearsec=mov_grupos$YEARSEC,
                                     coenti=mov_grupos$COENTI,
                                     gracodigo=mov_grupos$GRACODIGO
                                     )
                             ,FUN=sum,na.rm=TRUE)
        mov_grupos_melt <- mov_grupos
        mov_grupos <- dcast(data=mov_grupos,
                              formula= yearsec + coenti ~ gracodigo,
                              fun.aggregate=sum,
                              value.var="x")
        mov_grupos_names <- colnames(mov_grupos)[3:ncol(mov_grupos)]
        grupos_sel <- stri_trans_general(as.character(
                         grupos[grupos$GRACODIGO %in% mov_grupos_names,2]),
                        "Latin-ASCII")
        grupos_sel <- unlist(strsplit(grupos_sel,"- "))
        grupos_sel <- grupos_sel[c(2*(1:(length(grupos_sel)/2)))]
        colnames(mov_grupos) <- c(colnames(mov_grupos)[1:2],grupos_sel)
        write.csv2(mov_grupos,paste0("proc_data",slash,"resseg_grupos.csv"))

        mov_ramos <- aggregate(mov_ramos[,6],
                                by=list(
                                     yearsec=mov_ramos$yearsec,
                                     coenti=mov_ramos$coenti,
                                     #ramcodigo=mov_ramos$ramcodigo,
                                     gracodigo=mov_ramos$gracodigo
                                     )
                             ,FUN=sum,na.rm=TRUE)
        mov_ramos_melt <- mov_ramos
        mov_ramos <- dcast(data=mov_ramos,
                              formula= yearsec + coenti ~ gracodigo,
                              fun.aggregate=sum,
                              value.var="x")
        mov_ramos_names <- colnames(mov_ramos)[3:ncol(mov_ramos)]
        ramos_sel <- stri_trans_general(as.character(
                         grupos[grupos$GRACODIGO %in% mov_ramos_names,2]),
                        "Latin-ASCII")
        ramos_sel <- unlist(strsplit(ramos_sel,"- "))
        ramos_sel <- ramos_sel[c(2*(1:(length(ramos_sel)/2)))]
        colnames(mov_ramos) <- c(colnames(mov_ramos)[1:2],ramos_sel)
        write.csv2(mov_ramos,paste0("proc_data",slash,"resseg_ramos.csv"))

        addWorksheet(workbook,"resseg_grupos")
        addWorksheet(workbook,"resseg_ramos")

        insert_line <- 1

    #Merging All Companies
        if(params_2[[4]][1]=="TRUE")
            {

                mov_grupos <- aggregate(mov_grupos[,3:ncol(mov_grupos)],
                                by=list(yearsec=mov_grupos$yearsec),
                                FUN=sum,na.rm=TRUE)

                mov_ramos <- aggregate(mov_ramos[,3:ncol(mov_ramos)],
                                by=list(yearsec=mov_ramos$yearsec),
                                FUN=sum,na.rm=TRUE)

                writeData(workbook,
                          "resseg_grupos",
                          "Todas as Empresas",
                          startRow=insert_line)

                writeData(workbook,
                          "resseg_ramos",
                          "Todas as Empresas",
                          startRow=insert_line)

                insert_line <- insert_line + 1
                numCol_grupos <- ceiling((ncol(mov_grupos)-1)/3)
                numCol_ramos <- ceiling((ncol(mov_ramos)-1)/3)

            #Transpose Data Frame
                mov_grupos_t <- mov_grupos[,-1]
                mov_ramos_t <- mov_ramos[,-1]
                rownames(mov_grupos_t) <- mov_grupos[,1]
                rownames(mov_ramos_t) <- mov_ramos[,1]
                mov_grupos_t <- as.data.frame(t(mov_grupos_t))
                mov_ramos_t <- as.data.frame(t(mov_ramos_t))

                writeDataTable(workbook,
                               "resseg_grupos",
                               format(
                                mov_grupos_t,
                                decimal.mark=","),
                               rowNames=TRUE,
                               startRow=insert_line
                               )

                writeDataTable(workbook,
                               "resseg_ramos",
                               format(
                                mov_ramos_t,
                                decimal.mark=","),
                               rowNames=TRUE,
                               startRow=insert_line
                               )

                insert_line_grupos <- insert_line+nrow(mov_grupos_t)+1
                insert_line_ramos <- insert_line+nrow(mov_ramos_t)+1

                writeData(workbook,
                          "resseg_grupos",
                          "Ver Tabela 1",
                          startRow=insert_line_grupos)
                writeData(workbook,
                          "resseg_ramos",
                          "Ver Tabela 1",
                          startRow=insert_line_ramos)

                insert_line_grupos <- insert_line_grupos + 2
                insert_line_ramos <- insert_line_ramos + 2

            #Removing old files
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^resseg_grupos_",
                                 full.names=TRUE))
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^resseg_ramos_",
                                 full.names=TRUE))
            #Plotting data
                period <- c("Mensal",
                                   "Trimestral",
                                   "Semestral",
                                   "Anual")[params_2[[7]]]

                plot_data_grupos <- melt(mov_grupos,id="yearsec")
                plot_data_ramos <- melt(mov_ramos,id="yearsec")

                #Plot Indexes
                x_index_grupos <- rep(F,nrow(mov_grupos))
                x_length_grupos <- length(x_index_grupos)

                x_index_ramos <- rep(F,nrow(mov_ramos))
                x_length_ramos <- length(x_index_ramos)

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

                if(ceiling(x_length_ramos/3)!=x_length_ramos/3)
                    {
                        first_third_ramos <- ceiling(x_length_ramos/3)
                        second_third_ramos <- floor(x_length_ramos*2/3)+1
                    }
                else
                    {
                        first_third_ramos <- (x_length_ramos/3)+1
                        second_third_ramos <- x_length_ramos*2/3
                    }

                x_index_grupos[c(1,
                        first_third_grupos,
                        second_third_grupos,
                        length(x_index_grupos))
                        ]<-T

                x_index_ramos[c(1,
                        first_third_ramos,
                        second_third_ramos,
                        length(x_index_ramos))
                        ]<-T

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

                p1_ramos <- ggplot(data=plot_data_ramos,
                             aes(x=yearsec,y=value,colour=variable,group=variable))+
                                geom_line()+
                                scale_x_discrete(
                                    breaks=plot_data_ramos$yearsec[x_index_ramos]
                                    )+
                                facet_wrap(~variable, scales="free_y",ncol=3)+
                                ggtitle(paste0(
                                    "Tabela 1: Todas as Empresas\nPeriodicidade ",
                                    period,
                                    " \nDe ",
                                    mov_ramos[1,1],
                                    " a ",
                                    mov_ramos[nrow(mov_ramos),1])
                                )+
                                ylab("Valor")+
                                xlab("Periodo")
                if(nrow(mov_ramos)<15)
                    {
                        p1_ramos <- p1_ramos + geom_point()
                    }

                ggsave("resseg_grupos_total.jpeg",
                       plot=p1_grupos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_grupos,
                       units="cm"
                       )
                ggsave("resseg_ramos_total.jpeg",
                       plot=p1_ramos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_ramos,
                       units="cm"
                       )

                insertImage(workbook,
                            "resseg_grupos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "resseg_grupos_total.jpeg"),
                            width=28,
                            height=5*numCol_grupos,
                            units="cm",
                            startRow=insert_line_grupos
                            )
                insertImage(workbook,
                            "resseg_ramos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "resseg_ramos_total.jpeg"),
                            width=28,
                            height=5*numCol_ramos,
                            units="cm",
                            startRow=insert_line_ramos
                            )


            }


    #Plotting selected companies
        else
            {
                insert_line_grupos <- 1
                insert_line_ramos <- 1
                coenti_2 <- as.numeric(strsplit(params_2[[4]][2],",")[[1]])

            #Getting companies names
                cias <- read.csv2(paste0("data",slash,"Ses_cias.csv"),
                                  encoding="latin1")
            #Removing old data
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^resseg_grupos_",
                                 full.names=TRUE))
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^resseg_ramos_",
                                 full.names=TRUE))

            #Making plots
                plots_grupos<- list()
                plots_ramos<- list()

                tabela_grupos <- 1
                tabela_ramos <- 1

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
                                  "resseg_grupos",
                                  cia,
                                  startRow=insert_line_grupos)

                        insert_line_grupos <- insert_line_grupos+1

                    #Transpose Data Frame
                        mov_grupos_sub_t <- mov_grupos_sub[,-1]
                        rownames(mov_grupos_sub_t) <- mov_grupos_sub[,1]
                        mov_grupos_sub_t <- as.data.frame(t(mov_grupos_sub_t))

                        writeDataTable(workbook,
                                       "resseg_grupos",
                                       format(
                                        mov_grupos_sub_t,
                                        decimal.mark=","),
                                       rowNames=TRUE,
                                       startRow=insert_line_grupos
                                       )

                        insert_line_grupos <- insert_line_grupos +
                                                nrow(mov_grupos_sub_t)+1

                        writeData(workbook,
                                  "resseg_grupos",
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

                        ggsave(paste0("resseg_grupos_",tabela_grupos,".jpeg"),
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
                ggsave("resseg_grupos_total.jpeg",
                       plot=plot_total_grupos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_grupos*length(plots_grupos),
                       limitsize=FALSE,
                       units="cm"
                       )

                insertImage(workbook,
                            "resseg_grupos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "resseg_grupos_total.jpeg"
                                   ),
                            width=28,
                            height=5*numCol_grupos*length(plots_grupos),
                            units="cm",
                            startRow=insert_line_grupos
                            )

                for(j in 1:length(coenti_2))
                    {
                        mov_ramos_sub <- mov_ramos[
                                        mov_ramos$coenti==coenti_2[j],-2]

                        if(nrow(mov_ramos_sub)==0)
                            {
                                next
                            }

                        numCol_ramos <- ceiling((ncol(mov_ramos_sub)-1)/3)

                        if(!params_2[[7]][1])
                            {
                                mov_ramos_sub <- mov_ramos_sub[1:
                                            (nrow(mov_ramos_sub)-1),]
                            }

                        cia <- as.character(cias[cias$Coenti==coenti_2[j],2])
                        cia <- stri_trans_general(cia,"Latin-ASCII")

                        writeData(workbook,
                                  "resseg_ramos",
                                  cia,
                                  startRow=insert_line_ramos)

                        insert_line_ramos <- insert_line_ramos+1

                    #Transpose Data Frame
                        mov_ramos_sub_t <- mov_ramos_sub[,-1]
                        rownames(mov_ramos_sub_t) <- mov_ramos_sub[,1]
                        mov_ramos_sub_t <- as.data.frame(t(mov_ramos_sub_t))

                        writeDataTable(workbook,
                                       "resseg_ramos",
                                       format(
                                        mov_ramos_sub_t,
                                        decimal.mark=","),
                                       rowNames=TRUE,
                                       startRow=insert_line_ramos
                                       )

                        insert_line_ramos <- insert_line_ramos + 
                                                nrow(mov_ramos_sub_t)+1

                        writeData(workbook,
                                  "resseg_ramos",
                                  paste0("Ver tabela ",tabela_ramos),
                                  startRow=insert_line_ramos)

                        insert_line_ramos <- insert_line_ramos+2

                    #Plotting data
                        period <- c("Mensal",
                                           "Trimestral",
                                           "Semestral",
                                           "Anual")[params_2[[7]]]

                        plot_data_ramos <- melt(mov_ramos_sub,id="yearsec")

                        #Plot Indexes
                        x_index_ramos <- rep(F,nrow(mov_ramos_sub))
                        x_length_ramos <- length(x_index_ramos)

                        if(ceiling(x_length_ramos/3)!=x_length_ramos/3)
                            {
                                first_third_ramos <- ceiling(x_length_ramos/3)
                                second_third_ramos <- floor(x_length_ramos*2/3)+1
                            }
                        else
                            {
                                first_third_ramos <- (x_length_ramos/3)+1
                                second_third_ramos <- x_length_ramos*2/3
                            }
        
                        x_index_ramos[c(1,
                                first_third_ramos,
                                second_third_ramos,
                                length(x_index_ramos))
                                ]<-T

                        p1_ramos <- ggplot(data=plot_data_ramos,
                                     aes(x=yearsec,
                                         y=value,
                                         colour=variable,
                                         group=variable))+
                                        geom_line()+
                                        scale_x_discrete(
                                            breaks=plot_data_ramos$yearsec[
                                                                x_index_ramos])+
                                        facet_wrap(~variable, 
                                                   scales="free_y",
                                                   ncol=3)+
                                        ggtitle(paste0(
                                            "Tabela ",
                                            tabela_ramos,
                                            ": ",
                                            cia,
                                            " \nPeriodicidade ",
                                            period,
                                            " \nDe ",
                                            mov_ramos_sub[1,1],
                                            " a ",
                                            mov_ramos_sub[nrow(mov_ramos_sub),1])
                                        )+
                                        ylab("Valor")+
                                        xlab("Periodo")
                        if(nrow(mov_ramos_sub)<15)
                            {
                                p1_ramos <- p1_ramos + geom_point()
                            }

                        ggsave(paste0("resseg_ramos_",tabela_ramos,".jpeg"),
                               plot=p1_ramos,
                               path=paste0("proc_data",slash,"plots"),
                               width=28,
                               height=5*numCol_ramos,
                               units="cm"
                               )

                        plots_ramos[[tabela_ramos]]<- p1_ramos

                        tabela_ramos <- tabela_ramos + 1

                    }

            #Merging plots and printing
                plot_total_ramos <- arrangeGrob(grobs=plots_ramos,ncol=1)
                ggsave("resseg_ramos_total.jpeg",
                       plot=plot_total_ramos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_ramos*length(plots_ramos),
                       limitsize=FALSE,
                       units="cm"
                       )

                insertImage(workbook,
                            "resseg_ramos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "resseg_ramos_total.jpeg"
                                   ),
                            width=28,
                            height=5*numCol_ramos*length(plots_ramos),
                            units="cm",
                            startRow=insert_line_ramos
                            )
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
        campos$noitem <- sapply(campos$noitem, function(x) stri_trans_tolower(x))
        campos$noitem <- sapply(campos$noitem, function(x) gsub("^\\(.*?\\) ","",x))

    #Getting filtered cmpids
        params_campos <- params_3[[7]]
        params_campos <- stri_trans_general(params_campos,"Latin-ASCII")
        params_campos <- stri_trim(params_campos)
        params_campos <- stri_trans_tolower(params_campos)
        params_campos <- gsub("^\\(.*?\\) ","",params_campos)
        params_campos <- unique(params_campos)
        campos_filter <- campos[campos$noitem %in% params_campos, 1]

    #Filtering data

        mov_grupos <- read.csv2(paste0("data",slash,"ses_valoresresmovgrupos.csv"))
        mov_grupos <- mov_grupos[,-6]
        mov_ramos <- read.csv2(paste0("data",slash,"SES_ValoresMovRamos.csv"))

        mov_grupos <- mov_grupos[mov_grupos$CMPID %in% campos_filter, ]
        mov_ramos <- mov_ramos[mov_ramos$cmpid %in% campos_filter, ]

    #Subsetting by Initial Date

        if(nchar(params_3[[2]][1])==10)
            {
                init_date_3 <- as.Date(params_3[[2]][1], format="%d/ %m/ %Y")
                if(!is.na(init_date_3))
                    {
                        init_date_3 <- format(init_date_3,"%Y%m")
                        if(nchar(init_date_3)==6)
                            {
                                mov_grupos <- mov_grupos[mov_grupos$DAMESANO >= as.numeric(init_date_3),]
                                mov_ramos <- mov_ramos[mov_ramos$damesano >= as.numeric(init_date_3),]
                            }
                    }
            }

    #Subsetting by Final Date
        if(nchar(params_3[[2]][2])==10)
            {
                final_date_3 <- as.Date(params_3[[2]][2], format="%d/ %m/ %Y")
                if(!is.na(final_date_3))
                    {
                        final_date_3 <- format(final_date_3,"%Y%m")
                        if(nchar(final_date_3)==6)
                            {
                                mov_grupos <- mov_grupos[mov_grupos$DAMESANO <= as.numeric(final_date_3),]
                                mov_ramos <- mov_ramos[mov_ramos$damesano <= as.numeric(final_date_3),]
                            }
                    }
            }

    #Subsetting by Company
        if(params_3[[3]][1]=="FALSE")
            {
                coenti_3 <- as.numeric(strsplit(params_3[[3]][2],",")[[1]])
                mov_grupos <- mov_grupos[mov_grupos$COENTI %in% coenti_3,]
                mov_ramos <- mov_ramos[mov_ramos$coenti %in% coenti_3,]
            }

    #Subsetting by Ramos
        if(params_3[[4]][1]=="FALSE")
            {
                ramos_3 <- as.numeric(strsplit(params_3[[4]][2],",")[[1]])
                mov_ramos <- mov_ramos[mov_ramos$ramcodigo %in% ramos_3,]
            }

    #Subsetting by Grupos
        if(params_3[[5]][1]=="FALSE")
            {
                grupos_3 <- as.numeric(strsplit(params_3[[5]][2],",")[[1]])
                mov_grupos <- mov_grupos[mov_grupos$GRACODIGO %in% grupos_3,]
                mov_ramos <- mov_ramos[mov_ramos$gracodigo %in% grupos_3,]
            }

    #Aggreggating by time interval
        
        #Aggregate by Year/Month
        if(params_3[[6]][1] == TRUE)
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
        else if(params_3[[6]][2] == TRUE)
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
        else if(params_3[[6]][3] == TRUE)
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
        else if(params_3[[6]][4] == TRUE)
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

        mov_grupos_campos <- stri_trans_general(
                                stri_trim(
                                merge(
                                    campos,
                                    mov_grupos,
                                    by.y="CMPID",
                                    by.x="nuitem")[,2]
                                ),"Latin-ASCII")

        mov_grupos[,3] <- mov_grupos_campos
        mov_grupos <- aggregate(mov_grupos[,5],
                                by=list(
                                     yearsec=mov_grupos$YEARSEC,
                                     coenti=mov_grupos$COENTI,
                                     #gracodigo=mov_grupos$GRACODIGO
                                     cmpid=mov_grupos$CMPID
                                     )
                             ,FUN=sum,na.rm=TRUE)
        mov_grupos_melt <- mov_grupos
        mov_grupos <- dcast(data=mov_grupos,
                              formula= yearsec + coenti ~ cmpid,
                              fun.aggregate=sum,
                              value.var="x")
        #mov_grupos_names <- colnames(mov_grupos)[3:ncol(mov_grupos)]
        #colnames(mov_grupos) <- c(colnames(mov_grupos)[1:2],mov_grupos_names)
        write.csv2(mov_grupos,paste0("proc_data",slash,"dem_cont_grupos.csv"))

        mov_ramos_campos <- stri_trans_general(
                                stri_trim(
                                merge(
                                    campos,
                                    mov_ramos,
                                    by.y="cmpid",
                                    by.x="nuitem")[,2]
                                ),"Latin-ASCII")

        mov_ramos[,3] <- mov_ramos_campos

        mov_ramos <- aggregate(mov_ramos[,6],
                                by=list(
                                     yearsec=mov_ramos$yearsec,
                                     coenti=mov_ramos$coenti,
                                     #ramcodigo=mov_ramos$ramcodigo,
                                     #gracodigo=mov_ramos$gracodigo
                                     cmpid=mov_ramos$cmpid
                                     )
                             ,FUN=sum,na.rm=TRUE)
        mov_ramos_melt <- mov_ramos
        mov_ramos <- dcast(data=mov_ramos,
                              formula= yearsec + coenti ~ cmpid,
                              fun.aggregate=sum,
                              value.var="x")
        write.csv2(mov_ramos,paste0("proc_data",slash,"dem_cont_ramos.csv"))

        addWorksheet(workbook,"dem_cont_grupos")
        addWorksheet(workbook,"dem_cont_ramos")

        insert_line <- 1

    #Merging All Companies
        if(params_3[[3]][1]=="TRUE")
            {

                mov_grupos <- aggregate(mov_grupos[,3:ncol(mov_grupos)],
                                by=list(yearsec=mov_grupos$yearsec),
                                FUN=sum,na.rm=TRUE)

                mov_ramos <- aggregate(mov_ramos[,3:ncol(mov_ramos)],
                                by=list(yearsec=mov_ramos$yearsec),
                                FUN=sum,na.rm=TRUE)

                writeData(workbook,
                          "dem_cont_grupos",
                          "Todas as Empresas",
                          startRow=insert_line)

                writeData(workbook,
                          "dem_cont_ramos",
                          "Todas as Empresas",
                          startRow=insert_line)

                insert_line <- insert_line + 1
                numCol_grupos <- ceiling((ncol(mov_grupos)-1)/3)
                numCol_ramos <- ceiling((ncol(mov_ramos)-1)/3)

            #Transpose Data Frame
                mov_grupos_t <- mov_grupos[,-1]
                mov_ramos_t <- mov_ramos[,-1]
                rownames(mov_grupos_t) <- mov_grupos[,1]
                rownames(mov_ramos_t) <- mov_ramos[,1]
                mov_grupos_t <- as.data.frame(t(mov_grupos_t))
                mov_ramos_t <- as.data.frame(t(mov_ramos_t))

                writeDataTable(workbook,
                               "dem_cont_grupos",
                               format(
                                mov_grupos_t,
                                decimal.mark=","),
                               rowNames=TRUE,
                               startRow=insert_line
                               )

                writeDataTable(workbook,
                               "dem_cont_ramos",
                               format(
                                mov_ramos_t,
                                decimal.mark=","),
                               rowNames=TRUE,
                               startRow=insert_line
                               )

                insert_line_grupos <- insert_line+nrow(mov_grupos_t)+1
                insert_line_ramos <- insert_line+nrow(mov_ramos_t)+1

                writeData(workbook,
                          "dem_cont_grupos",
                          "Ver Tabela 1",
                          startRow=insert_line_grupos)
                writeData(workbook,
                          "dem_cont_ramos",
                          "Ver Tabela 1",
                          startRow=insert_line_ramos)

                insert_line_grupos <- insert_line_grupos + 2
                insert_line_ramos <- insert_line_ramos + 2

            #Removing old files
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^dem_cont_grupos_",
                                 full.names=TRUE))
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^dem_cont_ramos_",
                                 full.names=TRUE))
            #Plotting data
                period <- c("Mensal",
                                   "Trimestral",
                                   "Semestral",
                                   "Anual")[params_3[[6]]]

                plot_data_grupos <- melt(mov_grupos,id="yearsec")
                plot_data_ramos <- melt(mov_ramos,id="yearsec")

                #Plot Indexes
                x_index_grupos <- rep(F,nrow(mov_grupos))
                x_length_grupos <- length(x_index_grupos)

                x_index_ramos <- rep(F,nrow(mov_ramos))
                x_length_ramos <- length(x_index_ramos)

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

                if(ceiling(x_length_ramos/3)!=x_length_ramos/3)
                    {
                        first_third_ramos <- ceiling(x_length_ramos/3)
                        second_third_ramos <- floor(x_length_ramos*2/3)+1
                    }
                else
                    {
                        first_third_ramos <- (x_length_ramos/3)+1
                        second_third_ramos <- x_length_ramos*2/3
                    }

                x_index_grupos[c(1,
                        first_third_grupos,
                        second_third_grupos,
                        length(x_index_grupos))
                        ]<-T

                x_index_ramos[c(1,
                        first_third_ramos,
                        second_third_ramos,
                        length(x_index_ramos))
                        ]<-T

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

                p1_ramos <- ggplot(data=plot_data_ramos,
                             aes(x=yearsec,y=value,colour=variable,group=variable))+
                                geom_line()+
                                scale_x_discrete(
                                    breaks=plot_data_ramos$yearsec[x_index_ramos]
                                    )+
                                facet_wrap(~variable, scales="free_y",ncol=3)+
                                ggtitle(paste0(
                                    "Tabela 1: Todas as Empresas\nPeriodicidade ",
                                    period,
                                    " \nDe ",
                                    mov_ramos[1,1],
                                    " a ",
                                    mov_ramos[nrow(mov_ramos),1])
                                )+
                                ylab("Valor")+
                                xlab("Periodo")
                if(nrow(mov_ramos)<15)
                    {
                        p1_ramos <- p1_ramos + geom_point()
                    }

                ggsave("dem_cont_grupos_total.jpeg",
                       plot=p1_grupos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_grupos,
                       units="cm"
                       )
                ggsave("dem_cont_ramos_total.jpeg",
                       plot=p1_ramos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_ramos,
                       units="cm"
                       )

                insertImage(workbook,
                            "dem_cont_grupos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "dem_cont_grupos_total.jpeg"),
                            width=28,
                            height=5*numCol_grupos,
                            units="cm",
                            startRow=insert_line_grupos
                            )
                insertImage(workbook,
                            "dem_cont_ramos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "dem_cont_ramos_total.jpeg"),
                            width=28,
                            height=5*numCol_ramos,
                            units="cm",
                            startRow=insert_line_ramos
                            )


            }

    #Plotting selected companies
        else
            {
                insert_line_grupos <- 1
                insert_line_ramos <- 1
                coenti_3 <- as.numeric(strsplit(params_3[[3]][2],",")[[1]])

            #Getting companies names
                cias <- read.csv2(paste0("data",slash,"Ses_cias.csv"),
                                  encoding="latin1")
            #Removing old data
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^dem_cont_grupos_",
                                 full.names=TRUE))
                file.remove(list.files(path=paste0("proc_data",slash,"plots"),
                                 pattern="^dem_cont_ramos_",
                                 full.names=TRUE))

            #Making plots
                plots_grupos<- list()
                plots_ramos<- list()

                tabela_grupos <- 1
                tabela_ramos <- 1

                for(j in 1:length(coenti_3))
                    {
                        mov_grupos_sub <- mov_grupos[
                                        mov_grupos$coenti==coenti_3[j],-2]

                        if(nrow(mov_grupos_sub)==0)
                            {
                                next
                            }

                        numCol_grupos <- ceiling((ncol(mov_grupos_sub)-1)/3)

                        if(!params_3[[6]][1])
                            {
                                mov_grupos_sub <- mov_grupos_sub[1:
                                            (nrow(mov_grupos_sub)-1),]
                            }

                        cia <- as.character(cias[cias$Coenti==coenti_3[j],2])
                        cia <- stri_trans_general(cia,"Latin-ASCII")

                        writeData(workbook,
                                  "dem_cont_grupos",
                                  cia,
                                  startRow=insert_line_grupos)

                        insert_line_grupos <- insert_line_grupos+1

                    #Transpose Data Frame
                        mov_grupos_sub_t <- mov_grupos_sub[,-1]
                        rownames(mov_grupos_sub_t) <- mov_grupos_sub[,1]
                        mov_grupos_sub_t <- as.data.frame(t(mov_grupos_sub_t))

                        writeDataTable(workbook,
                                       "dem_cont_grupos",
                                       format(
                                        mov_grupos_sub_t,
                                        decimal.mark=","),
                                       rowNames=TRUE,
                                       startRow=insert_line_grupos
                                       )

                        insert_line_grupos <- insert_line_grupos +
                                                nrow(mov_grupos_sub_t)+1

                        writeData(workbook,
                                  "dem_cont_grupos",
                                  paste0("Ver tabela ",tabela_grupos),
                                  startRow=insert_line_grupos)

                        insert_line_grupos <- insert_line_grupos+2

                    #Plotting data
                        period <- c("Mensal",
                                           "Trimestral",
                                           "Semestral",
                                           "Anual")[params_3[[6]]]

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

                        ggsave(paste0("dem_cont_grupos_",tabela_grupos,".jpeg"),
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
                ggsave("dem_cont_grupos_total.jpeg",
                       plot=plot_total_grupos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_grupos*length(plots_grupos),
                       limitsize=FALSE,
                       units="cm"
                       )

                insertImage(workbook,
                            "dem_cont_grupos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "dem_cont_grupos_total.jpeg"
                                   ),
                            width=28,
                            height=5*numCol_grupos*length(plots_grupos),
                            units="cm",
                            startRow=insert_line_grupos
                            )

                for(j in 1:length(coenti_3))
                    {
                        mov_ramos_sub <- mov_ramos[
                                        mov_ramos$coenti==coenti_3[j],-2]

                        if(nrow(mov_ramos_sub)==0)
                            {
                                next
                            }

                        numCol_ramos <- ceiling((ncol(mov_ramos_sub)-1)/3)

                        if(!params_3[[6]][1])
                            {
                                mov_ramos_sub <- mov_ramos_sub[1:
                                            (nrow(mov_ramos_sub)-1),]
                            }

                        cia <- as.character(cias[cias$Coenti==coenti_3[j],2])
                        cia <- stri_trans_general(cia,"Latin-ASCII")

                        writeData(workbook,
                                  "dem_cont_ramos",
                                  cia,
                                  startRow=insert_line_ramos)

                        insert_line_ramos <- insert_line_ramos+1

                    #Transpose Data Frame
                        mov_ramos_sub_t <- mov_ramos_sub[,-1]
                        rownames(mov_ramos_sub_t) <- mov_ramos_sub[,1]
                        mov_ramos_sub_t <- as.data.frame(t(mov_ramos_sub_t))

                        writeDataTable(workbook,
                                       "dem_cont_ramos",
                                       format(
                                        mov_ramos_sub_t,
                                        decimal.mark=","),
                                       rowNames=TRUE,
                                       startRow=insert_line_ramos
                                       )

                        insert_line_ramos <- insert_line_ramos + 
                                                nrow(mov_ramos_sub_t)+1

                        writeData(workbook,
                                  "dem_cont_ramos",
                                  paste0("Ver tabela ",tabela_ramos),
                                  startRow=insert_line_ramos)

                        insert_line_ramos <- insert_line_ramos+2

                    #Plotting data
                        period <- c("Mensal",
                                           "Trimestral",
                                           "Semestral",
                                           "Anual")[params_3[[6]]]

                        plot_data_ramos <- melt(mov_ramos_sub,id="yearsec")

                        #Plot Indexes
                        x_index_ramos <- rep(F,nrow(mov_ramos_sub))
                        x_length_ramos <- length(x_index_ramos)

                        if(ceiling(x_length_ramos/3)!=x_length_ramos/3)
                            {
                                first_third_ramos <- ceiling(x_length_ramos/3)
                                second_third_ramos <- floor(x_length_ramos*2/3)+1
                            }
                        else
                            {
                                first_third_ramos <- (x_length_ramos/3)+1
                                second_third_ramos <- x_length_ramos*2/3
                            }
        
                        x_index_ramos[c(1,
                                first_third_ramos,
                                second_third_ramos,
                                length(x_index_ramos))
                                ]<-T

                        p1_ramos <- ggplot(data=plot_data_ramos,
                                     aes(x=yearsec,
                                         y=value,
                                         colour=variable,
                                         group=variable))+
                                        geom_line()+
                                        scale_x_discrete(
                                            breaks=plot_data_ramos$yearsec[
                                                                x_index_ramos])+
                                        facet_wrap(~variable, 
                                                   scales="free_y",
                                                   ncol=3)+
                                        ggtitle(paste0(
                                            "Tabela ",
                                            tabela_ramos,
                                            ": ",
                                            cia,
                                            " \nPeriodicidade ",
                                            period,
                                            " \nDe ",
                                            mov_ramos_sub[1,1],
                                            " a ",
                                            mov_ramos_sub[nrow(mov_ramos_sub),1])
                                        )+
                                        ylab("Valor")+
                                        xlab("Periodo")
                        if(nrow(mov_ramos_sub)<15)
                            {
                                p1_ramos <- p1_ramos + geom_point()
                            }

                        ggsave(paste0("dem_cont_ramos_",tabela_ramos,".jpeg"),
                               plot=p1_ramos,
                               path=paste0("proc_data",slash,"plots"),
                               width=28,
                               height=5*numCol_ramos,
                               units="cm"
                               )

                        plots_ramos[[tabela_ramos]]<- p1_ramos

                        tabela_ramos <- tabela_ramos + 1

                    }

            #Merging plots and printing
                plot_total_ramos <- arrangeGrob(grobs=plots_ramos,ncol=1)
                ggsave("dem_cont_ramos_total.jpeg",
                       plot=plot_total_ramos,
                       path=paste0("proc_data",slash,"plots"),
                       width=28,
                       height=5*numCol_ramos*length(plots_ramos),
                       limitsize=FALSE,
                       units="cm"
                       )

                insertImage(workbook,
                            "dem_cont_ramos",
                            paste0("proc_data",
                                   slash,
                                   "plots",
                                   slash,
                                   "dem_cont_ramos_total.jpeg"
                                   ),
                            width=28,
                            height=5*numCol_ramos*length(plots_ramos),
                            units="cm",
                            startRow=insert_line_ramos
                            )
            }


    }

saveWorkbook(workbook,paste0("SES-Susep-",Sys.Date(),".xlsx"),overwrite=TRUE)
