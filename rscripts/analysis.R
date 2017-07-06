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

package_list <- c("openxlsx")
	for( i in length(package_list))
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

#Analysis 1

if(params_1[[1]]==TRUE)
    {
        seguros <- read.csv(paste0("data",slash,"Ses_seguros.csv"),sep=";")

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

    #Writing file

        write.csv2(seguros,paste0("proc_data",slash,"premios_sinistros_seg.csv"))
        write.xlsx(seguros,paste0("SES-Susep-",Sys.Date(),".xlsx"))
    }
head(seguros)
