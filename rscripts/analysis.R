#Retrieving passed arguments

myArgs <- commandArgs(trailingOnly = TRUE)
slash <- myArgs[1]
data_folder <- myArgs[2]
config_path <- myArgs[3]
log_file <- myArgs[4]

setwd(data_folder)

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

list(run_analysis_1,param_1_1,param_1_2,param_1_3,param_1_4,param_1_5)
