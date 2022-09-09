# a1<-c(1,2,4,1,2)
# a2<-c(2,2,2,1,1)
# a3<-c(1,2,1,3,1)
# a<-data.frame(a1,a2,a3)
# 
# b1<-c(2,1,3,1,2)
# b2<-c(2,1,2,1,1)
# b3<-c(1,1,1,3,1)
# b<-data.frame(b1,b2,b3)
# rownames(a)<-rownames(b)<-c("nacional","Ent1","Ent2","Ent3","Ent4")
# 
# a<-cbind(rownames(a),a);rownames(a)<-NULL;
# b<-cbind(rownames(b),b);rownames(b)<-NULL
# 
# colnames(b)<-colnames(a)<-c("Ent","NEM_1","NEM_2","NEM_3")

#--- formas de comparar --------------------------------------------------------

#identical(a,b)
carga_excel_a <- function(ruta1, ruta2, num_hoja, busca){
  h1 <- openxlsx::readWorkbook(ruta1,num_hoja, 
                               skipEmptyRows = F, skipEmptyCols = F)
  h2 <- openxlsx::readWorkbook(ruta2,num_hoja,
                               skipEmptyRows = F, skipEmptyCols = F)
  
  
  inicia <- which(str_detect(h1[,1],busca)%in%T)
  finaliza <- which(str_detect(h1[,1], "Fuente:")%in%T)
  h_1 <- h1[inicia:finaliza, ]
  h_2 <- h2[inicia:finaliza, ]
  
  dfs <- list()
  dfs[[1]] <- h_1
  dfs[[2]] <- h_2
  return(dfs)
}

caracter_a <- function(df1, df2){
  inicia <-  which(df1[,1]%in%"Estados Unidos Mexicanos")
  
  for(i in inicia:dim(df1)[1]){
    for(j in 2:dim(df1)[2]){
        df1[i, j] <- as.character(df1[i, j]) 
        df2[i, j] <- as.character(df2[i, j]) 
        
    }
  }
  
  dfs <- list()
  dfs[[1]] <- df1
  dfs[[2]] <- df2
  
  return(dfs)
}

rec_dig_M <- function(x, digits, chars = TRUE) {
  if(grepl(x = x, pattern = "\\.")) {
    y=as.character(x)
    pos=grep(unlist(strsplit(x = y, split = "")), pattern = "\\.", value = FALSE)
    if(chars) {
      return(substr(x = x, start = 1, stop = pos + digits))
    }
    return(
      as.numeric(substr(x = x, start = 1, stop = pos + digits))
    )
  } else {
    return(
      #format(round(x, 2), nsmall = 2)
      x
    )
  }
}

compara_a <- function(df1, df2, nom_hoja){
#-------------------------------------------------------------------------------
library(tidyverse)
library(compare)
estilo <- openxlsx::createStyle(fontName = "Arial",
                                fontSize = 8,
                                fontColour = NULL,
                                numFmt = "TEXT",
                                border = NULL,
                                borderColour = getOption("openxlsx.borderColour", "black"),
                                borderStyle = getOption("openxlsx.borderStyle", "dotted"),
                                bgFill = NULL,
                                fgFill = "#B9C558",
                                halign = "right",
                                valign = "center",
                                textDecoration = "bold",
                                wrapText = FALSE,
                                textRotation = NULL,
                                indent = NULL,
                                locked = NULL,
                                hidden = NULL)
#-------------------------------------------------------------------------------
#Acomodo primero los dataframes para hacer la revisión
dfs <- caracter_a(df1, df2)

inicia <-  which(dfs[[1]][,1]%in%pa_buscar)

for(i in inicia:dim(dfs[[1]])[1]){
  for(j in 2:dim(dfs[[1]])[2]){
    df1[i, j] <- rec_dig_M(dfs[[1]][i, j],5) 
    df2[i, j] <- rec_dig_M(dfs[[2]][i, j],5) 
    
  }
}

#Voy a usa el anti_join para revisar tabulados.
renglon <- anti_join(df1,df2)
renglon <- as.vector(renglon[[1]])
mals_r <- which(df1[,1]%in%renglon)
#semi_join(a,b)


columnas <- as.vector(compare(df1, df2)[[2]])
mals_c <- which(columnas%in%"FALSE")

if(!length(mals_r)==0){
#wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, nom_hoja)
openxlsx::writeData(wb, nom_hoja, df1 , startCol = 1, colNames = T)


for(i in 1:length(mals_r)){
  for (j in 1:length(mals_c)){

    if(isTRUE(df1[mals_r[i], mals_c[j]] == df2[mals_r[i], mals_c[j]])){
       df1[mals_r[i], mals_c[j]] <- df1[mals_r[i], mals_c[j]]
    }else{
    openxlsx::addStyle(wb, nom_hoja, style = estilo, rows = mals_r[i] + 1, cols = mals_c[j]  )
  }
 }
 }
}else{}

}
#nom_arch <- "prueba_comparacion"
#dir <- "D:/Varios/"
#arch<-paste0(dir,nom_arch,".xlsx")
#openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)
#-------------------------------------------------------------------------------
# w1 <- openxlsx::loadWorkbook("D:/ENCIG/2021/Tabs/tabs_2021/Secc3/III_experiencia_pagos_tramites_solicitudes_encig2021_TC_01_est.xlsx")
# w2 <- openxlsx::loadWorkbook("D:/ENCIG/2021/Tabs/tabs_ENCIG_2021_1_3_6_8_006/S_3/III_experiencia_pagos_tramites_solicitudes_encig2021_est.xlsx")
# num_hoja <- 2
# 
# ruta1 <- w1; ruta2 <- w2


# 
# #-----------------------------------------------------------------------------------
# dfs<-carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = 2, busca = "Estados Unidos Mexicanos")
# 
#compara_a(df1 = dfs[[1]], df2 = dfs[[2]], dir = "D:/Varios/" ,nom_arch = "prueba_comparacion" , nom_hoja = "alex")
w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccII/enve2022_est_II_caracterizacion_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_est_II_caracterizacion_delitoM.xlsx"
nom_arch <- "comparacion_secc2_ENVE_2022_est"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

#-----
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = 2 , busca = pa_buscar)

# dfs <- caracter_a(dfs[[1]], dfs[[2]])
# 
# inicia <-  which(dfs[[1]][,1]%in%pa_buscar)
# 
# for(i in inicia:dim(dfs[[1]])[1]){
#   for(j in 2:dim(dfs[[1]])[2]){
#     dfs[[1]][i, j] <- rec_dig_M(dfs[[1]][i, j],5) 
#     dfs[[2]][i, j] <- rec_dig_M(dfs[[2]][i, j],5) 
#     
#   }
# }
#compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0("Tab2.2"))

#-----
for(i in 1:29)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab2.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccIII/enve2022_est_III_denuncia_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_est_III_denuncia_delitoMProp.xlsx"
nom_arch <- "comparacion_secc3_ENVE_2022_est"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab3.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccV/enve2022_est_V_desempeno_institucional.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_est_V_desempeno_MProp.xlsx"
nom_arch <- "comparacion_secc5_ENVE_2022_est"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab5.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)


#===============================================================================
w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVI/enve2022_est_VI_perdidas_consecuencia_inseguridad.xlsx"
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_est_VI_perdidas_M.xlsx"
nom_arch <- "comparacion_secc6_ENVE_2022_est"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab6.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)



#===============================================================================
w2 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVII/enve2022_est_VII_afectacion_comercio_informal.xlsx"
w1 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_est_VII_afectacion_M.xlsx"
nom_arch <- "comparacion_secc7_ENVE_2022_est"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab7.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

###############################################################################

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccII/enve2022_cv_II_caracterizacion_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_cv_II_caracterizacion_delitoM.xlsx"
nom_arch <- "comparacion_secc2_ENVE_2022_cv"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:29)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab2.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccIII/enve2022_cv_III_denuncia_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_cv_III_denuncia_delitoMProp.xlsx"
nom_arch <- "comparacion_secc3_ENVE_2022_cv"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab3.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccV/enve2022_cv_V_desempeno_institucional.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_cv_V_desempeno_MProp.xlsx"
nom_arch <- "comparacion_secc5_ENVE_2022_cv"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab5.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)


#===============================================================================
w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVI/enve2022_cv_VI_perdidas_consecuencia_inseguridad.xlsx"
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_cv_VI_perdidas_M.xlsx"
nom_arch <- "comparacion_secc6_ENVE_2022_cv"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab6.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)



#===============================================================================
w2 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVII/enve2022_cv_VII_afectacion_comercio_informal.xlsx"
w1 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_cv_VII_afectacion_M.xlsx"
nom_arch <- "comparacion_secc7_ENVE_2022_cv"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab7.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)


################################################################################


w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccII/enve2022_int_II_caracterizacion_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_int_II_caracterizacion_delitoM.xlsx"
nom_arch <- "comparacion_secc2_ENVE_2022_int"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:29)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab2.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccIII/enve2022_int_III_denuncia_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_int_III_denuncia_delitoMProp.xlsx"
nom_arch <- "comparacion_secc3_ENVE_2022_int"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab3.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccV/enve2022_int_V_desempeno_institucional.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_int_V_desempeno_MProp.xlsx"
nom_arch <- "comparacion_secc5_ENVE_2022_int"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab5.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)


#===============================================================================
w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVI/enve2022_int_VI_perdidas_consecuencia_inseguridad.xlsx"
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_int_VI_perdidas_M.xlsx"
nom_arch <- "comparacion_secc6_ENVE_2022_int"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab6.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)



#===============================================================================
w2 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVII/enve2022_int_VII_afectacion_comercio_informal.xlsx"
w1 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_int_VII_afectacion_M.xlsx"
nom_arch <- "comparacion_secc7_ENVE_2022_int"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab7.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)


################################################################################


w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccII/enve2022_se_II_caracterizacion_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_err_II_caracterizacion_delitoM.xlsx"
nom_arch <- "comparacion_secc2_ENVE_2022_err"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:29)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab2.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccIII/enve2022_se_III_denuncia_delito.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_err_III_denuncia_delitoMProp.xlsx"
nom_arch <- "comparacion_secc3_ENVE_2022_err"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab3.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)

#===============================================================================

w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccV/enve2022_se_V_desempeno_institucional.xlsx" #alex
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_err_V_desempeno_MProp.xlsx"
nom_arch <- "comparacion_secc5_ENVE_2022_err"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:14)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab5.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)


#===============================================================================
w1 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVI/enve2022_se_VI_perdidas_consecuencia_inseguridad.xlsx"
w2 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_err_VI_perdidas_M.xlsx"
nom_arch <- "comparacion_secc6_ENVE_2022_err"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab6.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)



#===============================================================================
w2 <- "D:/ENVE/ENVE 2022/Tabulados/Tabs_2022/SeccVII/enve2022_se_VII_afectacion_comercio_informal.xlsx"
w1 <- "D:/ENVE/ENVE 2022/comparacion/manu/enve2022_err_VII_afectacion_M.xlsx"
nom_arch <- "comparacion_secc7_ENVE_2022_err"
wb <- openxlsx::createWorkbook()
dir <- "D:/ENVE/ENVE 2022/comparacion/alex_compara/"
arch<-paste0(dir,nom_arch,".xlsx")
pa_buscar <- "Estados Unidos Mexicanos"

for(i in 1:7)eval(parse(text = paste0("
dfs <- carga_excel_a(ruta1 = w1, ruta2 = w2, num_hoja = i+1 , busca = pa_buscar)
compara_a(df1 = dfs[[1]], df2 = dfs[[2]], nom_hoja = paste0(\"Tab7.\", i))
print(i)
")))

openxlsx::saveWorkbook(wb, file = arch, overwrite = TRUE)


################################################################################
























