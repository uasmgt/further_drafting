
# Пакеты ----
library(dplyr)
library(openxlsx)
library(lubridate)

# Путь до рабочей папки ----
setwd("~/aiso/2019")

# Функция для обработки выгрузок АИСО ----
# Объединение выгрузок в одну таблицу
CreateDataset <- function(x)
{
  files <- list.files(path = x, pattern = "Реестр") # создаёт список файлов
  files.list <- lapply(files, read.xlsx) # создаёт лист необработанных таблиц
  dataset <- lapply(files.list, CreateTable) # обрабатывает таблицы в листе
  y <- do.call(rbind.data.frame, dataset) # объединяет таблицы
  return(y) # возвращает таблицу
}

# Вспомогательная подготовка таблиц к объединению
CreateTable <- function(x)
{
  names(x) <- as.character(unlist(x[1, ])) # забирает заголовки столбцов
  x <- x[-c(1), ] # удаляет строки без данных
}

# Центральный АО ----
setwd("cao/")
cao <- CreateDataset("./")
setwd("..")

# Новомосковский АО ----
setwd("nao/")
nao <- CreateDataset("./")
setwd("..")

# Северный АО ----
setwd("sao/")
sao <- CreateDataset("./")
setwd("..")

# Северо-восточный АО ----
setwd("svao/")
svao <- CreateDataset("./")
setwd("..") 

# Северо-западный АО ----
setwd("szao/")
szao <- CreateDataset("./")
setwd("..") 

# Троицкий АО ----
setwd("tao/")
tao <- CreateDataset("./")
setwd("..") 

# Восточный АО ----
setwd("vao/")
vao <- CreateDataset("./")
setwd("..") 

# Южный АО ----
setwd("yao/")
yao <- CreateDataset("./")
setwd("..") 

# Юго-восточный АО ----
setwd("yvao/")
yvao <- CreateDataset("./")
setwd("..") 

# Юго-восточный АО ----
setwd("yzao/")
yzao <- CreateDataset("./")
setwd("..") 

# Западный АО ----
setwd("zao/")
zao <- CreateDataset("./")
setwd("..") 

# Зеленоградский АО ----
setwd("zelao/")
zelao <- CreateDataset("./")
setwd("..") 

# Создать вектор с названиями переменных ----
districts <- grep("ao", ls(), value = TRUE)

# Присовение названий столбцов ----
for(x in districts) data.table::setnames(get(x),  
                                         c("app_no", "app_no_portal", "voucher_no",
                                           "date_time", "status", "denial_reason",
                                           "exec_auth", "office", "list", 
                                           "payment", "purpose", "vacation_spot",
                                           "vacation_address", "period", "theme",
                                           "category", "surname_camper", "name_camper",
                                           "patronym_camper", "gender_camper", "birthdate_camper",
                                           "age_camper", "birthplace_camper", "SNILS_camper", 
                                           "ID_camper", "IDser_camper", "IDno_camper",
                                           "IDissuedate_camper", "IDissueplace_camper", "disorder_cat",
                                           "disorder_sub", "benefit", "reg_address", 
                                           "ticket_to_rejection", "ticket_from_rejection", 
                                           "surname_applicant", "name_applicant", 
                                           "patronym_applicant", "ID_applicant",
                                           "IDser_applicant", "IDno_applicant", "phone",
                                           "email", "name_mother", "birthdate_mother",
                                           "name_father", "birthdate_father"))
rm(x)

SubSetKids <- function(x){
  x$birthdate_camper <- dmy(x$birthdate_camper)
  x$date_time <- unlist(strsplit(x$date_time, split = " "))[1]
  x$date_time <- dmy(x$date_time)
  x$age <- year(as.Date(x$date_time)) - year(as.Date(x$birthdate_camper))
  x <- x %>% filter(age >= 6 & age < 18)
  return(x)
}

districts.list <- do.call("list", mget(grep("ao", ls(), value = TRUE)))

districts.list <- lapply(districts.list, SubSetKids)

list2env(districts.list, .GlobalEnv)

TblFun <- function(x){
  tbl <- table(x)
  res <- cbind(tbl)
  colnames(res) <- c("Count")
  res
}


data.frame(lapply(districts.list, nrow))
