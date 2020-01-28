# Назначение: подсчитывает количество фактически отдохнувших
# детей-инвалидов по административным округам Москвы.

# Пакеты ----
library(dplyr)
library(openxlsx)

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

# Функция для создания подмассива из детей-инвалидов ----
SubsetDisabled <- function(x){
  disabled <- x %>% 
    filter(benefit == 
             "Дети-инвалиды, дети с ограниченными возможностями здоровья" | 
             benefit == "Дети-инвалиды")
  disabled <- disabled %>% filter(as.numeric(age_camper) >= 7 &
                                    as.numeric(age_camper) <= 17)
  id <- paste(disabled$IDser_camper,
              disabled$IDno_camper)
  disabled <- cbind(disabled, id)
  disabled$id <- as.character(disabled$id)
  return(disabled)
}

districts.list <- do.call("list", mget(grep("ao", ls(), value = TRUE)))

districts.list <- lapply(districts.list, SubsetDisabled)

list2env(districts.list, .GlobalEnv)

# Подсчёт количества инвалидов и его сохранение в переменную -----------
# Воспользоваться функцией "найти и заменить" (ctrl + f в RStudio) для
# подстановки нужного округа
CountDisabled <- function(x){
  x <- read.xlsx(x, sheet = 3)
  names(x) <- as.character(unlist(x[4, ]))     # название колонок
  x <- x[-c(1:4), ]                            # удаление пустых рядов
  x <- x[c(1, 2, 3, 5, 6, 9, 10, 23)]             # удаление ненужных колонок
  portal <- subset(x, grepl("[0-9]{4}\\-[0-9]{7}\\-[0-9]{6}\\-[0-9]{7}\\/[0-9]{2}", `Номер заявления`))
  portal <- portal %>% filter(`Цель обращения` != "Отдых для сирот (совместный отдых)")
  portal <- portal %>% filter(`Цель обращения` != "Молодёжный отдых для лиц из числа детей-сирот и детей, оставшихся без попечения родителей, 18-23 лет")
  # подсчёт количества
  portal.arrived <- nrow(portal %>% filter(`Заехал` == "Заехал"))
  portal.kids <- portal %>% filter(`Заехал` == "Заехал" & `Ребёнок / сопровождающий` == "Ребёнок")
  diasbled.kids <- nrow(portal.kids[portal.kids$`Номер документа` %in% szao$id, ])
# Возрат количества детей
  return(diasbled.kids)
}

setwd("~/aism/2019/arrivals_fam/")
file.names <- list.files(path = "./", recursive = TRUE, 
                         pattern = "*.xlsx") # чтение названий xlsx-файлов
szao <- lapply(file.names, CountDisabled)
szao <- sum(unlist(szao))

# Сохранение итоговой таблицы
districts
family.disabled <- rbind(cao, nao, sao, svao, szao, tao, vao, yao, yvao, 
                         yzao, zao, zelao)
write.csv2(family.disabled, "~/data/family_disabled.csv")
