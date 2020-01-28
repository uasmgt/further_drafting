# Библиотеки -----------------------------------------------------------
library(openxlsx)

# Функции --------------------------------------------------------------
GetInfo <- function(x){
  camp <- x[1, 2]
  if (ncol(x) >= 30){
    term <- as.character(x[1, 26])
  } else if (ncol(x) >= 23){
    term <- as.character(x[1, 20])
  } else if (ncol(x) >= 16){
    term <- as.character(x[1, 14])
  } else if (ncol(x) <= 15){
    term <- as.character(x[1, 8])
  } else {
    term <- " - "
  }
  term <- unlist(strsplit(term, split = " - "))
  date.in  <- term[1]
  date.out <- term[2]
  string <- cbind(camp, date.in, date.out)
}

# 15-дневные заезды
TrmByDayFam15 <- function(x){
  if (ncol(x) == 17) {
    trm.con  <- as.numeric(x[9, 2:17])
    tbl.day <- data.frame(rbind(trm.con))
  } else {
    tbl.day <- data.frame(rbind(rep(NA, 16)))
  }
  return(tbl.day)
}

# 14-дневные заезды
TrmByDayFam14 <- function(x){
  if (ncol(x) == 16) {
    trm.con  <- as.numeric(x[9, 2:16])
    tbl.day <- data.frame(rbind(trm.con))
  } else {
    tbl.day <- data.frame(rbind(rep(NA, 15)))
  }
  return(tbl.day)
}

# Подгрузить дополнительные данные -------------------------------------
data.arrivals <- get(load("~/data/data_fam_2019.rda"))
camps <- read.csv2("~/data/camps.csv")

# Сборка массива -------------------------------------------------------
# Рабочая папка
setwd("~/aism/2019/arrivals_fam/")
# Обработать файлы заездов
file.names <- list.files(path = "./", recursive = TRUE, 
                         pattern = "*.xlsx")
file.list <- lapply(file.names, read.xlsx)
# Создать подмассив со сведениями о заездах
session.data <- lapply(file.list, GetInfo)
session.data <- data.frame(matrix(unlist(session.data), 
                                  nrow=length(session.data), byrow=TRUE))

# 15-дневные заезды ----------------------------------------------------
# Создать подмассив со сведениями об обращениях
by.day15.data <- lapply(file.list, TrmByDayFam15)
by.day15.data <- data.frame(matrix(unlist(by.day15.data), 
                                   nrow=length(by.day15.data), byrow=TRUE))

# Объединить массивы
by.day15.data <- cbind(session.data, by.day15.data)
colnames(by.day15.data) <- c("camp_name", "date_in", "date_out",
                             as.character(1:15), "trm_sum")

# Обработка дат
by.day15.data$date_in <- as.Date(by.day15.data$date_in, 
                                 format = "%d.%m.%Y")
by.day15.data$date_out <- as.Date(by.day15.data$date_out, 
                                  format = "%d.%m.%Y")

# Задать расположение лагеря
by.day15.data$region <- camps$region[match(by.day15.data$camp_name, 
                                           camps$camp_name)]

# Удаление дубликатов
by.day15.data <- unique(by.day15.data)

# Добавить количество отдыхающих
by.day15.data$visitors <- fam2019$visits_total

# Удаление NA
by.day15.data <- na.omit(by.day15.data)

# 14-дневные заезды ----------------------------------------------------
# Создать подмассив со сведениями об обращениях
by.day14.data <- lapply(file.list, TrmByDayFam14)
by.day14.data <- data.frame(matrix(unlist(by.day14.data), 
                                   nrow=length(by.day14.data), byrow=TRUE))

# Объединить массивы
by.day14.data <- cbind(session.data, by.day14.data)
colnames(by.day14.data) <- c("camp_name", "date_in", "date_out",
                             as.character(1:14), "trm_sum")

# Обработка дат
by.day14.data$date_in <- as.Date(by.day14.data$date_in, 
                                 format = "%d.%m.%Y")
by.day14.data$date_out <- as.Date(by.day14.data$date_out, 
                                  format = "%d.%m.%Y")

# Задать расположение лагеря
by.day14.data$region <- camps$region[match(by.day14.data$camp_name, 
                                           camps$camp_name)]

# Удаление дубликатов
by.day14.data <- unique(by.day14.data)

# Добавить количество отдыхающих
by.day14.data$visitors <- fam2019$visits_total

# Удаление NA
by.day14.data <- na.omit(by.day14.data)

# Переименование результирующей переменной -----------------------------
by.day14.data -> trm.daily14.fam2019
by.day15.data -> trm.daily15.fam2019
rm(list = setdiff(ls(), c("trm.daily14.fam2019", "trm.daily15.fam2019")))

# Сохранение результатов -----------------------------------------------
save(trm.daily14.fam2019, file = "~/data/trm_daily14_fam2019")
save(trm.daily14.fam2019, file = "~/data/trm_daily15_fam2019")
