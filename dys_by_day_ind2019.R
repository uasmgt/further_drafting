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

# Обращения из-за заболеваний по дням
DysByDayInd <- function(x){
  if (ncol(x) == 23) {
    dys.con  <- as.numeric(x[28,  2:23])
    tbl.day <- data.frame(rbind(dys.con))
  } else {
    tbl.day <- data.frame(rbind(rep(NA, 22)))
  }
  return(tbl.day)
}

# Подгрузить дополнительные данные -------------------------------------
data.arrivals <- get(load("~/data/data_ind_2019.rda"))
camps <- read.csv2("~/data/camps.csv")

# Сборка массива -------------------------------------------------------
# Рабочая папка
setwd("~/aism/2019/arrivals_ind/")
# Обработать файлы заездов
file.names <- list.files(path = "./", recursive = TRUE, 
                         pattern = "*.xlsx")
file.list <- lapply(file.names, read.xlsx)
# Создать подмассив со сведениями о заездах
session.data <- lapply(file.list, GetInfo)
session.data <- data.frame(matrix(unlist(session.data), 
                                  nrow=length(session.data), byrow=TRUE))
# Создать подмассив со сведениями об обращениях
by.day.data <- lapply(file.list, DysByDayInd)
by.day.data <- data.frame(matrix(unlist(by.day.data), 
                                 nrow=length(by.day.data), byrow=TRUE))

# Объединить массивы
by.day.data <- cbind(session.data, by.day.data)
colnames(by.day.data) <- c("camp_name", "date_in", "date_out",
                           as.character(1:21), "dys_sum")

# Обработка дат
by.day.data$date_in <- as.Date(by.day.data$date_in, 
                                format = "%d.%m.%Y")
by.day.data$date_out <- as.Date(by.day.data$date_out, 
                                 format = "%d.%m.%Y")

# Задать расположение лагеря
by.day.data$region <- camps$region[match(by.day.data$camp_name, 
                                          camps$camp_name)]

# Удаление дубликатов
by.day.data <- unique(by.day.data)

# Добавить количество отдыхающих
by.day.data$kids <- ind2019$fact_total

# Удаление NA
by.day.data <- na.omit(by.day.data)

# Переименование результирующей переменной -----------------------------
by.day.data -> dys.daily.ind2019
rm(list = setdiff(ls(), "dys.daily.ind2019"))

# Сохранение результатов -----------------------------------------------
save(dys.daily.ind2019, file = "~/data/dys_daily_ind2019")
