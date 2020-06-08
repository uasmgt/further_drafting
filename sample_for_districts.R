library(dplyr)
library(openxlsx)

setwd("~/aiso/survey_sample/")
dir.names <- grep("ao", list.dirs(full.names = FALSE), value = TRUE)
districts <- c("Центральный", "Новомосковский", "Северный", 
               "Северо-Восточный", "Северо-Западный", "Троицкий",
               "Восточный", "Южный", "Юго-Восточный", "Юго-Западный",
               "Западный", "Зеленоградский")

CreateDataset <- function(x) {
  files <- list.files(path = x, pattern = ".xlsx", recursive = TRUE)
  files.list <- lapply(files, read.xlsx)
  dataset <- lapply(files.list, CreateTable)
  y <- do.call(rbind.data.frame, dataset)
  names(y) <- c("status", "email")
  y <- y %>% 
    filter(status == "Услуга оказана" | 
             status == "Отказ в предоставлении услуги")
  emails <- unique(y$email)
  return(emails)
}
CreateTable <- function(x)
{
  names(x) <- as.character(unlist(x[1, ]))
  x <- x[-c(1), c(5, 44)]
}
CreateList <- function(x) {
  setwd(x)
  dataset <- CreateDataset("./")
  setwd("..")
  return(dataset)
}

sample.base <- lapply(dir.names, CreateList)
names(sample.base) <- dir.names

districts.table <- data.frame(districts)
districts.table$emails <- as.numeric(lapply(sample.base, length))
districts.table$per <- districts.table$emails / sum(districts.table$emails) * 100

# параметры выборки
td = 1.96 # для уровня доверия в 95%
p = .5 # доля людей, обладающих исследуемым признаком в совокупности
ci = .05 # доверительный интервал
population <- sum(districts.table$emails) # размер генеральной совокупности
rr = 0.1 # ожидаемая доля отклика
sample.size <- round(1 + (1 + td^2 * p * (1 - p) / ci ^ 2) / 
  (1 + (td^2 * p * (1 - p) / ci ^ 2 / population)) / rr, 0)

districts.table$sample <- round(districts.table$per * sample.size / 100, 0)


Sampler <- function(x){
  base <- x
  size <- round(length(x) / population * sample.size, 0)
  sampled <- sample(base, size)
  return(sampled)
}

sampled <- lapply(sample.base, Sampler)
sampled <- unique(unlist(sampled))
write.csv(sampled, "~/Documents/work/Исследование - качество услуг и конкуренция/sample.csv",
          row.names = FALSE)
