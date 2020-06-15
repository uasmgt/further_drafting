# библиотеки
library(openxlsx)
library(stringr)

# указать правильное название файла и путь до него
emails <- read.xlsx("input.xlsx", rowNames = FALSE)
expression <- c("([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+.[a-zA-Z0-9_-]+)")
#вместо "emails$string" указать нужную колонку таблицы
emails$email <- str_match(emails$string, expression) 

write.xlsx(emails, "output.xlsx")
