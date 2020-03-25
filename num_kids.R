EstimateKids <- function(x){
  x$id1 <- paste(x$IDser_applicant, x$IDno_applicant, sep = " ")
  x$id2 <- paste(x$IDser_camper, x$IDno_camper, sep = " ")
  set1 <- x %>%
    filter(age_camper < 18) %>%
    distinct(id2, .keep_all = TRUE) %>%
    count(id1)
  x$n <- set1$n[match(x$id1, set1$id1)]
  x1 <- x %>% filter(is.na(n) == FALSE)
  x2 <- x %>% filter(is.na(n) == TRUE)
  set2 <- x2 %>% 
    filter(age_camper < 18) %>%
    distinct(id2, .keep_all = TRUE) %>%
    count(id1)
  x2$n <- set2$n[match(x2$id1, set2$id1)]
  x <- rbind(x1, x2)
  return(x)
}