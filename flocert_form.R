library(dplyr)
library(readxl)
library(tidyverse)


x <- read_excel("~/Documents/CartONG/XLSForms/Fair Trade/FLOCERT_form_PDF.xlsx", 
                sheet = "Table 1", col_types = "text", col_names = TRUE)

x$Time <- as.character(x$Time)
x$Reference <- as.character(x$Reference)


for (i in 2:81){
  tempo <- read_excel("~/Documents/CartONG/XLSForms/Fair Trade/FLOCERT_form_PDF.xlsx",
                      sheet = paste("Table", i))
  
  tempo$Time <- as.character(tempo$Time)
  tempo$Reference <- as.character(tempo$Reference)
  
  x <- bind_rows(x, tempo)       

}
view(x)
ncol(x)
write_csv(x, "~/Documents/CartONG/XLSForms/Fair Trade/FLOCERT_pdf_to_xls.csv")

y <- x %>% 
  select(c(3, 5,6,7,8,9))
y[1,1] <- "empty"

for (r in 1:nrow(y)){
  for (c in 2:6){
    if (is.na(y[r,1]))
      {t <- y[r,c] #copy value col 2 row l
      collector <- 0 #determine how many rows above we need to paste the value
      while (is.na(y[r-collector,1])){ # while the value of the question column in the lines above is empty:
             collector <- collector + 1 # add 1 to the collector
             }
           y[r-collector,c] <- paste(y[r-collector,c], y[r,c], sep = "_" )
    }
  }
  }


view(y)

z <- y %>% 
  filter(question != "question") %>% 
  gather(key="rank", value="rank_text", -question) %>% 
  arrange(question) %>% 
  mutate(value = substr(rank, nchar(rank), nchar(rank)),
         list_name = paste("s", gsub(".","_",question, fixed = T),sep =""))

z$rank_text <- str_replace_all(z$rank_text, "_", " ")
view(z)

write_csv(z, "~/Documents/CartONG/XLSForms/Fair Trade/FLOCERT_form_csv.csv")



# extract criteria column -------------------------------------------------
view(x)

yy <- x %>% 
  select(c(3, 5,6,7,8,9,11))
yy[1,1] <- "empty"

for (r in 1:nrow(yy)){
  for (c in 2:7){
    if (is.na(yy[r,1]))
    {t <- yy[r,c] #copy value col 2 row l
    collector <- 0 #determine how many rows above we need to paste the value
    while (is.na(yy[r-collector,1])){ # while the value of the question column in the lines above is empty:
      collector <- collector + 1 # add 1 to the collector
    }
    yy[r-collector,c] <- paste(yy[r-collector,c], yy[r,c], sep = "_" )
    }
  }
}

yy <- yy %>% 
  rename(criteria_type = `Criteria£Type`)

zz <- yy %>% 
  filter(question != "question") %>% 
  filter(!is.na(criteria_type ))

index_criteria <- substr(zz$criteria_type,1,1)
question <- zz$question

zz <- data.frame(question, index_criteria)

zz <- zz %>%
  mutate(list_name = paste("s", gsub(".","_",question, fixed = T),sep ="")) %>% 
  select(list_name, index_criteria) %>% 
  mutate(denominator = paste("if(${", list_name, "}!='',1,0)", sep = ''),
         numerator = paste("if(${", list_name, "}!='', ${", list_name, "},0)", sep = ''),
         section = substr(list_name, 2,2))


table(zz$index_criteria, zz$section)  


# Global development average 
 
zz[1,3]

d_glb_denominator <- NULL

for(l in 1:nrow(zz)){
  if (zz[l,2] == 'D') 
    {d_glb_denominator <- paste(d_glb_denominator, zz[l,3],' + ',sep = '')
    
  }
}

d_glb_numerator <- NULL

for(l in 1:nrow(zz)){
  if (zz[l,2] == 'D') 
  {d_glb_numerator <- paste(d_glb_numerator, zz[l,4],' + ',sep = '')
  
  }
}

#section 3

d_3_denominator <- NULL

for(l in 1:nrow(zz)){
  if (zz[l,2] == 'D' & zz[l, 5] == "3") 
  {d_3_denominator <- paste(d_3_denominator, zz[l,3],' + ',sep = '')
  }
}

d_3_numerator <- NULL

for(l in 1:nrow(zz)){
  if (zz[l,2] == 'D' & zz[l, 5] == 3) 
  {d_3_numerator <- paste(d_3_numerator, zz[l,4],' + ',sep = '')
  
  }
}

#section 4

d_4_denominator <- NULL

for(l in 1:nrow(zz)){
  if (zz[l,2] == 'D' & zz[l, 5] == "3") 
  {d_4_denominator <- paste(d_4_denominator, zz[l,3],' + ',sep = '')
  }
}

d_4_numerator <- NULL

for(l in 1:nrow(zz)){
  if (zz[l,2] == 'D' & zz[l, 5] == 3) 
  {d_4_numerator <- paste(d_4_numerator, zz[l,4],' + ',sep = '')
  
  }
}




write_csv(zz, "~/Documents/CartONG/XLSForms/Fair Trade/FLOCERT_index_critria.csv")
 



