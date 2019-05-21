pkgs <- c("rio","readxl","jiebaR","ngram","dplyr","tidyverse","parallel","taRifx","stringr","Nippon")
pkgs_ins <- pkgs [!(pkgs %in% installed.packages()[,"Package"])]
if( length(pkgs_ins)>0 )
{
  for (i in 1:length(pkgs_ins))
  {
    install.packages(pkgs_ins[i])
  }
}

library(rio)
library(readxl)
library(jiebaR)
library(ngram)
library(tidyverse)
library(parallel)
library(taRifx)
library(stringr)
library(Nippon)

#data_set
data.name <- "RANDRANDOMSAMPLESTATEMENT_DET.xlsx"

#data_import
data.raw <- read_excel(data.name)[1:10000,]
data.raw <- memo
col.analysis <- "Contact_Memo"

#啟動結巴
cutter <- worker(bylines = TRUE)

#斷詞準備
data.content <- data.raw[,col.analysis] %>% unlist()

fth <- function(x){zen2han(as.character(x))} #FulltoHalf
bigger <- function(x){toupper(x)} #轉大寫 
space <- function(x){gsub("　",replacement="",x)}
data.content <- sapply(data.content,fth)
data.content <- sapply(data.content,bigger)
data.content <- sapply(data.content,space)

#進行斷詞
data.content.cutter <- cutter[data.content]
data.content.con <- concatenate(data.content.cutter)

#ngram進行1~n個詞單位的斷詞
ng.sum <- function(article_con = data.content.con,freq_ngrams_num = 5,ngram_num = 3)
{
  ng_table_sum <- as.character()
  pb_j <- txtProgressBar(min = 1, max = ngram_num, style = 3)
  for (j in 1:ngram_num)
  {
    ng_words <- ngram(article_con,n=j)
    ng_table <- get.phrasetable(ng_words)
    if(nrow(ng_table)>0)
    {
      Encoding(ng_table[,1]) <- 'UTF-8'
      ng_table[,1] <- gsub("c\\(",replacement = "",ng_table[,1])
      ng_table[,1] <- gsub("[[:punct:]]|[[:cntrl:]]",replacement = "",ng_table[,1])
      ng_table[,1] <- gsub(" ",replacement = "",ng_table[,1]) #取代其他符號
      ng_table <- aggregate(freq ~ ngrams,ng_table,sum)#把符號去掉後，將相同的詞加回去
      ng_table$ng_n <- j
      ng_table$wordlen <- nchar(ng_table$ngrams)
      ng_table <- ng_table %>% as_tibble()
      ng_table1 <- ng_table %>% 
        filter(freq > freq_ngrams_num) %>% 
        filter(wordlen > 1) #保留詞長>2
        #filter(wordlen < wordlen_num)
      ng_table_sum <- rbind(ng_table_sum,ng_table1)
    }
    setTxtProgressBar(pb_j, j)
  }
  ng_table_sum <- arrange(ng_table_sum, desc(freq))
  ng_table_sum  <- ng_table_sum %>% group_by(ngrams) %>% top_n(1,freq) %>% ungroup()
  return(ng_table_sum)
}

ng.sum1 <- ng.sum(data.content.con,100000,3)
ng.unique <- ng.sum1["ngrams"]

#中英文標記調整
fun_dim_nonc <- function(x){grepl("^[A-Za-z0-9]+$",x)}
label_notc <- "[\u4e00-\u9fa5| ]subname[\u4e00-\u9fa5| ]|^subname[\u4e00-\u9fa5| ]|[\u4e00-\u9fa5| ]subname$|^subname$"
fun_notc_sub <- function(x){gsub("subname",x,label_notc)}
dim.label <- function(ng_table_unique=ng_unique)
{
  #啟動多核心
  cl_num <- detectCores()
  cl <- makeCluster(cl_num - 1)
  
  #標記純數字英文的文章
  ng_table_dim_notc_label <- as.character()
  clusterExport(cl,"fun_dim_nonc")
  ng_table_dim_notc_label <- parLapply(cl,ng_table_unique[,1],fun_dim_nonc) %>% unlist()
  ng_table_dim_notc_label <- sub("TRUE",replacement = 1, ng_table_dim_notc_label)
  ng_table_dim_notc_label <- sub("FALSE",replacement = 0, ng_table_dim_notc_label)
  ng_table_dim_notc_label <- ng_table_dim_notc_label %>% as.integer() %>% as.data.frame()
  colnames(ng_table_dim_notc_label) <- "notchinese"
  ng_table_sum2 <- cbind(ng_table_unique,ng_table_dim_notc_label)

  #分別歸為純數字英文及非純數字英文
  ng_table_dim_notc <- ng_table_sum2[ng_table_sum2$notchinese == 1,] %>% remove.factors()
  ng_table_dim_isc <- ng_table_sum2[ng_table_sum2$notchinese == 0,] %>% remove.factors()
  
  #將純數字英文進行關鍵字正規處理整理
  if (is_empty(ng_table_dim_notc[,1])==FALSE)
  {
    clusterExport(cl,"label_notc")
    data_notc_sub <- parLapply(cl,ng_table_dim_notc[,1],fun_notc_sub) %>% unlist() %>% tibble()
    colnames(data_notc_sub) <- "keyword"
    ng_table_dim_notc <- cbind(ng_table_dim_notc,data_notc_sub)
  }
  
  #將非純數字英文進行關鍵字整理
  ng_table_dim_isc_sub <- ng_table_dim_isc[,1] %>% unlist() %>% as.data.frame()
  colnames(ng_table_dim_isc_sub) <- "keyword"
  ng_table_dim_isc <- cbind(ng_table_dim_isc,ng_table_dim_isc_sub)

  #將上述兩個檔案合併
  ng_table_dimension <- rbind(ng_table_dim_notc,ng_table_dim_isc) %>% remove.factors()
  stopCluster(cl)
  return(ng_table_dimension)
}

ng_table_dimension <- dim.label(ng.unique)

#啟動多核心
cl_num <- detectCores()
cl <- makeCluster(cl_num - 1)

#詞頻和文章數統計表單建立
dimension_data_doc <- dimension_data_freq <- 1:nrow(data.raw) %>% tibble()
content_labeled <- data.content

#進行標記
ng_count <- as.character()
pb_l <- txtProgressBar(min = 1, max = nrow(ng_table_dimension), style = 3)
for(l in 1:nrow(ng_table_dimension))
{
  fun_freq <- function(x){str_count(x,ng_table_dimension[l,3])}
  clusterExport(cl,"l")
  clusterExport(cl,"str_count")
  clusterExport(cl,"ng_table_dimension")
  clusterExport(cl,"fun_freq")
  
  freq_num <- parLapply(cl,content_labeled,fun_freq) %>% unlist() %>% as.data.frame()
  freq_num[is.na(freq_num)==TRUE,] <- 0
  freq_sum <- sum(freq_num) %>% tibble()
  dimension_data_freq <- cbind(dimension_data_freq,freq_num)
  
  doc_num <- freq_num
  doc_num[doc_num > 1,] <-1
  doc_sum <- sum(doc_num) %>% tibble()
  dimension_data_doc <- cbind(dimension_data_doc,doc_num)
  
  ng_freq_doc <- cbind(ng_table_dimension[l,1],freq_sum,doc_sum) %>% as.data.frame()
  ng_count <- rbind(ng_count,ng_freq_doc)
  setTxtProgressBar(pb_l, l)
}
stopCluster(cl)
rm(fun_freq,freq_num,freq_sum,doc_num,doc_sum,ng_freq_doc,pb_l)

#欄位命名
ng_count[,2:3] <- ng_count[,2:3] %>% unlist() %>% as.integer()
colnames(ng_count) <- c("ngram","freq","doc")

#進行篩選
ng_count_filter <- ng_count %>% filter(freq > 100000 & doc > 100000) %>% arrange(desc(doc)) %>% remove.factors()
ng_count_filter$filter <- 0

#詞頻數和文章數_欄位命名
t_dimension<-as.character(t(ng_table_dimension[ ,1]))
colnames(dimension_data_freq)<-c("NO.",t_dimension)
colnames(dimension_data_doc)<-c("NO.",t_dimension)

filter_token <- ng_count_filter[,1] %>% t() %>% as.list
sav <- colnames(dimension_data_freq) %in% filter_token
dimension_data_freq <- cbind(dimension_data_freq[,1],dimension_data_freq[,sav])
dimension_data_doc <- cbind(dimension_data_doc[,1],dimension_data_doc[,sav])
colnames(dimension_data_freq)[1] <- "NO."
colnames(dimension_data_doc)[1] <- "NO."

##output
export(ng_count_filter,"words_stat.xlsx")
#dimension_freq <- paste0(file_name,"_dimension_freq.csv")
write.csv(dimension_data_freq,file="dimension_freq.csv",row.names = FALSE)
#dimension_doc <- paste0(file_name,"_dimension_doc.csv")
write.csv(dimension_data_doc,file="dimension_doc.csv",row.names = FALSE)