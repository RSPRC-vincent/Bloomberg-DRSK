########## Sigma兩種算法：(1)年度最高點與最低點(2)以股票日報酬率波動率換算成年化波動率

library(CreditRisk)
library(kableExtra)
library(dplyr)
library(readxl)
library(writexl)
library(xlsx)
library(tidyr)
library(utils)
library(data.table)
library(stats)
library(plyr)

getwd()
setwd("D:/Drivers/台大風險中心/14.模型與指標/Merton違約機率")
setwd("/Users/Galong/Insync/Drivers/台大風險中心/14.模型與指標/Merton違約機率")

############################################################################## (1)Sigma用年度最高點與最低點計算。
# 26家南科電子業廠商違約機率(以2020起算每五年的存活機率到2050年)
# L和Vo單位為千元, 資料來源:TEJ IFRS以合併為主財務(累計)_一般行業IV
# Sigma:log(sqrt(股價最高點/股價最低點))，股價資料來源：TEJ股價資料庫-TEJ調整股價(年)-除權息調整
# r為目前政府十年公債利率的資料來源：中華民國統計資訊網https://statdb.dgbas.gov.tw/pxweb/Dialog/varval.asp?ma=FM3007A1A&ti=%A7Q%B2v%B2%CE%ADp-%B8%EA%A5%BB%A5%AB%B3%F5%A7Q%B2v-%A6%7E&path=../PXfile/FinancialStatistics/&lang=9&strList=L
# 碳排量為TEJ CSR資料庫的碳排放量（範疇一加範疇二排放量，單位公噸）
stock_price <- read_xlsx("rawdata/2020TEJ未調整股價(年).xlsx",skip = 2)
IFRS <- read_xlsx("rawdata/2020TEJ_IFRS以合併為主財務(累計)_一般行業IV(all).xlsx",skip = 2)
SP26 <- read_xlsx("rawdata/26家南科製造業.xlsx")
carbon <- read_xlsx("rawdata/2014-2020_all公司碳排放量.xlsx", skip = 2) #TEJ CSR資料庫
rate <- read_xlsx("rawdata/1997-2021利率統計-資本市場利率-年.xlsx") #需要手動修改.csv檔再轉成.xlsx檔

# select rate for 2020
rate2020 <- rate[which(rate$年 == 2020), 5] %>% as.numeric()
rate2020

# create stock
stock_price$stock <- substr(stock_price$證券代碼, regexpr("[0-9A-Z]*", stock_price$證券代碼), attr(regexpr("[0-9A-Z]*", stock_price$證券代碼), "match.length"))
IFRS$stock <- substr(IFRS$公司, regexpr("[0-9A-Z]*", IFRS$公司), attr(regexpr("[0-9A-Z]*", IFRS$公司), "match.length"))
SP26$stock <- substr(SP26$name, regexpr("[0-9A-Z]*", SP26$name), attr(regexpr("[0-9A-Z]*", SP26$name), "match.length"))
carbon$stock <- substr(carbon$公司代碼, regexpr("[0-9A-Z]*", carbon$公司代碼), attr(regexpr("[0-9A-Z]*", carbon$公司代碼), "match.length"))

# select needed columns
stock_price_s <- stock_price[, c(30,4,5)]
IFRS_s <- IFRS[, c("stock", "公司", "資產總額", "負債總額")]
SP26_s <- SP26[, c(8,1)]
carbon_s <- carbon %>% filter(報導年度== "2020" & 最新揭露註記 == "Y")

# combine SP26 & carbon_s
carbon26 <- SP26_s %>% 
  left_join(carbon_s, by = c("stock"="stock"))

carbon26 <- carbon26[-c(3,13,15,27), ] # 非台灣廠區
carbon26 <- carbon26[-c(18,24), ] # 無資產負債排放量資料的公司
colnames(carbon26)[12] <- "碳排放總量"


carbon26_s <- ddply(carbon26,~stock,summarise,TWCO2 = sum(as.numeric(碳排放總量)))
            
# combine data
data26 <- SP26_s %>% 
  left_join(IFRS_s, by = c("stock" = "stock")) %>%
  left_join(stock_price_s, by = c("stock" = "stock")) %>%
  left_join(carbon26_s, by = c("stock" = "stock"))

# remove any row with NA
data_f <- data26[complete.cases(data26), ] #3267泰陞、2366亞旭、6244茂迪

# as.numeric columns
data_f$`最高價(元)_年` <- as.numeric(data_f$`最高價(元)_年`)
data_f$`最低價(元)_年` <- as.numeric(data_f$`最低價(元)_年`)

# model 23家南科廠商違約機率
firm.merge <- data.frame()

for (i in 1:nrow(data_f)) {
  
  for (j in 0:5) {
    
    CO2 <- data_f[i,8] %>% as.numeric()
    
    CO2cost <- CO2*j
    
    Liability <- data_f[i,5] %>% as.numeric()
    
    Value <- data_f[i,4] %>% as.numeric()
    
    Sigma <- log(sqrt(data_f[i,6]/data_f[i,7])) %>% as.numeric()
    
    mod <- Merton(L = Liability + CO2cost, V0 = Value, sigma = Sigma, r = rate2020,
                  t = c(5.00, 10.00,15.00,20.00,25.00,30.00))
    
    # 新增公司名
    mod$name <- data_f[i,3] %>% as.character() 
    
    # 新增到期年
    mod$Year <- 2020 + mod$Maturity
    
    # 新增碳費
    mod$carbon_price <- j*1000
    
    firm.merge <- rbind(firm.merge, mod)
    
  }
}

firm.merge <- firm.merge[, c(6:8,1:5)]

# save files
write.csv(firm.merge, "output/23家南科電子廠商存活機率.csv", row.names = F)
xlsx::write.xlsx(firm.merge, "output/23家南科電子廠商存活機率.xlsx", showNA = F, row.names = F)

# spread firm.merge
firm.merge_y <- firm.merge %>% dplyr::select(., c(1:4,8)) %>% spread(., carbon_price, Survival)
firm.merge_plot <- firm.merge_y[,-3]

# save files
write.csv(firm.merge_plot, "output/23家南科電子廠商存活機率(製圖用).csv", row.names = F)
xlsx::write.xlsx(firm.merge_plot, "output/23家南科電子廠商存活機率(製圖用).xlsx", showNA = F, row.names = F)

write.csv(data26, "output/26家Merton基本資料.csv", row.names = F)
#xlsx::write.xlsx(data26, "output/26家Merton基本資料.xlsx", showNA = T, row.names = F)
writexl::write_xlsx(data26, "output/26家Merton基本資料.xlsx")


############################################################ 畫圖
library(ggplot2)

# read file
table <- read.table("output/23家南科電子廠商存活機率(製圖用).csv", header=T, sep=",", row.names = NULL, stringsAsFactors = F)

# select columns
table_s <- table[, c(1,2,5)]
colnames(table_s)[3] <- "Survival Rate"

# create No for name
table_s$No <- ""

j = 1
for (i in seq(1, 133, by = 6)){
  table_s[c(i:(i+5)),4] <- j
  j = j + 1
}

firmlist <- unique(table_s$name)

# make plot for each firm
for (i in 1:23){
  png(file = paste0("output/figure/", firmlist[i], ".png"), width = 600, height = 400)
  lineplot <- ggplot(table_s[which(table_s$No == i),], aes(x = Year, y = `Survival Rate`, group = 1)) + 
    geom_line() + 
    geom_point()+
    ggtitle(paste(firmlist[i], "在碳費2000元的存活機率"))+
    ylim(0, 1.00)+
    theme(legend.position="top")+
    theme(text=element_text(family="Taipei Sans TC Beta", size=10))
  print(lineplot)
  dev.off()
}

# 23 firms in one plot
png(file = "output/figure/23家南科電子廠商存活機率.png", width = 600, height = 400)
lineplot2 <- ggplot(table_s, aes(x = Year, y = `Survival Rate`, group = name)) + 
  geom_line(aes(color=name)) + 
  geom_point()+
  ylim(0, 1.00)+
  theme(legend.position="bottom")+
  theme(text=element_text(family="Taipei Sans TC Beta", size=10))
print(lineplot2)
dev.off()

############################################################################ (2)Sigma用股票日報酬率波動性轉為年化波動性
# 26家南科電子業廠商違約機率(以2020起算每五年的存活機率到2050年)
# L和Vo單位為千元, 資料來源:TEJ IFRS以合併為主財務(累計)_一般行業IV
# Sigma:年股票報酬率波動性=股票日報酬率波動性* sqrt(T- day) 。股價資料來源：TEJ股價資料庫-TEJ調整股價(日)-除權息調整
# r為目前政府十年公債利率的資料來源：中華民國統計資訊網https://statdb.dgbas.gov.tw/pxweb/Dialog/varval.asp?ma=FM3007A1A&ti=%A7Q%B2v%B2%CE%ADp-%B8%EA%A5%BB%A5%AB%B3%F5%A7Q%B2v-%A6%7E&path=../PXfile/FinancialStatistics/&lang=9&strList=L
# 碳排量為TEJ CSR資料庫的碳排放量（範疇一加範疇二排放量，單位公噸）
stock_price1 <- fread("rawdata/2020TEJ全部普通股日收盤價.csv", encoding="UTF-8") #先把檔案另存成UTF8檔再讀入
IFRS <- read_xlsx("rawdata/2020TEJ_IFRS以合併為主財務(累計)_一般行業IV(all).xlsx",skip = 2)
SP26 <- read_xlsx("rawdata/26家南科製造業.xlsx")
carbon <- read_xlsx("rawdata/2014-2020_all公司碳排放量.xlsx", skip = 2) #TEJ CSR資料庫
rate <- read_xlsx("rawdata/1997-2021利率統計-資本市場利率-年.xlsx") #需要手動修改.csv檔再轉成.xlsx檔

# select rate for 2020
rate2020 <- rate[which(rate$年 == 2020), 5] %>% as.numeric()
rate2020 #0.427

# create stock
colnames(stock_price1)[1] <- "stock"
stock_price1$stock <- as.character(stock_price1$stock)
#unique(nchar(stock_price1$年月日))
IFRS$stock <- substr(IFRS$公司, regexpr("[0-9A-Z]*", IFRS$公司), attr(regexpr("[0-9A-Z]*", IFRS$公司), "match.length"))
SP26$stock <- substr(SP26$name, regexpr("[0-9A-Z]*", SP26$name), attr(regexpr("[0-9A-Z]*", SP26$name), "match.length"))
carbon$stock <- substr(carbon$公司代碼, regexpr("[0-9A-Z]*", carbon$公司代碼), attr(regexpr("[0-9A-Z]*", carbon$公司代碼), "match.length"))

# select needed columns
IFRS_s <- IFRS[, c("stock", "公司", "資產總額", "負債總額")]
SP26_s <- SP26[, c(8,1)]
carbon_s <- carbon %>% filter(報導年度== "2020" & 最新揭露註記 == "Y")
stock_price1_s <- stock_price1[, c(1:3,6)]
stock_price1_s <- SP26_s %>%
  left_join(stock_price1_s, by = c("stock" = "stock"))


# combine SP26 & carbon_s
carbon26 <- SP26_s %>% 
  left_join(carbon_s, by = c("stock"="stock"))

carbon26 <- carbon26[-c(3,13,15,27), ] # 非台灣廠區
colnames(carbon26)[12] <- "碳排放總量"
# remove 碳排放總量 is NA
carbon26 <- carbon26[-which(is.na(carbon26$碳排放總量)), ] #remove 6244茂迪、3267泰陞、2366亞旭
carbon26_s <- ddply(carbon26,~stock,summarise,TWCO2 = sum(as.numeric(碳排放總量)))


# check all firms have stock price
nostock <- stock_price1_s[which(is.na(stock_price1_s$`收盤價(元)`)), ] # 代碼3267, 2366無2020日股票收盤價
stock_price1_s <- stock_price1_s[-which(is.na(stock_price1_s$`收盤價(元)`)), ] ## 只剩24家有日股票

#stockdays <- stock_price1_s %>% group_by(stock) %>% plyr::count(., 'stock') #2448晶電只有239交易日，非245日
#stock_price1_s2448 <- stock_price1_s[which(stock_price1_s$stock == "2448"), ] #2448晶電
#stock_price1_s <- stock_price1_s[-which(stock_price1_s$stock == "2448"), ] ## 移除2448晶電,只剩23家有日股票

# calculate Sigma for each firm
stock_price1_s$LnStock <- "" %>% as.numeric()

for (i in 1:(nrow(stock_price1_s)-1)){
    stock_price1_s$LnStock[i+1] <- log(stock_price1_s[(i+1),5]/stock_price1_s[i,5]) %>% unlist() 
}

#stock_price1_s1 <- stock_price1_s
#stock_price1_s <- stock_price1_s1
stock_price1_s[!duplicated(stock_price1_s$stock), 6] <- NA
a <- stock_price1_s[which(is.na(stock_price1_s1$LnStock)), ]

# calculate Sigma (group by stock)
df <- ddply(stock_price1_s,~stock,summarise,mean=mean(LnStock, na.rm = T),sd=sd(LnStock, na.rm = T))
df$Sigma0 <- df$sd*sqrt(245) #245個交易日

# combine data
data26 <- SP26_s %>% 
  left_join(IFRS_s, by = c("stock" = "stock")) %>%
  left_join(df, by = c("stock" = "stock")) %>%
  left_join(carbon26_s, by = c("stock" = "stock"))

# remove any row with NA
data_f <- data26[complete.cases(data26), ] #3267泰陞、2366亞旭、6244茂迪

data_f <- data_f[, -6]

# model 23家南科廠商違約機率
firm.merge <- data.frame()

for (i in 1:nrow(data_f)) {
  
  for (j in 0:5) {
    
    CO2 <- data_f[i,8] %>% as.numeric()
    
    CO2cost <- CO2*j
    
    Liability <- data_f[i,5] %>% as.numeric()
    
    Value <- data_f[i,4] %>% as.numeric()
    
    Sigma <- data_f[i,7] %>% as.numeric()
    
    mod <- Merton(L = Liability + CO2cost, V0 = Value, sigma = Sigma, r = rate2020,
                  t = c(5.00, 10.00,15.00,20.00,25.00,30.00))
    
    # 新增公司名
    mod$name <- data_f[i,3] %>% as.character() 
    
    # 新增到期年
    mod$Year <- 2020 + mod$Maturity
    
    # 新增碳費
    mod$carbon_price <- j*1000
    
    firm.merge <- rbind(firm.merge, mod)
    
  }
}

firm.merge <- firm.merge[, c(6:8,1:5)]

# save files
write.csv(firm.merge, "output/23家南科電子廠商存活機率(日波動率).csv", row.names = F)
xlsx::write.xlsx(firm.merge, "output/23家南科電子廠商存活機率(日波動率).xlsx", showNA = F, row.names = F)

# spread firm.merge
firm.merge_y <- firm.merge %>% dplyr::select(., c(1:4,8)) %>% spread(., carbon_price, Survival)
firm.merge_plot <- firm.merge_y[,-3]

# save files
write.csv(firm.merge_plot, "output/23家南科電子廠商存活機率(製圖用)(日波動率).csv", row.names = F)
xlsx::write.xlsx(firm.merge_plot, "output/23家南科電子廠商存活機率(製圖用)(日波動率).xlsx", showNA = F, row.names = F)

write.csv(data26, "output/26家Merton基本資料(日波動率).csv", row.names = F)
#xlsx::write.xlsx(data26, "output/26家Merton基本資料.xlsx", showNA = T, row.names = F)
writexl::write_xlsx(data26, "output/26家Merton基本資料(日波動率).xlsx")

############################################################ 畫圖
# read file
table <- read.table("output/23家南科電子廠商存活機率(製圖用)(日波動率).csv", header=T, sep=",", row.names = NULL, stringsAsFactors = F)

# select columns
table_s <- table[, c(1,2,5)]
colnames(table_s)[3] <- "Survival Rate"

# create No for name
table_s$No <- ""

j = 1
for (i in seq(1, 133, by = 6)){
  table_s[c(i:(i+5)),4] <- j
  j = j + 1
}

firmlist <- unique(table_s$name)

# make plot for each firm
for (i in 1:23){
  png(file = paste0("output/figure/", firmlist[i], "(日波動率).png"), width = 600, height = 400)
  lineplot <- ggplot(table_s[which(table_s$No == i),], aes(x = Year, y = `Survival Rate`, group = 1)) + 
    geom_line() + 
    geom_point()+
    ggtitle(paste(firmlist[i], "在碳費2000元的存活機率"))+
    ylim(0, 1.00)+
    theme(legend.position="top")+
    theme(text=element_text(family="Taipei Sans TC Beta", size=10))
  print(lineplot)
  dev.off()
}



Table1[!duplicated(Table1$Var2, fromLast = T), ]# 月年化
