#install.packages("nleqslv")
library(nleqslv)


getwd()
setwd("D:/Drivers/台大風險中心/14.模型與指標/Merton違約機率")
setwd("/Users/Galong/Insync/Drivers/台大風險中心/14.模型與指標/Merton違約機率")

############################################################################## Sigma以股票日報酬率波動率換算成年化波動率
# 26家南科電子業廠商違約機率(以2020/12/31起算每五年的存活機率到2050年)
# L和Vo單位為千元, 資料來源:TEJ IFRS以合併為主財務(累計)_一般行業IV
# Sigma:log(sqrt(股價最高點/股價最低點))，股價資料來源：TEJ股價資料庫-TEJ調整股價(年)-除權息調整
# r為目前政府十年公債利率的資料來源：中華民國統計資訊網https://statdb.dgbas.gov.tw/pxweb/Dialog/varval.asp?ma=FM3007A1A&ti=%A7Q%B2v%B2%CE%ADp-%B8%EA%A5%BB%A5%AB%B3%F5%A7Q%B2v-%A6%7E&path=../PXfile/FinancialStatistics/&lang=9&strList=L
# 碳排量為TEJ CSR資料庫的碳排放量（範疇一加範疇二排放量，單位公噸）

## read files
stock_price <- read_xlsx("rawdata/2020TEJ調整股價(年)-除權息調整(all).xlsx",skip = 2)
IFRS <- read_xlsx("rawdata/2020TEJ_IFRS以合併為主財務(累計)_一般行業IV(all).xlsx",skip = 2)
SP26 <- read_xlsx("rawdata/26家南科製造業.xlsx")
carbon <- read_xlsx("rawdata/2014-2020_all公司碳排放量.xlsx", skip = 2) #TEJ CSR資料庫
rate <- read_xlsx("rawdata/1997-2021利率統計-資本市場利率-年.xlsx") #需要手動修改.csv檔再轉成.xlsx檔


## 政府10年公債利率
rate2020 <- rate[which(rate$年 == 2020), 5] %>% as.numeric() # select rate for 2020
rate2020

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
df$Sigma0 <- df$sd*sqrt(245)

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

##输入时间
T<-1; ##输入流动负债、非流动负债 SD<-1e8;
LD<-50000000;
##计算违约点
DP<-SD+0.5*LD;
##可以根据fair value修改违约点
D<-DP;
##输入股权波动率和股权价值
PriceTheta<-0.2893 
EquityTheta<-PriceTheta*sqrt(12) #以月波动率为例 E=141276427;
##KMV 方程变形及求解 EtoD <- E/D;
x0 <- c(1,1); KMVfun<-function(x) {
  y <- numeric(2);
  d1<-( log(x[1]*EtoD)+(r+0.5*x[2]^2)*T ) / ( x[2]*sqrt(T)); d2<-d1-x[2]*sqrt(T); y[1]<-x[1]*pnorm(d1)-exp(-r*T)*pnorm(d2)/EtoD-1; y[2]<-pnorm(d1)*x[1]*x[2]-EquityTheta;
  y
}
z<-nleqslv(x0, KMVfun, method="Newton")
Va<-z$x[1]*E
AssetTheta<-z$x[2]
##计算违约距离 ##DD=(log(Va/DP)+(r-0.5*AssetTheta^2)*T)/(AssetTheta*T) DD<-(Va-DP)/(Va*AssetTheta)
##计算违约率
EDF<-pnorm(-DD)