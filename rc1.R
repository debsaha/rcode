setwd("C:/Users/Admin/Desktop")
library(dplyr)
library(openxlsx)
library(ReporteRs)
library(descr)
CHH <- read.xlsx(file.choose(), sheet = 1, colNames = TRUE, rowNames = TRUE)
#tab <- table(chh1$Vidhansabha, chh1$`34..Whom.will.you.cast.your.vote.in.upcoming.Vidhan.Sabha.Elections`)
#tab <- as.data.frame.matrix(tab)
doc <- docx()
highlight <- textProperties(color = '#7030a0', font.size = 20, font.weight = 'bold', font.family = 'calibri')

#for (i in c(14,15,17,19,20,21)) {


tab <- crosstab(CHH$`14.Occupation`,CHH$`19.Which.party.you.will.support/vote.in.corporation.election?`, prop.r = TRUE, percent = TRUE)
tab1 <- (tab$prop.row)*100
tab1 <- round(tab1)
tab1<-t(tab1)
tab<-t(tab)
#tab
text2<-"Vote share by"
text1<-names(CHH[17])
text1<-paste(text2,text1)
text1<-pot(text1,highlight)

doc <- addParagraph(doc, value = text1)

#percent <- prop.table(tab)*100
#percent <- round(percent)
#tab <- desc(tab)
MyFTable <- FlexTable(tab, add.rownames = TRUE, header.cell.props = cellProperties( background.color = "#1E90FF"),
                      header.columns = TRUE, header.text.props = textProperties(color = "black",font.size = 20, font.weight = "bold", font.family = "calibri"), body.text.props = textProperties(font.size = 18, font.family  = "calibri"))
MyFTable1 <- FlexTable(tab1, add.rownames = TRUE, header.cell.props = cellProperties( background.color = "#1E90FF"),
                       header.columns = TRUE, header.text.props = textProperties(color = "black",font.size = 20, font.weight = "bold", font.family = "calibri"), body.text.props = textProperties(font.size = 18, font.family  = "calibri"))

#For Table Border
MyFTable <- setFlexTableBorders(MyFTable, inner.vertical = borderProperties(color = "#A6A6A6",width = 1, style = "solid"), inner.horizontal =  borderProperties(color = "#A6A6A6",width = 1, style = "solid"), outer.vertical =  borderProperties(color = "#A6A6A6",width = 1, style = "solid"), outer.horizontal =  borderProperties(color = "#A6A6A6",width = 1, style = "solid"))
MyFTable1 <- setFlexTableBorders(MyFTable1, inner.vertical = borderProperties(color = "#A6A6A6",width = 1, style = "solid"), inner.horizontal =  borderProperties(color = "#A6A6A6",width = 1, style = "solid"), outer.vertical =  borderProperties(color = "#A6A6A6",width = 1, style = "solid"), outer.horizontal =  borderProperties(color = "#A6A6A6",width = 1, style = "solid"))

#doc <- addParagraph(doc, value = text1)
doc<-addTitle(doc,"   ")
doc <- addFlexTable(doc, MyFTable)
doc <- addFlexTable(doc, MyFTable1)

#}
writeDoc(doc, file = "Output_cr_test1.docx")

names(CHH)
nrow(CHH)
