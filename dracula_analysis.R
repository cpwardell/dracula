#### Load packages ####
library(openxlsx) # for reading Excel files
library(RColorBrewer) # nice color palettes
library(ggplot2) # fancy-looking plots
library(corrplot) # correlation plot
library(network) # for network plot
library(sna) # for network plot
library(tm) # Text mining package
library(wordcloud) # word cloud plotting

#### Read in data and setup ####

## Read from Excel.  Note that we have word counts for chapters but never bother using them
chron = read.xlsx("dracula.xlsx",sheet="chronology", startRow=1, colNames=TRUE, skipEmptyRows=TRUE, rowNames=FALSE)
wc = read.xlsx("dracula.xlsx",sheet="wordcounts", startRow=1, colNames=TRUE, skipEmptyRows=TRUE, rowNames=FALSE)

## Reorder type column for aesthetic reasons and cast chapters as a factor
chron$Type=factor(chron$Type,levels=c("Diary","Letter","Telegram","Memo","Newspaper","Ship's log"))
chron$Chapter=factor(chron$Chapter)

## To reorder the characters, we manually define the order to ensure proper grouping
## A frequency-based approach looks more confused
primarycharacters=c("John Seward",
                    "Jonathan Harker",
                    "Wilhelmina Murray-Harker",
                    "Lucy Westenra",
                    "Abraham Van Helsing", 
                    "Arthur Holmwood",     
                    "Quincey P Morris")     

secondarycharacters=c("Captain of the Demeter",
                      "The Dailygraph",
                      "The Westminster Gazette",   
                      "The Pall Mall Gazette",
                      "Carter, Paterson & Co",
                      "Samuel F Billington & Son",
                      "Mitchell, Sons & Candy",
                      "Rufus Smith",
                      "Patrick Hennessey",    
                      "Sister Agatha")        
chron$From=factor(chron$From,levels=c(primarycharacters,secondarycharacters))
chron$To=factor(chron$To,levels=c(primarycharacters,secondarycharacters))

## Store the initials of each character for labelling plots
initials=c("JS","JH","WMH","LW","AVH","AH","QPM","CD","DG","WG","PMG","CPC","SBS","MSC","RS","PH","SA")

## Calculate days of the week - this cannot be done in Excel before the year 1900
chron$Date=as.Date(gsub(" ","-",chron$Date))
chron$Weekday=factor(weekdays(chron$Date),levels=c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"))

## Define color palettes
pcols=brewer.pal(length(primarycharacters),"Set1")
scols=brewer.pal(length(secondarycharacters),"Set3")
acols=c(pcols,scols)

gothic=c("#6B090F","#360511","#080004","#230024","#3D1D54")

#### End of reading data and setup ###

#### Basic plots ####

## Which days documents were written on
ggplot(chron,aes(Weekday,fill=Type))+geom_bar()+ggtitle("Documents by day of the week")

## Plot document types with who wrote them and who to:
ggplot(chron,aes(Type,fill=From))+geom_bar()+scale_fill_manual(values=acols)+ggtitle("Documents From")
ggplot(chron,aes(Type,fill=To))+geom_bar()+scale_fill_manual(values=acols)+ggtitle("Documents To")

## Who writes what and when, split by chapters
ggplot(chron,aes(Chapter,fill=Type))+geom_bar()+ggtitle("Document type by chapter")
ggplot(chron,aes(Chapter,fill=From))+geom_bar()+scale_fill_manual(values=acols)+ggtitle("Document sender by chapter")
ggplot(chron,aes(Chapter,fill=To))+geom_bar()+scale_fill_manual(values=acols)+ggtitle("Document receiver by chapter")

## A heatmap of who sends what to who...
fromtotable=table(chron$From,chron$To)
fromtoproportions=fromtotable/rowSums(fromtotable)
corrplot(fromtoproportions,method="color",tl.col="black")

## corrplot() messes with the device setup, so reset it
dev.off()

#### End of basic plots ####

#### Timeline ####

## Crop off the last entry, as it's 7 years later
main=chron[1:197,]

## Define heights for points
height=rep(0,nrow(main))
height[main$Type=="Diary"]= 1
height[main$Type=="Ship's log"]=2
height[main$Type=="Memo"]= 3
height[main$Type=="Telegram"]=4
height[main$Type=="Newspaper"]= 5
height[main$Type=="Letter"]= 6 

## Define vectors of colours for the plot
fromcols=as.character(main$From)
tocols=as.character(main$To)
for(i in 1:length(levels(chron$From))){
  fromcols[fromcols==levels(chron$From)[i]]=acols[i]
  tocols[tocols==levels(chron$From)[i]]=acols[i]
}

## Use chapter information to determine whether the point is above or below the line
## Odd numbers above, even below
main$Chapter=as.numeric(as.character(main$Chapter))
height[main$Chapter%%2==0]=height[main$Chapter%%2==0]*-1

## Perform plot
plot(x=main$Date,y=rep(1,nrow(main)),type="n",xaxt = "n",yaxt="n",
     bty = "n",xlab = "Date", ylab=NA,ylim=c(-6,6),main="Dracula Timeline")
abline(h=-6:6,col="lightgrey")
points(main$Date,height,pch=15,col=fromcols)
segments(main$Date,rep(0,nrow(main)),main$Date,height,col=fromcols,lty=1)
#points(main$Date,height,pch=15,col=tocols) #uncomment for "to" colours
#segments(main$Date,rep(0,nrow(main)),main$Date,height,col=tocols) #uncomment for "to" colours
u = par("usr")
arrows(u[1], 0, u[2], 0, xpd=TRUE,lwd=3)
xlabels=paste(c(month.abb[5:11]),"1890")
ylabels=c("Diary","Ship's log","Memo","Telegram","Newspaper","Letter")
axis(1, at=as.Date(c("1890-05-01","1890-06-01","1890-07-01","1890-08-01",
                     "1890-09-01","1890-10-01","1890-11-01")),labels=xlabels) 
axis(2,at=c(-6:6),labels=c(rev(a),"",a),las=2,lty=0,cex.axis=0.7)

#### End of timeline ####


#### Network ####

## Create network of who communicates with who
comtable=table(chron$From,chron$To)
comtablebinary=comtable/comtable # divide by itself so that all values are 1
comnetwork=network(comtablebinary,directed=TRUE,loops=TRUE)

## Number of people who write to each character
to=(colSums(comtablebinary,na.rm=TRUE))
## Number of people who each character writes to
from=(rowSums(comtablebinary,na.rm=TRUE))
sentreceived=cbind(from,to)

## Barplot of results
barplot(t(sentreceived),main="Characters interacted with",
        ylab="Characters",beside=TRUE,names=initials,
        las=3,col=c("#078E53","#111987"))
legend("topright",c("Sent","Received"),pch=rep(15,2),col=c("#078E53","#111987"))

## Generate some scaled edge widths
edgewidths=as.vector(comtable)
edgewidths=edgewidths[edgeswidths!=0]
edgewidths=sapply(log2(edgewidths),max,1)

## Perform network plot
gplot(comnetwork,
      interactive=TRUE, # Uncomment for interactive plotting
      diag=TRUE, # self-loops
      loop.steps=5, # angular self-loops
      usecurve=TRUE, # curvy arrows
      arrowhead.cex=0.5, # smaller arrowheads
      label=initials, # initials not full names
      label.pos=5, # labels inside shapes
      label.cex=ifelse(levels(chron$From)%in%primarycharacters,1.1,0.75),
      label.col=ifelse(levels(chron$From)%in%primarycharacters,"black","white"),
      vertex.cex=(log2(table(c(as.character(chron$From),as.character(chron$To)))[levels(chron$From)])+3)/2, # log2 scale
      vertex.sides=ifelse(levels(chron$From)%in%primarycharacters,6,3),
      vertex.rot=ifelse(levels(chron$From)%in%primarycharacters,0,30),
      vertex.col=ifelse(levels(chron$From)%in%primarycharacters,"#078E53","#111987"),
      edge.lwd=edgewidths
      )

#### End of network ####


#### Arc diagram ####

## IMPORTANT NOTE: loading this package will break the network-plotting
## code, which is why it's all commented out here:
# https://github.com/gastonstat/arcdiagram
# http://gastonsanchez.com/work/starwars/
# library(devtools) # for installing arcdiagram
# install_github('arcdiagram',  username='gastonstat')
# library(arcdiagram) # for plotting arc diagram

## Generate a suitable object, order it for easy access and 
## remove self-communication, as this can't be represented
## on an arc diagram
arcmat=as.matrix(chron[,c("From","To")])
arcmat=arcmat[!(arcmat[,"From"]==arcmat[,"To"]),]
arcmat=arcmat[order((arcmat[,2])),]
arcmat=arcmat[order((arcmat[,1])),]

## Define which arcs we will place above and which below
AVH2JS=c(1:6,14:16)
AVH2JH=c(7)
AVH2WMH=c(8:9,31:32)
WMH2SA=30
WMH2LW=c(21:23,33:36)

## Define a custom order of nodes
ordering=c("Abraham Van Helsing",      
           "John Seward",             
           "Jonathan Harker",
           "Wilhelmina Murray-Harker",
           "Arthur Holmwood",         
           "Quincey P Morris",     
           "Lucy Westenra",
           "Carter, Paterson & Co", 
           "Samuel F Billington & Son",
           "Mitchell, Sons & Candy", 
           "Sister Agatha",
           "Rufus Smith",
           "Patrick Hennessey")

## Perform the arc plot
## Col.nodes has a bug, so colours are defined manually
## Could be improved with weighted line arcs and/or nodes
arcplot(arcmat,above=c(AVH2JS,AVH2JH,AVH2WMH,WMH2SA,WMH2LW),
        ordering=ordering,
        pch.nodes=15,
        col.nodes=c(rep("#078E53",6),rep("#111987",2),"#078E53",rep("#111987",4)),
        cex.nodes=2,
        cex.labels=0.5,
        lwd.arcs=3,
        col.arcs="#A71930")
legend("topleft",c("Primary character","Secondary character"),pch=rep(15,2),col=c("#078E53","#111987"),box.lty=0)

#### End of arc diagram ####


#### Word cloud ####

## Let's plot a word cloud. It's not really quantitative, but it's pretty

## Create a corpus and do some basic cleaning up
book=read.delim("pg345.txt",header=FALSE,stringsAsFactors=FALSE)
corpus = Corpus(DataframeSource(book), readerControl = list(language = "en"))
corpus = tm_map(corpus, content_transformer(removePunctuation))
corpus = tm_map(corpus, content_transformer(tolower))
corpus = tm_map(corpus, function(x) removeWords(x, stopwords("english")))

## Create the necessary objects for plotting
tdm = TermDocumentMatrix(corpus)
tdmatrix = as.matrix(tdm)
tdcounts = sort(rowSums(tdmatrix),decreasing=TRUE)
tddf = data.frame(word = names(tdcounts),freq=tdcounts)

## Perform plot
wordcloud(tddf$word,tddf$freq,min.freq=3,max.words=500, random.order=FALSE, rot.per=0.15, colors=gothic,random.color=TRUE)
wordcloud(tddf$word,tddf$freq,min.freq=3, random.order=FALSE, rot.per=0.15, colors=gothic,random.color=TRUE)

#### End of word cloud ####

