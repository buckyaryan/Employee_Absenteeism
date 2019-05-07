rm(list = ls())

library(xlsx)
library(ggplot2)
library(DMwR)
library(gridExtra)
library(MASS)
library(corrgram)
library(rpart)
library(DataCombine)
library(inTrees)
library(caret)
library(gbm)
library(randomForest)
library(mlr)
library(dummies)
library(usdm)
library(rpart.plot)

#libraries to install
#x = c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071", "Information",
#"MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees')

df = read.xlsx("Absenteeism_at_work_Project.xls",sheetIndex = 1)

# Univariate Analysis and Variable Consolidation
str(df)

df$ID = as.factor(as.character(df$ID))

df$Reason.for.absence[df$Reason.for.absence %in% 0] = 20
df$Reason.for.absence = as.factor(as.character(df$Reason.for.absence))
                                  
df$Month.of.absence[df$Month.of.absence %in% 0] = NA
df$Month.of.absence = as.factor(as.character(df$Month.of.absence))

df$Day.of.the.week = as.factor(as.character(df$Day.of.the.week))
df$Seasons = as.factor(as.character(df$Seasons))
df$Disciplinary.failure = as.factor(as.character(df$Disciplinary.failure))
df$Education = as.factor(as.character(df$Education))
df$Social.drinker = as.factor(as.character(df$Social.drinker))
df$Social.smoker = as.factor(as.character(df$Social.smoker))
df$Son = as.factor(as.character(df$Son))
df$Pet = as.factor(as.character(df$Pet))

str(df)

##################################Missing Values Analysis###############################################

missing_val = data.frame(apply(df,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] = "Missing_Percentage" 
missing_val$Missing_Percentage = (missing_val$Missing_Percentage/nrow(df))*100
missing_val = missing_val[order(-missing_val$Missing_Percentage),]
row.names(missing_val) = NULL
missing_val = missing_val[,c(2,1)]
write.csv(missing_val,"Missing_Percentage.csv",row.names = F)

# Missing value percentage representation by bar chart

ggplot(data = missing_val[1:3,], aes(x=reorder(Columns, -Missing_Percentage),y = Missing_Percentage))+
  geom_bar(stat = "identity",fill = "grey")+xlab("Columns")+
  ggtitle("Missing data percentage") + theme_bw()

######## Missing value imputation by KNN Method ##############

#Actual Value = 29
#mean   = 26.68
#median = 25
#knn = 29

#mean method
#df$Body.mass.index[is.na(df$Body.mass.index)] = mean(df$Body.mass.index, na.rm = T)

# median method
#df$Body.mass.index[is.na(df$Body.mass.index)] = median(df$Body.mass.index, na.rm = T)

#KNN
df = knnImputation(df,5)

#Note--- Compare to all the methods, knn method results are more near to the actual values so i decided to go with knn method for further processing--

#Separating continuous and catagorical variables

#for numerical variables
numeric_index = sapply(df, is.numeric)
numeric_data = df[,numeric_index]
continuous_variables = colnames(numeric_data)

#continuous_variables = ["Transportation.expense", "Distance.from.Residence.to.Work", 
                          "Service.time" , "Age" , "Work.load.Average.day." ,
                          "Hit.target", "Weight" , "Height", "Body.mass.index",
                          "Absenteeism.time.in.hours"
                        ]


#for catagorical variables
factor_index = sapply(df, is.factor)
factor_data = df[,factor_index]
catagorical_variables = colnames(factor_data)

#catagorical_variables = [ "ID", "Reason.for.absence", "Month.of.absence", "Day.of.the.week",
                           "Seasons", "Disciplinary.failure", "Education", "Son",                
                           "Social.drinker",  "Social.smoker", "Pet"
                          ]

############ Outlier Analysis ##################

#Outlier analysis for continuous variables.

for (i in 1:length(continuous_variables)){
  assign(paste0("gn",i), ggplot(aes_string(y = (continuous_variables[i]), x = "Absenteeism.time.in.hours"), data = subset(df))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=continuous_variables[i],x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot of Employee Absenteeism for",continuous_variables[i])))
}

#generating plots

gridExtra::grid.arrange(gn1,gn2,ncol=2)
gridExtra::grid.arrange(gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,ncol=2)
gridExtra::grid.arrange(gn7,gn8,ncol=2)
gridExtra::grid.arrange(gn9,gn10,ncol=2)
#################################EDDDDDDDDD########################

#Distribution of factor data using bar plot

bar1 = ggplot(data = df, aes(x = ID)) + geom_bar() +
      ggtitle("Count of ID") + theme_minimal()

bar2 = ggplot(data = df, aes(x = Reason.for.absence)) +
      geom_bar() + ggtitle("Count of Reason for absence") + theme_bw()

bar3 = ggplot(data = df, aes(x = Month.of.absence)) + geom_bar()+
      ggtitle("Count of Month") + theme_bw()

bar4 = ggplot(data = df, aes(x = Disciplinary.failure)) + 
      geom_bar() + ggtitle("Count of Disciplinary failure") + theme_bw()

bar5 = ggplot(data = df, aes(x = Education)) + geom_bar()+
      ggtitle("Count of Education")+ theme_bw()

bar6 = ggplot(data = df, aes(x = Son)) + geom_bar()+
        ggtitle("Count of Son") + theme_bw() 

bar7 = ggplot(data = df, aes(x = Social.smoker)) + geom_bar() + 
        ggtitle("Count of Social smoker") + theme_bw()


#Check the distribution of numerical data using histogram

hist1 = ggplot(data = numeric_data, aes(x =Transportation.expense)) + ggtitle("Transportation.expense") + geom_histogram(bins = 25)

hist2 = ggplot(data = numeric_data, aes(x =Height)) + ggtitle("Distribution of Height") + geom_histogram(bins = 25)

hist3 = ggplot(data = numeric_data, aes(x =Body.mass.index)) + ggtitle("Distribution of Body.mass.index") + geom_histogram(bins = 25)

hist4 = ggplot(data = numeric_data, aes(x =Absenteeism.time.in.hours)) + ggtitle("Distribution of Absenteeism.time.in.hours") + geom_histogram(bins = 25)


#Replace all outliers with NA and impute

for(i in continuous_variables){
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  #print(length(val))
  df[,i][df[,i] %in% val] = NA
}

#Imputing missing values with Knn method
df = knnImputation(df,k=5)

############# Feature selection #######################################
#Checking for multicollinearity using VIF
vifcor(numeric_data)

#checking multicollinearity using correlation graph
corrgram(df[,numeric_index], order= F,
         upper.panel=panel.pie, text.panel = panel.txt, main = "Correlation Plot")


#Dimension Reduction

df_new = subset(df,select = -c(Body.mass.index))
df_new_cleaned = df_new

################ Feature Scaling #######################################

#Normality Checking
qqnorm(df_new$Absenteeism.time.in.hours)
hist(df_new$Absenteeism.time.in.hours)


#Updating the continuous variables for normalization

continuous_variables = c('Transportation.expense', 'Distance.from.Residence.to.Work', 'Service.time',
                         'Age', 'Work.load.Average.day.',
                         'Hit.target', 'Height', 
                         'Weight')

catagorical_variables = c('ID','Reason.for.absence','Disciplinary.failure', 
                          'Social.drinker', 'Son', 'Pet', 'Month.of.absence', 'Day.of.the.week', 'Seasons',
                          'Education', 'Social.smoker')


#Normalization
for(i in continuous_variables)
{
  print(i)
  df_new[,i] = (df_new[,i] - min(df_new[,i]))/(max(df_new[,i])-min(df_new[,i]))
}

#Creating dummy variables for catagorical varibales to maintain the levels.

df_new = dummy.data.frame(df_new, catagorical_variables)


#Cleaning the Environment

rmExcept(keepers = c("df_new","df","df_new_cleaned"))

######################## Model development ###################

#dividing data into test and train using simple random sampling method (because the target variable is continuous variable)

set.seed(1)
train_index = sample(1:nrow(df_new),0.8 * nrow(df_new))
train = df_new[train_index,]
test = df_new[-train_index,]

#----------------Decission tree --------------------#

#rpart for regression on training data

model_DT = rpart(Absenteeism.time.in.hours ~ ., data = train, method="anova")

#Summary of the model
summary(model_DT)

#summary in the form of tree
rpart.plot(model_DT)

#predicting for test data
predictions_DT = predict(model_DT,test[,-115])

#creating separate dataframe for analysing actual and predicted observations.

df_new_dt_pred = data.frame("actual"=test[,115],"predicted" = predictions_DT)

#Error Metrics using PostResample method
#Calculates Performance Across Resamples

print(postResample(pred = predictions_DT, obs = test[,115]))

#Alternate method for error metrics
#regr.eval(test[,115], predictions_DT, stats = c('mae','rmse','mse'))

#Validation

#RMSE     = 2.276450
#Rsquared = 0.440758
#MAE      = 1.694686


#------------ Random Forest --------------------#

#create model using training data
model_RF = randomForest(Absenteeism.time.in.hours ~., data = train, ntree=500)

#predict for test data
predictions_RF = predict(model_RF,test[,-115])

#creating separate dataframe for actual and predicted data
df_new_rf_pred = data.frame("actual" = test[,115], "predicted" = predictions_RF)

#calculate mae, rmse, and rsquared
print(postResample(pred = predictions_RF, obs = test[,115]))

#Validation

#RMSE     = 2.1941162
#Rsquared = 0.4798489
#MAE      = 1.6107736
   

#------------------- Linear Regression ----------------------#

#Train the model with training data
model_LR = lm(Absenteeism.time.in.hours ~ ., data = train)

#summary of the model
summary(model_LR)

#predict for test data
predictions_LR = predict(model_LR,test[,-115])

#creating separate dataframe for actual and predicted data
df_new_lr_pred = data.frame("actual" = test[,115], "predicted" = predictions_LR)

#calculate mae, rmse, and rsquared
print(postResample(pred = predictions_LR, obs = test[,115]))

#Validation
   
#RMSE     = 2.5593810
#Rsquared = 0.3587518
#MAE      = 1.8604850


#Compared to all the models the performance is not optimizing due to dimension of the data,
#this can be achived using the principal component analysis

######## Dimension Reduction by PCA ############################

#principal component analysis
prin_comp = prcomp(train)

#sdev for each of the principal components

pr_stdev = prin_comp$sdev

#calculate variance
pr_var = pr_stdev^2

#Proportion of variance
pro_var = pr_var/sum(pr_var)

#Screen plot
plot(cumsum(pro_var), xlab = "Principal Component",
     ylab = "Cumulative Proportion of Variance Explained",
     type = "l")

#Add a training set with principal of components
train.data = data.frame(Absenteeism.time.in.hours = train$Absenteeism.time.in.hours, prin_comp$x)

#From the above plot selecting 45 components since it explains almost 95+ % data variance
train.data = train.data[,1:45]

#Transform test data into PCA
test.data = predict(prin_comp, newdata = test)
test.data = as.data.frame(test.data)

#select the first 45 components
test.data = test.data[,1:45]

############ Model development after principal component analysis ###############

#----------------Decission tree --------------------#

#rpart for regression on training data

model_DTP = rpart(Absenteeism.time.in.hours ~ ., data = train.data, method="anova")

#Summary of the model
summary(model_DTP)

#summary in the form of tree
rpart.plot(model_DTP)

#predicting for test data
predictions_DTP = predict(model_DTP,test.data)

#creating separate dataframe for analysing actual and predicted observations.

df_new_dtp_pred = data.frame("actual"=test[,115],"predicted" = predictions_DTP)

#Error Metrics using PostResample method
#Calculates Performance Across Resamples

print(postResample(pred = predictions_DTP, obs = test[,115]))

#Alternate method for error metrics
#regr.eval(test[,115], predictions_DTP, stats = c('mae','rmse','mse'))

#Validation

#RMSE     = 0.4427499
#Rsquared = 0.9787753
#MAE      = 0.3012380

   
#------------ Random Forest --------------------#

#create model using training data
model_RFP = randomForest(Absenteeism.time.in.hours ~., data = train.data, ntree=500)

#predict for test data
predictions_RFP = predict(model_RFP,test.data)

#creating separate dataframe for actual and predicted data
df_new_rfp_pred = data.frame("actual" = test[,115], "predicted" = predictions_RFP)

#calculate mae, rmse, and rsquared
print(postResample(pred = predictions_RFP, obs = test[,115]))

#Validation

#RMSE     = 0.4361794
#Rsquared = 0.9828053
#MAE      = 0.2505895
   

#------------------- Linear Regression ----------------------#

#Train the model with training data
model_LRP = lm(Absenteeism.time.in.hours ~ ., data = train.data)

#summary of the model
summary(model_LRP)

#predict for test data
predictions_LRP = predict(model_LRP,test.data)

#creating separate dataframe for actual and predicted data
df_new_lrp_pred = data.frame("actual" = test[,115], "predicted" = predictions_LRP)

#calculate mae, rmse, and rsquared
print(postResample(pred = predictions_LRP, obs = test[,115]))

#Validation

#RMSE     = 0.003016381
#Rsquared = 0.999999023
#MAE      = 0.002247516
   