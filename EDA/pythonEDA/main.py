import pandas as pd
from pandas_profiling import ProfileReport

df_StudentDetails=pd.read_csv('C:\\Users\\TITUS\\PycharmProjects\\pythonpractice\\StudentDetails1.csv')
profile=ProfileReport(df_StudentDetails, title='EDA STUDENTSRESULTS')
profile.to_file('EDA_STUDENTRESULTS_Analysis.html')
