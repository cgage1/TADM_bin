
# Will also need to install c++ -> https://visualstudio.microsoft.com/visual-cpp-build-tools/
import pandas as pd
import numpy as np
from sklearn.pipeline import Pipeline
import pickle
import sklearn.metrics as skm
from sklearn.metrics import accuracy_score
import os 
import tsfresh 

# Change this to the working directory of where these scripts will be stored. 
os.chdir('C:/Users/cgagel/Box/Hut 8/TADMCurves/Python_dev/MLModelApplications')


with open("model_pipeline_classify.pkl", "rb") as f:
    ppl_1 = pickle.load(f) 


test = pd.read_csv("test_data.csv") #Input new incoming data here. The new data should be in the format as shown in this csv. To get this csv, look at train.py
df_test = test.drop(columns=['classify','measuredvol'],axis=1)

y_test = test[['timekey_dot_studynum','classify']].drop_duplicates().set_index('timekey_dot_studynum').squeeze(axis=1).rename_axis(None, axis=0)

ppl_1.set_params(augmenter__timeseries_container=df_test)
X_test = pd.DataFrame(index=y_test.index)
pred_classify = ppl_1.predict(X_test)
print("Prediction accuracy:{:.2f}%".format(accuracy_score(np.array(y_test), pred_classify)*100))

with open("model_pipeline_regression.pkl", "rb") as f:
    ppl_2 = pickle.load(f) 
y_test = test[['timekey_dot_studynum','measuredvol']].drop_duplicates().set_index('timekey_dot_studynum').squeeze(axis=1).rename_axis(None, axis=0)
ppl_2.set_params(augmenter__timeseries_container=df_test)
X_test = X_test[pred_classify != 0] # Filter rows where tube is not a forced faliure like signature and forward that as input below to the prediction
pred = ppl_2.predict(X_test)
mse = skm.mean_squared_error(y_test, pred,squared=False)
r2 = skm.r2_score(y_test, pred)
maxe = skm.max_error(y_test, pred)
print("Results for the validation set")
print("The mean squared error (MSE) on validation set: {:.4f}".format(mse))
print("The coefficient of determination(R²) on validation set: {:.4f}".format(r2))
print("The max error on validation set: {:.4f}".format(maxe))
