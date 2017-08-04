import pandas as pd
import os
import tkinter
import matplotlib.pyplot as plt
import seaborn as sns

def Ticker_Weights_Sum(index):
    Index_Weighter = index['Weight'].groupby(index['TICKER'])
    Index_Weights = pd.DataFrame(Index_Weighter.sum())
    Index_Weights = Index_Weights.rename(columns={'Weight':'Index_Weight'})

    Ticker_Weighter = index['PSA_Weights'].groupby(index['TICKER'])
    Ticker_Weights = pd.DataFrame(Ticker_Weighter.sum())
    Ticker_Weights = Ticker_Weights.rename(columns={'PSA_Weights':'Fund_Weight'})

    Ticker_Matrix = Index_Weights.join(Ticker_Weights)
    Ticker_Matrix = Ticker_Matrix.reset_index()
    Ticker_Matrix['Weight_Difference'] = (Ticker_Matrix['Fund_Weight'] - Ticker_Matrix['Index_Weight']) * 100

    return Ticker_Matrix

def Include_Rebal(index, dater):
    os.chdir('R:/Fixed Income/IVY/Index Holdings')
    path = str('R:/Fixed Income/IVY/Index Holdings/' + dater + '_SP5MAIG_PRO.SPFIC')
    next_index = pd.read_csv(path, sep='\t')
    index_link = next_index.merge(index, on='CUSIP', how='outer', indicator=True)
    index_link.to_csv('index_comparison.csv')
    index_link = index_link[index_link._merge == 'right_only']
    bad_cusips = index_link['CUSIP'].tolist()

    return bad_cusips

#def inputwidgets(question, response):

def IVY_Plots(master_set):
    indx_ratings = master_set['Index_Weight'].groupby(master_set['SP RATING'])
    indx_ratings_gp = pd.DataFrame(indx_ratings.sum())
    print(indx_ratings_gp)
    fund_ratings = master_set['Fund_Weight'].groupby(master_set['SP RATING'])
    fund_ratings_gp = pd.DataFrame(fund_ratings.sum())
    print(fund_ratings_gp)
