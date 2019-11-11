import pandas as pd
import numpy as np
import ezodf
import os
from IPython.display import display, HTML
from pprint import pprint

def read_ods(filename, sheet_no=0, header=0):
    tab = ezodf.opendoc(filename=filename).sheets[sheet_no]
    return pd.DataFrame({col[header].value:[x.value for x in col[header+1:]]
                         for col in tab.columns()})

def readQuiz(filename):
    df=read_ods(filename,sheet_no=1)
    # team stats
    dfteam=df[:3].loc[:,:'Errors']
    # quizzer stats
    quizzerHdr=df[4:5].iloc[0].values.tolist()[:6]
    dfquizzer=df[5:20]
    dfquizzer=dfquizzer.loc[:,:'Errors']
    dfquizzer.columns=quizzerHdr
    for k in ['#N/A',None]:
        dfquizzer.drop(dfquizzer[dfquizzer.Quizzer == k].index,inplace=True)
    
    return dfteam,dfquizzer

def readMeet(dname,room):
    dfteam=[]
    dfquizzer=[]
    for qi in range(4):
        fnq=os.path.join(dname,'R%dQ%d.ods'%(room,qi+1))
        print('reading %s'%fnq)
        dfteam0,dfquizzer0=readQuiz(fnq)
        dfteam.append(dfteam0)
        dfquizzer.append(dfquizzer0)

    dfteam=pd.concat(dfteam)
    dfquizzer=pd.concat(dfquizzer)
    return dfteam,dfquizzer
    
def meetStats(df,termList=None):
    nquiz=3
    
    # getterm
    if('Quizzer' in df.columns):
        term='Quizzer'
    elif('Team' in df.columns):
        term='Team'
    else:
        raise Exception('%s not supported'%type)
    if(termList==None):
        # unique term
        termList=list(df[term].unique())
    
    # unique quizzes
    uq=list(np.sort(df['Quiz'].unique()))
    
    # drop NONE and #N/A
    for k in [None,'#N/A']:
        if(k in termList):
            termList.pop(termList.index(k))
        
    # get scores for quizzers
    scores=[]
    for ii,t in enumerate(termList):
        # init dict for this term
        d0={'points':[]}
        if(term=='Quizzer'): d0['quizzer']=t

        # loop through all quizzes
        for q in uq: 
            # get a data frame for this quiz and term (quizzer or team)
            df0=df[(df['Quiz']==q) & (df[term]==t)]
            # if there are rows, then get the score and team
            if(df0.shape[0]):
                d0['points'].append(int(df0.iloc[0].Points))
                d0['team']=df0.iloc[0].Team
        scores.append(d0)

    # form into dataframe
    L=[]
    for ii,v in enumerate(scores):
        pts=v['points']
        d0={term:termList[ii],'Total':sum(pts)}
        # loop through up to the nquiz quizzes this quizzer/team was involved in 
        for jj in range(nquiz):
            if(len(pts)>jj): val=pts[jj]
            else:  val=''
            d0['Q%d'%(jj+1)]=val

        if(term=='Quizzer'):
            if('team' in v):
                d0['team']=v['team']
            else: d0['team']='?'
        L.append(d0)
    dfscores=pd.DataFrame(L)
    
    if(term=='Quizzer'):
        colout=['team',term,"Q1","Q2","Q3","Total"]
    else:
        colout=[term,"Q1","Q2","Q3","Total"]
    dfscores=dfscores[colout]
    #dfscores=dfscores.sort_values(by="Total",ascending=False)
    return dfscores
