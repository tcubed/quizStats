import pandas as pd
import numpy as np
import ezodf
import os
from IPython.display import display, HTML
import openpyxl
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

def readMeet(dname,room,verbose=True):
    dfteam=[]
    dfquizzer=[]
    for qi in range(4):
        fnq=os.path.join(dname,'R%dQ%d.ods'%(room,qi+1))
        if(verbose):
            print('reading %s'%fnq)
        dfteam0,dfquizzer0=readQuiz(fnq)
        dfteam.append(dfteam0)
        dfquizzer.append(dfquizzer0)

    dfteam=pd.concat(dfteam)
    dfquizzer=pd.concat(dfquizzer)
    return dfteam,dfquizzer
    
def meetStats(df,termList=None,nquiz=3,sort=False):
    # input
    #   df -- dataframe of meet (all quizzes)
    #   termList -- unique list of quizzers or teams to aggregate stats for
    #               if not provided, it will determine this from dataframe
    #   nquiz -- the number of quizzes a quizzer/team is involved in during a meet
    # returns: df
    #   df -- a dataframe of compiled scores
    #
    # TODO: if nquiz>3, then have to change the number of quizzes reported out
    
    #
    # generate term list for compiling stats
    #
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
    uq=sorted(list(np.sort(df['Quiz'].unique())))
    #print(uq)
    
    # drop NONE and #N/A
    for k in [None,'#N/A']:
        if(k in termList):
            termList.pop(termList.index(k))
    
    #
    # whether teams or quizzers, get scores for quizzers
    #
    # the scores list is of the following dicts:
    # -- quizzers: {'quizzer':string,'team':string,'points':[]}
    # -- teams: {'team':string,'points':[]}
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
            # if the term is Quizzer, then assign the name
            if('team' in v):
                d0['Team']=v['team']
            else:
                d0['Team']='?'
        L.append(d0)
    dfscores=pd.DataFrame(L)
    
    if(term=='Quizzer'):
        colout=['Team',term,"Q1","Q2","Q3","Total"]
    else:
        colout=[term,"Q1","Q2","Q3","Total"]
    dfscores=dfscores[colout]
    #dfscores=dfscores.sort_values(by="Total",ascending=False)
    if(sort==True):
        dfscores=dfscores.sort_values(by='Total',ascending=False)
    return dfscores

def readDivision(meetPaths,divisions=None,verbose=False):
    # create division data structure
    # div{'A':[{'path':<path>,'dfq':<meetQuizzerDataFrame>,'dft':<meetTeamDataFrame>}]}
    if(divisions==None): divisions=['A','B']
    D={}
    for ri,div in enumerate(divisions):
        D[div]=[]
        for m in meetPaths:
            room=ri+1
            dfteam,dfquizzer=readMeet(m,room=room,verbose=verbose)
            d0={'path':m,'dfq':dfquizzer,'dft':dfteam}
            D[div].append(d0)
    return D

def uniqueTeams(meetList):
    # return array of unique team names
    # input
    #    meetList -- list of meet dicts {'path','dfq','dft'}  (as from Division data structure)
    # returns
    #    np array of unique team names
    F=[]
    for m in meetList:
        F.append(m['dft'])
    df=pd.concat(F)
    u=df['Team'].unique()
    return u

def uniqueQuizzers(meetList):
    # return array of unique quizzer names
    # input
    #    meetList -- list of meet dicts {'path','dfq','dft'}  (as from Division data structure)
    # returns
    #    np array of unique quizzer names
    F=[]
    for m in meetList:
        F.append(m['dfq'])
    df=pd.concat(F)
    u=df['Quizzer'].unique()
    return u

def MeetQuizzerCumulativeScores(df,sort=False):
    # return dataFrame of total quizzer scores from a meet [Quizzer, Team, Points, Errors, Jumps]
    qteams=dict(zip(df['Quizzer'],df['Team']))
    qx=df.groupby(['Quizzer'],as_index=False).sum()
    qx=qx.drop(columns=['Quiz'])
    qx['Team']=[qteams[x] for x in qx['Quizzer']]
    if(sort==True):
        qx=qx.sort_values(by='Points',ascending=False)
    return qx

def MeetTeamCumulativeScores(df,sort=False):
    # return dataFrame of total team scores from a meet
    tx=df.groupby(['Team'],as_index=False).sum()
    tx=tx.drop(columns=['Quiz','Place'])
    if(sort==True):
        tx=tx.sort_values(by='Points',ascending=False)
    return tx

def writeStats(fn,meetPaths):
    writer = pd.ExcelWriter(fn) 

    qteams={}
    stats={}
    # meet data, by quizzer and team
    MQ={}
    MT={}
    for room in [1,2]:
        if(room==1): sheet='A-Division Stats'
        else: sheet='B-Division Stats'
        qteams[room]={}
        stats[room]={}
        
        # monthly team and quizzer
        MT[room]=[]
        MQ[room]=[]
        for m in meetPaths:
            dfteam,dfquizzer=readMeet(m,room=room)
            MQ[room].append(dfquizzer)
            MT[room].append(dfteam)
        
        # concat all quizzes/months for this room
        TT=pd.concat(MT[room])
        QQ=pd.concat(MQ[room])
        
        #
        # unique quizzers
        #
        uqz=QQ['Quizzer'].unique()
        #dict(zip(df['Quizzer'],df['team']))
        # unique teams
        utm=TT['Team'].unique()

        #
        # write each month
        #
        # -- for quizzers
        for ii,q in enumerate(MQ[room]):
            # calc meet stats by quizzer
            df=QS.meetStats(q,termList=list(uqz))
            print('Quizzer stats for %s'%meetPaths[ii])
            display(df.head())
            qteams[room].update(dict(zip(df['Quizzer'],df['Team'])))
            
            if(ii==0):
                # for the first meet, keep all columns
                df0=df
                startcol=0
            else:
                # for every other month, only write out the quiz results
                df0=df.loc[:,'Q1':]
                startcol=2+5*ii
            
            # write quizzer monthly stats on division sheet
            df0.to_excel(writer,sheet_name=sheet,startrow=1, startcol=startcol,index=False) 
        
        # -- for teams
        for ii,t in enumerate(MT[room]):
            # calc meet stats by team
            df=meetStats(t,termList=list(utm))

            if(ii==0):
                # for the first meet, keep all columns
                df0=df
                startcol=1
            else:
                # for every other month, only write quiz results
                df0=df.loc[:,'Q1':]
                startcol=2+5*ii
            
            # write team monthly stats on division sheet
            df0.to_excel(writer,sheet_name=sheet,startrow=20, startcol=startcol,index=False) 
        
        # 
        # YTD - all quizzers
        #
        #Q2=QQ.groupby(['Quizzer']).sum()

        Q2=MeetQuizzerTotalScores(QQ)
        display(Q2)
        #Q2=Q2.drop(columns=['Team','Quiz'])
        Q2.to_excel(writer,sheet_name=sheet,startrow=1, startcol=25,index=False) 
        
        #
        # YTD - all team
        #
        #T2=TT.groupby(['Team']).sum()
        T2=MeetTeamTotalScores(TT)
        display(T2)
        #T2=T2.drop(columns=['Team','Quiz'])
        T2.to_excel(writer,sheet_name=sheet,startrow=20, startcol=25,index=False) 

        d={'qstats':Q2,'tstats':T2}
        #cumStats.append(d)

        
    #
    # Monthly Results Tab
    #
    sheet='Monthly Results'
    Q2.to_excel(writer,sheet_name=sheet,startrow=1, startcol=25,index=False) 

    #
    # Year Eng Results Tab
    #

    writer.save()
