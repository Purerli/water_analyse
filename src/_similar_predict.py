#coding:utf-8
def similar_ayalyse(YearMonth='202201',Infilename=None,Outdirname=None):
    import pandas as pd
    nums=[]
    word_nums=[]
    def output_word_nums(demand_type):
        Year=int(YearMonth[:4])
        Month=int(YearMonth[4:])
        word_nums.append(pd.read_excel(Infilename).shape[0])
        word_nums.append(len(nums))
        word_nums.extend(demand_type)
    def string_similar(s0,s1,s2,s3):
        import difflib

        d0=difflib.SequenceMatcher(None, s0, s1).quick_ratio()
        d1=difflib.SequenceMatcher(None, s2, s3).quick_ratio()
        return (d0+d1)/2.0
    def similar_excel():
        orgin_df=pd.read_excel(Infilename,header=None)
        data_df=pd.read_excel(Infilename)
        data0=data_df['诉求内容'].tolist()
        data1=data_df['区级'].tolist()
        excel_shape=data_df.shape[0]

        for i in range(excel_shape):
            
            j=i+1
            while(j<excel_shape):
                
                if data0[i]=='0' or data0[j]=='0' or data1[i]=='0' or data1[j]=='0':
                    j+=1
                    continue
                else:
                    simi_rate=string_similar(data0[i],data0[j],str(data1[i]),str(data1[j]))
                    if(simi_rate>=0.8):
                        if i not in nums:
                            nums.extend([0,i])
                        nums.append(j)
                        data0[j]='0'
                        data1[j]='0'
                    j+=1
            
                # end_nums.extend(nums)
            # [end_nums.append(i) for i in nums if not i in end_nums]
        scrdf = pd.DataFrame(orgin_df.iloc[nums].values)
        savepath = Outdirname+'\\similar.xlsx'
        writer = pd.ExcelWriter(savepath)
        scrdf.to_excel(writer, sheet_name="Sheet1", index=False)
        writer.save()
        # return nums
    def count_similar():
        ans=0
        id1 = [i for i,x in enumerate(nums) if x==0]
        print(id1)


    similar_excel()
    count_similar()

infile=r'D:\Users\MR\Desktop\北京市水务局涉水数据分析软件v1.1\直派水务局\2022.4.xlsx'
outfile=r'D:\Users\MR\Desktop\北京市水务局涉水数据分析软件v1.1'
similar_ayalyse(Infilename=infile,Outdirname=outfile)

def predict_plitti(demand='趋势预测',YearMonth='202201',Infilename=None,Outdirname=None):
    import datetime
    import itertools
    import os
    # import matplotlib.pyplot as plt
    import numpy as np
    import pandas as pd
    import scipy as sp
    import seaborn as sns  # 热力图
    import statsmodels.api as sm
    import statsmodels.tsa.stattools as ts
    from statsmodels.tsa.arima.model import ARIMA
    def predict():
        # print(Get_data(df,21,1))
        def Get_data(df,year,month,day):
            num=[]
            for i in range(len(df['年'])):
                # print(str(df['年'].iloc[i]),str(df['月'].iloc[i]))
                if df['年'].iloc[i]==year and df['月'].iloc[i]==month:
                    # if df['二级问题'].iloc[i]=='供水' or df['三级问题'].iloc[i]=='供水':
                    if df['日'].iloc[i]==day :
                        num.append(i)
            if num:
                scrdf = pd.DataFrame(df.iloc[num].values)
                scrdf.columns=df.columns
                # print(scrdf)
                return scrdf
            else:
                return np.array([])
        def YearandMonth():
            data20=[]
            data21=[]
            data22=[]
            bigmonth=[1,3,5,7,8,10,12]
            smallmonth=[4,6,9,11]
            specialmonth=[2]
            Year=int(YearMonth[:-2])
            Month=int(YearMonth[-2:])
            data=pd.DataFrame()
            if Year==2020:
                for i in range(Month):
                    if (i+1)!=Month:
                        if (i+1) in bigmonth:
                            for j in range(31):
                                data20.append(Get_data(df,20,i+1,j+1).shape[0])
                        elif (i+1) in smallmonth:
                            for j in range(30):
                                data20.append(Get_data(df,20,i+1,j+1).shape[0])
                        elif (i+1) in specialmonth:
                            for j in range(29):
                                data20.append(Get_data(df,20,i+1,j+1).shape[0])
                    else:
                        for j in range(18):
                            data20.append(Get_data(df,20,i+1,j+1).shape[0])
                data['数值']=data20
            elif Year==2021:
                for i in range(12):
                    if (i+1) in bigmonth:
                        for j in range(31):
                            data20.append(Get_data(df,20,i+1,j+1).shape[0])
                    elif (i+1) in smallmonth:
                        for j in range(30):
                            data20.append(Get_data(df,20,i+1,j+1).shape[0])
                    elif (i+1) in specialmonth:
                        for j in range(29):
                            data20.append(Get_data(df,20,i+1,j+1).shape[0])
                for i in range(Month):
                    if (i+1)!=Month:
                        if (i+1) in bigmonth:
                            for j in range(31):
                                data21.append(Get_data(df,21,i+1,j+1).shape[0])
                        elif (i+1) in smallmonth:
                            for j in range(30):
                                data21.append(Get_data(df,21,i+1,j+1).shape[0])
                        elif (i+1) in specialmonth:
                            for j in range(28):
                                data21.append(Get_data(df,21,i+1,j+1).shape[0])
                    else:
                        for j in range(18):
                            data21.append(Get_data(df,21,i+1,j+1).shape[0])
                data['数值']=data20+data21
            elif Year==2022:
                for i in range(12):
                    if (i+1) in bigmonth:
                        for j in range(31):
                            data20.append(Get_data(df,20,i+1,j+1).shape[0])
                            data21.append(Get_data(df,21,i+1,j+1).shape[0])
                    elif (i+1) in smallmonth:
                        for j in range(30):
                            data20.append(Get_data(df,20,i+1,j+1).shape[0])
                            data21.append(Get_data(df,21,i+1,j+1).shape[0])
                    elif (i+1) in specialmonth:
                        for j in range(29):
                            data20.append(Get_data(df,20,i+1,j+1).shape[0])
                            data21.append(Get_data(df,21,i+1,j).shape[0])
                for i in range(Month):
                    if (i+1)!=Month:
                        if (i+1) in bigmonth:
                            for j in range(31):
                                data22.append(Get_data(df,22,i+1,j+1).shape[0])
                        elif (i+1) in smallmonth:
                            for j in range(30):
                                data22.append(Get_data(df,22,i+1,j+1).shape[0])
                        elif (i+1) in specialmonth:
                            for j in range(28):
                                data22.append(Get_data(df,22,i+1,j+1).shape[0])
                    else:
                        for j in range(18):
                            data22.append(Get_data(df,22,i+1,j+1).shape[0])
                data['数值']=data20+data21+data22

            data.index=pd.Index(pd.date_range(start='20200101',end=YearMonth+'18'))
            
            plt.plot(data.index,data['数值'],'-')
            plt.savefig(Outdirname+'\\pic\\原始数据.jpg')
            # data['数值'].diff(1).plot()
            # plt.savefig(Outdirname+'\\pic\\一阶差分后数据.jpg')
            return data
        
        def judge_stationarity(data):
            dftest = ts.adfuller(data)
            # print(dftest)
            dfoutput = pd.Series(dftest[0:4], index=['Test Statistic','p-value','#Lags Used','Number of Observations Used'])
            stationarity = 1
            for key, value in dftest[4].items():
                dfoutput['Critical Value (%s)'%key] = value 
                if dftest[0] > value:
                        stationarity = 0
            # print(dfoutput)
            print("是否平稳(1/0): %d" %(stationarity))
            return stationarity
        def diff(timeseries,num):
            data_diff = timeseries.diff(num)
            data_diff = data_diff.dropna()
            plt.figure()
            plt.plot(data_diff)
            plt.savefig(Outdirname+'\\pic\\差分后数据.jpg')
            if num==1:
                plt.title('一阶差分')
            elif num==2:
                plt.title('二阶差分')
            # plt.show()
            return data_diff
        def get_P_Q(data):    #获得arim的p值和q值
            pmax = int(5)    #一般阶数不超过 length /10
            qmax = int(5)
            bic_matrix = []
            for p in range(pmax +1):
                temp= []
                for q in range(qmax+1):
                    try:
                        temp.append(ARIMA(data['数值'].diff(1), order=(p, 1, q)).fit().bic)
                    except:
                        temp.append(None)
                    bic_matrix.append(temp)
            # print(bic_matrix)
            bic_matrix = pd.DataFrame(bic_matrix)   #将其转换成Dataframe 数据结构
            p,q = bic_matrix.stack().idxmin()   #先使用stack 展平， 然后使用 idxmin 找出最小值的位置
            # print(u'BIC 最小的p值 和 q 值：%s,%s' %(p,q))  #  BIC 最小的p值 和 q 值：0,1
            return p,q
        df = pd.read_excel(Infilename)
        plt.rcParams['font.sans-serif'] = ['SimHei']
        if not os.path.exists(Outdirname+'\\pic\\'):
            os.makedirs(Outdirname+'\\pic\\')
        data=YearandMonth()
        if judge_stationarity(data)==0:
            myts_diff = diff(data,1)
        if judge_stationarity(myts_diff)==0:
            myts_diff = diff(data,2)
        elif judge_stationarity(myts_diff)==1:
            p,q=get_P_Q(data)
            result_arima = ARIMA(data, order=(p,1,q)).fit()
        # result_arima_fit=model.fit()
            predict_ts = result_arima.predict(start=str(YearMonth)+'19',end=str(int(YearMonth)+1)+'18',typ='levels')  #若不设置typ，预测的值为差分值，需要自己还原
            print(predict_ts)
            sum_predict=np.sum(list(map(int,predict_ts.values.tolist())))
            print(sum_predict)
        # data = data[predict_ts.index]  # 过滤没有预测的记录

        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.figure()
        plt.plot(data.index.tolist(),data['数值'].tolist(),color='skyblue', label='Predict')
        plt.plot(predict_ts.index.tolist(),predict_ts.values.tolist(),color='tomato', label='Original')
        plt.savefig(Outdirname+'\\pic\\预测数据.jpg')

    def plitti():
        df=pd.read_excel(Infilename)
    # if __name__=='__main__':
    if demand=='趋势预测':
        predict()
    elif demand=='突变点标识':
        plitti()
