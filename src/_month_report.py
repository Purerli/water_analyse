def month_report(Yearmonth=None,Infilename1=None,Infilename2=None,Infilename3=None,Infilename4=None,Outdirname=None):
    Count('直派水务局',Yearmonth,'诉求变化趋势',Infilename2,'./data')
    import codecs
    import os
    import re
    import jieba
    jieba.set_dictionary("./data/dict.txt")
    jieba.initialize()
    # import matplotlib as plot
    import numpy as np
    import pandas as pd
    from django.shortcuts import render
    from docx.shared import Inches, Mm, Pt
    from docxtpl import DocxTemplate, InlineImage
    from numpy.core.numeric import outer
    from pyecharts import options as opts
    from pyecharts.charts import Geo, Map
    from pyecharts.render import make_snapshot
    from snapshot_selenium import snapshot as driver
    Year=int(Yearmonth[:-2])
    Inmonth=int(Yearmonth[-2:])
    class Search_river():
        def __init__(self):
            self.rDict_filename='./data/river.txt' # 河流名称字典
            self.stopkey_filename= './data/stopWord_river.txt'# 停用词字典
            self.tDict_filename='./data/total_dict.txt' # 自建水务词库
        def txt_read(self,files):
            txt_dict = {}
            fopen = open(files,encoding='utf-8')
            for line in fopen.readlines():
                line = str(line).replace("\n","")
                txt_dict[line.split(' ',1)[0]] = line.split(' ',1)[1]
                #split（）函数用法，逗号前面是以什么来分割，后面是分割成n+1个部分，且以数组形式从0开始
            fopen.close()
            return txt_dict


        def get_content(self,df):
            '''
            读取excel中指定两列的内容
            '''
            content1=[]
            content=[]
            for i in range(len(df)):
                content1.append(str(df['诉求内容'].iloc[i]))
                content1.append(str(df['主办单位'].iloc[i]))  
            for j in range(len(content1)):
                if j%2==0:
                    content.append(''.join(content1[j:j+2]))
                else:
                    continue
            return content
        
        def data_clean(self,df):
            '''
            去停用词
            '''
            jieba.load_userdict(self.tDict_filename) # 添加自建词库
            content=self.get_content(df)
            stopkey=[w.strip() for w in codecs.open(self.stopkey_filename, 'r',encoding="utf-8").readlines()]
            m={}
            # 按行边历，先分词，再去除无用词
            for i in range(0,len(content)):
                l=[]
                seg = jieba.lcut(content[i])  
                for j in seg:
                    if j not in stopkey:  
                        l.append(j)
                        m[i]=l
            return m

        def search_river(self,df):
            '''
            在字典中查找河流名称
            '''
            l = []
            suggestion=self.data_clean(df)
            print(suggestion)
            riverdict=self.txt_read(self.rDict_filename)
            keys=list(riverdict.keys())
            # 按行遍历，对于每行中的词，与河流词库匹配。
            for j in range(len(suggestion)):
                    for i in suggestion[j]:
                        count=0
                        for k in riverdict.keys():
                        #  if i in riverdict[k]:
                            c=re.search(i,riverdict[k])
                            a=list(riverdict[k])
                            if c==None or len(i)<2:
                                continue
                            elif c.re.pattern==k :
                                l.append(k)
                                count+=1
                                break
                            elif  a[min(c.span())-1]!=" "  :
                                continue
                            elif a[max(c.span())]==' ':
                                l.append(k)
                                count+=1
                                break  
                        if(count==1):
                            break              
                    if(count==1):
                        continue
                    else:
                        l.append('不详')
            return l

        
    class Pythonword():
        def __init__(self):
            if not os.path.exists(Outdirname+'\\pic\\'):
                os.makedirs(Outdirname+'\\pic\\')
        def Initial(self, df):
            numdf = df['解决类型']
            indexs=[]
            for index, value in enumerate(numdf):
                if '非权属' != value:
                    indexs.append(index)
            scrdf = pd.DataFrame(df.iloc[indexs].values)
            scrdf.columns=df.columns
            return scrdf
        def scrdf(self,df,indexname,judge):#
            df0=df[indexname]
            indexs = []
            for index, value in enumerate(df0):
                if judge == value:
                    indexs.append(index)
            scrdf = pd.DataFrame(df.iloc[indexs].values)
            scrdf.columns=df.columns
            return scrdf

        def getChangeIndex(self,num1,num2,judge):#num1是上个月数据，num2是这个月的数据
            num=[]
            for i in range(len(num2)):
                for j in range(len(num1)):
                    if num2.index[i]==num1.index[j]:
                        if judge==0:
                            if num2.iloc[i]>num1.iloc[j] and (num2.iloc[i]-num1.iloc[j])>0.2:#如果这个月的数据大于上个月，就把这个诉求类型提取,judge=0求取上升的诉求
                                num.append(num2.index[i])
                        elif judge==1:
                            if num2.iloc[i]<num1.iloc[j] and (num1.iloc[j]-num2.iloc[i])>0.2:#judge=1求取下降的诉求
                                num.append(num2.index[i])
            return num
        def getChangeData(self,num1,num2):#num1是上个月数据，num2是这个月的数据
            box=[]
            for i in range(len(num2)):
                for j in range(len(num1)):
                    if num2.index[i]==num1.index[j]:
                        box.append([num2.index[i],num2.iloc[i],num1.iloc[j]])
            df=pd.DataFrame(box)[:10]
            df.columns=['诉求类型', str(int(Inmonth)) + '月数据', str(int(Inmonth) - 1) + '月数据']
            df.set_index('诉求类型',inplace=True)
            return df
        def list2str(self,list0):
            if list0:
                listend=[]
                for i in range(len(list0)):
                    listend.append(list0[i])
                    listend.append('、')
                del(listend[-1])
                str0="".join(listend)
                return str0
            else:
                return None
        def countVar(self,Tdf,Tdf1,Ldf,Ldf1):  # n1是前一个月的数据，n2是后一个月
            num1=[]
            num1.append(Ldf.shape[0])
            num1.append(Ldf.shape[0]-Ldf1.shape[0])
            num1.append(Ldf1.shape[0])
            num2=[]
            num2.append(Tdf.shape[0])
            num2.append(Tdf.shape[0] - Tdf1.shape[0])
            num2.append(Tdf1.shape[0])
            num = []
            for i in range(len(num2)):
                num.append(int(num2[i]))
            if num1[2] > num2[2]:
                num.append(int(num1[2] - num2[2]))
                num.append('下降')
                num.append(round(((num1[2] / num2[2] - 1) * 100), 2))
            else:
                num.append(int(num2[2] - num1[2]))
                num.append('上升')
                num.append(round(((num2[2] / num1[2] - 1) * 100), 2))
            return num

        def getformat2f(self,dfstu):
            count = [float('{:.4f}'.format(i)) for i in dfstu.values]
            count = [round(i * 100, 2) for i in count]
            return count

    class excel2word1(Pythonword):
        def __init__(self,lastmonth,thismonth):
            Pythonword.__init__(self)
            self.key1='诉求类型一级'
            self.key2='诉求类型二级'
            self.key3='诉求类型三级'
            self.keywater='河流名称'
            self.keyname0 = '解决类型'
            self.keyname1 = '非权属'
            self.Tdf = pd.read_excel(thismonth)
            self.Ldf = pd.read_excel(lastmonth)
            river=Search_river().search_river(self.Tdf)
            self.Tdf['河流名称']=river
            self.Tdf=pd.concat([self.Tdf,self.gettime(self.Tdf)],axis=1)
            river1 = Search_river().search_river(self.Ldf)
            self.Ldf['河流名称'] = river1
            self.Ldf = pd.concat([self.Ldf, self.gettime(self.Ldf)],axis=1)
            self.FTdf = self.Initial(self.Tdf)
            self.FLdf = self.Initial(self.Ldf)
            
        def getdict0(self):
            c0 = [f'c{n}' for n in range(6)]
            list=self.countVar(self.Tdf,self.FTdf,self.Ldf,self.FLdf)
            return dict(zip(c0,list)) 
        def getpic1(self):
            pt = self.getChangeData(self.FLdf[self.key1].value_counts(), self.FTdf[self.key1].value_counts())
            plt.rcParams['font.sans-serif'] = ['SimHei']
            pt.plot.bar()
            # plt.bar(pt.index,pt[str(int(Inmonth)) + '月数据'])
            # plt.bar(pt.index,pt[str(int(Inmonth) - 1) + '月数据'])
            x = range(0, len(pt.index.tolist()), 1)
            plt.xticks(x,rotation=60)
            plt.subplots_adjust(bottom=0.35)
            plt.title('直派水务局诉求')
            plt.savefig(Outdirname + '\\pic\\pic0.jpg')
            plt.close()
        def getdict1(self,lastlevel1,thislevel1):
            a0 = [f'a{n}' for n in range(10)]       
            list1 = thislevel1.index.tolist()[:6]
            listc0 = self.getChangeIndex(lastlevel1, thislevel1,0)
            listc1 = self.getChangeIndex(lastlevel1, thislevel1,1)
            list1.append(self.list2str(listc0))
            list1.append(self.list2str(listc1))
            a0dict = dict(zip(a0, list1))
            return a0dict
        def getdict2(self,valuecounts,valuecountsp):
            list0=valuecounts.values.tolist()[:6]
            list1=self.getformat2f(valuecountsp)
            list0.extend(list1)
            a0 = [f'b{n}' for n in range(len(list0))]
            return dict(zip(a0,list0))
        
        def gettime(self,df):
            time=[]
            timestamp=df['登记时间']
            for i in range(len(timestamp)):
                time.append([timestamp.iloc[i][:4],timestamp.iloc[i][5:7],timestamp.iloc[i][8:10],timestamp.iloc[i][11:13]])
            dftime = pd.DataFrame(time)
            dftime.columns = ['年', '月', '日','时']
            return dftime
        def getdicttime(self,cbi):
            listtime=[]
            # listtimeh=[]
            for i in range(len(cbi)):
                listtime.extend(self.get2levelData(cbi[i],'日',3)[:6])
                # listtimeh.extend(self.get2levelData(cbi[i], '时', '时', 2)[:4])
            a0 = [f'time{n}' for n in range(len(listtime))]
            # ah = [f'timeh{n}' for n in range(len(listtimeh))]
            new_dict={}
            new_dict.update(dict(zip(a0, listtime)))
            # new_dict.update(dict(zip(ah, listtimeh)))
            return new_dict
        def get2levelData(self,tmindex,key,num):
            la0=[]
            dfvc0=self.scrdf(self.Tdf,self.key1, tmindex)[key].value_counts()
            if '其它' in dfvc0.index.tolist():
                dfvc0=dfvc0.drop('其它',axis=0)
                # dfvc00=dfvc00.drop('其它',axis=0)
            c0a =dfvc0.index.tolist()[:num]
            c0a.extend('' for _ in range(num-len(c0a)))
            # if key2==key3:
            c0a1 = dfvc0.values.tolist()[:num]
            c0a1.extend('' for _ in range(num-len(c0a1)))
            la0.extend(c0a)
            la0.extend(c0a1)
            # la0.extend(c0a2)
            return la0
        def getleveldata(self,ldf,tdf,judgename):#ldf是上个月的数据，tdf是这个月得数据，judgename是一级诉求里判断条件，二级诉求
            dictlist=[]
            dfv0=self.scrdf(tdf,'诉求类型一级',judgename)['诉求类型二级'].value_counts()
            dfv00=self.scrdf(tdf,'诉求类型一级',judgename)['诉求类型二级'].value_counts(normalize=True)
            if '其它' in dfv0.index.tolist():
                dfv0=dfv0.drop('其它',axis=0)
                dfv00=dfv00.drop('其它',axis=0)
            dfv00=self.getformat2f(dfv00)
            if len(dfv0.index.tolist())>=2:
                dictlist.extend(dfv0.index.tolist()[:2])
                dictlist.extend(dfv0.values.tolist()[:2])
                dictlist.extend(dfv00[:2])
            else:
                dictlist.extend(dfv0.index.tolist())
                dictlist.extend('' for _ in range(2-len(dfv0.index.tolist())))
                dictlist.extend(dfv0.values.tolist())
                dictlist.extend(0 for _ in range(2-len(dfv0.values.tolist())))
                dictlist.extend(dfv00)
                dictlist.extend(0 for _ in range(2-len(dfv00)))
            # print(dictlist)
            dfvl0=self.scrdf(ldf,'诉求类型一级',judgename)['诉求类型二级'].value_counts()

            if len(dfv0.index.tolist())>=1:
                for i in range(len(dfvl0)):
                    if dfv0.index[0]==dfvl0.index[i]:
                        if dfv0.iloc[0]>dfvl0.iloc[i]:
                            dictlist.append('增长')
                            dictlist.append(dfv0.iloc[0]-dfvl0.iloc[i])
                        else:
                            dictlist.append('下降')
                            dictlist.append(dfvl0.iloc[i]-dfv0.iloc[0])
            else:
                dictlist.extend(['',''])
            if len(dfv0.index.tolist())>=2:
                for i in range(len(dfvl0)):
                    if dfv0.index[1]==dfvl0.index[i]:
                        if dfv0.iloc[1]>dfvl0.iloc[i]:
                            dictlist.append('增长')
                            dictlist.append(dfv0.iloc[1]-dfvl0.iloc[i])
                        else:
                            dictlist.append('下降')
                            dictlist.append(dfvl0.iloc[i]-dfv0.iloc[1])
            else:
                dictlist.extend(['',''])
            
            dictlist.append(self.list2str(self.scrdf(tdf,'诉求类型二级',dictlist[0])['诉求类型三级'].value_counts().index.tolist()))
            dictlist.append(self.list2str(self.scrdf(tdf,'诉求类型二级',dictlist[1])['诉求类型三级'].value_counts().index.tolist()))
            # print(dictlist)
            return dictlist
        def cbdw(self,df,cbi,name):
            strlist=[]
            cbdfindex=[]
            if name=='主办单位':          
                cbdf=self.scrdf(df,'诉求类型一级',cbi)[name].value_counts()
                if len(cbdf)>=5:
                    for i in range(5):
                        cbdfindex.append(cbdf.index.tolist()[i])
                        cbdfindex.append('（')
                        cbdfindex.append(str(cbdf.values.tolist()[i]))
                        cbdfindex.append('）件')
                        cbdfindex.append('、')
                    del(cbdfindex[-1])
                    str0="".join(cbdfindex)
                    strlist.append(str0)
                else:
                    for i in range(len(cbdf)):
                        cbdfindex.append(cbdf.index.tolist()[i])
                        cbdfindex.append('（')
                        cbdfindex.append(str(cbdf.values.tolist()[i]))
                        cbdfindex.append('）件')
                        cbdfindex.append('、')
                    del(cbdfindex[-1])
                    str0="".join(cbdfindex)
                    strlist.append(str0)
            elif name=='河流名称':
                cb=self.scrdf(df,'诉求类型一级',cbi)['诉求类型二级'].value_counts().index.tolist()
                if len(cb)<=1:
                    cbdf=self.scrdf(df,'诉求类型二级',cb[0])['河流名称'].value_counts()
                    if len(cbdf)>=5:
                        for i in range(5):
                            cbdfindex.append(cbdf.index.tolist()[i])
                            cbdfindex.append('（')
                            cbdfindex.append(str(cbdf.values.tolist()[i]))
                            cbdfindex.append('）件')
                            cbdfindex.append('、')
                        del(cbdfindex[-1])
                        str0="".join(cbdfindex)
                        strlist.append(str0)
                    else:
                        for i in range(len(cbdf)):
                            cbdfindex.append(cbdf.index.tolist()[i])
                            cbdfindex.append('（')
                            cbdfindex.append(str(cbdf.values.tolist()[i]))
                            cbdfindex.append('）件')
                            cbdfindex.append('、')
                        del(cbdfindex[-1])
                        str0="".join(cbdfindex)
                        strlist.append(str0)
                if len(cb)>=2:
                    cbdf=self.scrdf(df,'诉求类型二级',cb[1])['河流名称'].value_counts()
                    if len(cbdf)>=5:
                        for i in range(5):
                            cbdfindex.append(cbdf.index.tolist()[i])
                            cbdfindex.append('（')
                            cbdfindex.append(str(cbdf.values.tolist()[i]))
                            cbdfindex.append('）件')
                            cbdfindex.append('、')
                        del(cbdfindex[-1])
                        str0="".join(cbdfindex)
                        strlist.append(str0)
                    else:
                        for i in range(len(cbdf)):
                            cbdfindex.append(cbdf.index.tolist()[i])
                            cbdfindex.append('（')
                            cbdfindex.append(str(cbdf.values.tolist()[i]))
                            cbdfindex.append('）件')
                            cbdfindex.append('、')
                        del(cbdfindex[-1])
                        str0="".join(cbdfindex)
                        strlist.append(str0)

            return strlist
        def get_row_data(self,name):
            df=pd.read_excel('./data/诉求风险趋势变化表.xlsx')[-32:]
            ptdf = df.iloc[df['全部诉求'].values.tolist().index(name)]
            datadf = ptdf.values.tolist()
            datadf.remove(name)
            datadf = list(map(int, datadf))
            return datadf
        def compareDict(self,tli):
            a0a0=[]
            for i in range(5):
                monthdata=self.get_row_data(tli[i])
                lastY=monthdata[-13]
                lastM=monthdata[-2]
                thisM=monthdata[-1]
                a0a0.append(self.compareLastMandY(tli[i],lastM,lastY,thisM))
            a0 = [f'aa{n}' for n in range(len(a0a0))]
            return dict(zip(a0,a0a0))
        def compareLastMandY(self,leiname,lastM,lastY,thisM):
            if thisM> lastM and thisM >lastY:
                if lastY!=0:
                    stra0=leiname+'类诉求较上月增加'+str(thisM-lastM)+'件，较去年同期增加'+str(thisM-lastY)+'件，增长率为'+"{:.2f}".format((thisM-lastY)/lastY*100)+'%。'
                else:
                    stra0 = leiname + '类诉求较上月增加' + str(thisM - lastM) + '件，较去年同期增加' + str(thisM - lastY) + '件，下降率为' + 'None。'
            elif thisM> lastM and thisM < lastY:
                if lastY!=0:
                    stra0=leiname+'类诉求较上月增加'+str(thisM-lastM)+'件，较去年同期减少'+str(lastY-thisM)+'件，下降率为'+"{:.2f}".format((lastY-thisM)/lastY*100)+'%。'
                else:
                    stra0 = leiname + '类诉求较上月增加' + str(thisM - lastM) + '件，较去年同期减少' + str(lastY - thisM) + '件，下降率为' +'None。'
            elif thisM < lastM and thisM > lastY:
                if lastY!=0:
                    stra0=leiname+'类诉求较上月减少'+str(lastM-thisM)+'件，较去年同期增加'+str(thisM-lastY)+'件，增长率为'+"{:.2f}".format((thisM-lastY)/lastY*100)+'%。'
                else:
                    stra0 = leiname + '类诉求较上月减少' + str(lastM-thisM) + '件，较去年同期增加' + str(thisM - lastY) + '件，增长率为' +'None。'
            elif thisM < lastM and thisM < lastY:
                if lastY!=0:
                    stra0=leiname+'类诉求较上月减少'+str(lastM-thisM)+'件，较去年同期减少'+str(lastY-thisM)+'件，下降率为'+"{:.2f}".format((lastY-thisM)/lastY*100)+'%。'
                else:
                    stra0 = leiname + '类诉求较上月减少' + str(lastM-thisM) + '件，较去年同期减少' + str(lastY - thisM) + '件，下降率为' +'None。'
            else:
                stra0=None
            return stra0
        def getConstrast(self,name,idx):
            def draw_line_base(month,datadf,yearlist,length):
            # plt.figure()
                plt.plot()
                step = int(length/12)
                # print(length,step)
                for i in range(step):
                    # print("%d",i*12)
                    plt.plot(month[:12],datadf[int(i*12):int((i+1)*12)], marker='o', markersize=3)
                    for a, b in zip(month[:12], datadf[int(i*12):int((i+1)*12)]):
                        plt.text(a, b, b, ha='center', va='bottom', fontsize=10)  # 设置数据标签位置及大小
                plt.plot(month[:length-step*12],datadf[step*12:length],marker='o',markersize=3)
                for a, b in zip(month[:length-step*12], datadf[step*12:length]):
                    plt.text(a, b, b, ha='center', va='bottom', fontsize=10)
                plt.legend(yearlist[:step+1])  # 设置折线名称
            # plt.show()
            df=pd.read_excel('./data/诉求风险趋势变化表.xlsx')[-32:]
            plt.rcParams['font.sans-serif'] = ['SimHei']
            ptdf = df.iloc[df['全部诉求'].values.tolist().index(name)]
            datadf = ptdf.values.tolist()
            datadf.remove(name)
            datadf = list(map(int, datadf))
            yearlist=['2020年','2021年','2022年','2023年','2024年','2025年','2026年','2027年','2028年']
            month = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']
            length=len(datadf)
            draw_line_base(month,datadf,yearlist,length)
            plt.title(label=name)
            plt.xlabel('时间')  # x轴标题
            plt.ylabel('诉求值')  # y轴标题
            plt.savefig(Outdirname  + '\\pic\\pica'+ str(idx) + '.jpg')
            plt.close()

        def getWordmain(self):
            self.getpic1()
            thislevel1 = self.FTdf[self.key1].value_counts().drop('其它',axis=0)
            thislevel1p = self.FTdf[self.key1].value_counts(normalize=True).drop('其它',axis=0)
            lastlevel1 = self.FLdf[self.key1].value_counts().drop('其它',axis=0)
            thislevel1i = thislevel1.index.tolist()
            list0=[]
            list1=[]
            list2=[]
            for i in range(5):
                list0.extend(self.getleveldata(self.FLdf,self.FTdf,thislevel1i[i]))
                self.getConstrast(thislevel1i[i],i)
                list1.extend(self.cbdw(self.FTdf,thislevel1i[i],'河流名称'))
                list2.extend(self.cbdw(self.FTdf,thislevel1i[i],'主办单位'))
                
            # dict0=zip(dict([f'a0{n}' for n in range(len(list0))],list0))
            new_dict = {}
            # new_dict.update(self.getMonthTime())
            new_dict.update(self.getdict0())
            new_dict.update(dict(zip([f'a0{n}' for n in range(len(list0))],list0)))
            new_dict.update(dict(zip([f'ab{n}' for n in range(len(list1))],list1)))
            new_dict.update(dict(zip([f'ac{n}' for n in range(len(list2))],list2)))
            new_dict.update(self.getdict1(lastlevel1, thislevel1))
            new_dict.update(self.getdict2(thislevel1, thislevel1p))
            new_dict.update(self.compareDict(thislevel1i))
            new_dict.update(self.getdicttime(thislevel1i))
            return new_dict

    class excel2word2(Pythonword):
        def __init__(self,Lmonth,Tmonth):
            Pythonword.__init__(self)
            self.tdf = pd.read_excel(Tmonth,sheet_name='区街道乡镇')
            self.ldf = pd.read_excel(Lmonth,sheet_name='区街道乡镇')
            self.cbcounts = self.tdf['承办单位'].value_counts()


        def getdict1(self):
            chengban1 = '北京市自来水集团有限责任公司'
            chengban2 = '北京城市排水集团有限责任公司'
            d0=[]
            d0.append(self.tdf.shape[0])
            if chengban1 in self.cbcounts.index.tolist():
                if chengban2 in self.cbcounts.index.tolist():
                    d0.append(self.cbcounts.loc[chengban1])
                    d0.append(self.cbcounts.loc[chengban2])
            listd0 = [f'd{n}' for n in range(3)]
            dict1 = dict(zip(listd0,d0))
            return dict1
        def getdict2(self,cbs,cbp):
            cbi = cbs.index.tolist()[:10]
            cbv=cbs.values.tolist()[:10]
            cbi.extend(cbv)
            list1 = self.getformat2f(cbp)
            cbi.extend(list1)
            listd00=[f'd0{n}' for n in range(len(cbi))]
            dict2=dict(zip(listd00,cbi))
            return dict2
        def getpic2(self,pt):
            plt.rcParams['font.sans-serif'] = ['SimHei']
            pt.plot.bar()
            # plt.bar(pt.index,pt[str(int(Inmonth)) + '月数据'])
            # plt.bar(pt.index,pt[str(int(Inmonth) - 1) + '月数据'])
            x = range(0, len(pt.index.tolist()), 1)
            plt.xticks(x,rotation=60)
            plt.subplots_adjust(bottom=0.35)
            plt.title('直派区街道乡镇诉求')
            plt.savefig(Outdirname +'\\pic\\pic1.jpg')
        def getdict3(self):
            num=[]
            num.append(self.list2str(self.getChangeIndex(self.ldf['三级问题'].value_counts(),self.tdf['三级问题'].value_counts(),0)))
            num.append(self.list2str(self.getChangeIndex(self.ldf['三级问题'].value_counts(),self.tdf['三级问题'].value_counts(),1)))
            listd0 = [f'd00{n}' for n in range(2)]
            dict3=dict(zip(listd0,num))
            return dict3
        def getpic(self,num0,num,cbi):
            pt = self.getChangeData(self.scrdf(self.ldf, '三级问题', cbi)['被反映区'].value_counts(),self.scrdf(self.tdf, '三级问题', cbi)['被反映区'].value_counts())
            plt.rcParams['font.sans-serif'] = ['SimHei']
            pt.plot.bar()
            # plt.bar(pt.index,pt[str(int(Inmonth)) + '月数据'])
            # plt.bar(pt.index,pt[str(int(Inmonth) - 1) + '月数据'])
            x = range(0, len(pt.index.tolist()), 1)
            plt.xticks(x, rotation=45)
            plt.subplots_adjust(bottom=0.35)
            plt.title(cbi)
            plt.savefig(Outdirname+ '\\pic\\pic'+str(num0+2)+'.jpg')
            quname=['密云区','延庆区','朝阳区','丰台区','石景山区','海淀区','门头沟区','房山区','通州区','顺义区','昌平区','大兴区','怀柔区','平谷区','东城区','西城区']
            ptindex=self.scrdf(self.tdf, '三级问题', cbi)['被反映区'].value_counts().index.tolist()
            # ptindex=list(set(ptindex+quname))
            ptindex=list((ptindex+ [item for item in quname if str(item) not in ptindex]))
            ptvalue=self.scrdf(self.tdf, '三级问题', cbi)['被反映区'].value_counts().values.tolist()
            if len(ptindex)!=len(ptvalue):
                ptvalue.extend(0 for _ in range(len(ptindex)-len(ptvalue)))
            maxvalue=max(map(int,ptvalue))
            while (maxvalue %5):
                maxvalue+=1
            c = (
                Map()
                .add("", [list(z) for z in zip(ptindex, ptvalue)], "北京",zoom=1.25,aspect_scale=0.9,layout_center=['60%','60%'],
                label_opts=opts.LabelOpts(is_show=True,position='outsideleft',font_size=10,color='#708069',rotate = '30',horizontal_align = 'center',font_weight='bold',vertical_align ='middle'))
                .set_global_opts(
                title_opts=opts.TitleOpts(title=cbi+'类诉求'),
                visualmap_opts=opts.VisualMapOpts(type_='color', orient='vertical',max_=maxvalue,
                                                    is_piecewise=True,range_color=['#71ae9b','#8bb08e', '#c1d389','#e3d76b','#f7b44a','#f27127','#cc8775']
                ))
            )
            make_snapshot(driver, c.render(), Outdirname +'\\pic\\picb'+ str(num)+".png")
            b = (
                Map()
                .add("", [list(z) for z in zip(ptindex, ptvalue)], "北京",zoom=1.25,aspect_scale=0.9,layout_center=['60%','60%'],
                label_opts=opts.LabelOpts(is_show=True,position='outsideleft',font_size=10,color='#708069',rotate = '30',horizontal_align = 'center',font_weight='bold',vertical_align ='middle'))
                .set_global_opts(
                title_opts=opts.TitleOpts(title=cbi+'类诉求'),
                visualmap_opts=opts.VisualMapOpts(type_='color', orient='vertical',is_piecewise=True,pieces=[{'min':0,'max':10,'label':'[0-10)',"color":'#71ae9b'},{'min':11,'max':30,'label':'[10-30)',"color":'#8bb08e'},{'min':31,'max':100,'label':'[30-100)',"color":'#c1d389'},{'min':101,'max':200,'label':'[100-200)',"color":'#e3d76b'},{'min':201,'max':350,'label':'[200-350)',"color":'#f7b44a'},
                                                            {'min':350,'label':'[350,)',"color":'#f27127'}]
                                                            # {'min':0,'max'=10,label:'0-10','#cc8775']
                                                    )
            ))
            make_snapshot(driver, b.render(), Outdirname +'\\pic\\picb'+ str(num+1)+".png")
            

        def get1level(self,indexname,var,key1,key2,num):
            lc0a = [str(var)+f'{n}' for n in range(3*num)]
            c0a = self.get2levelData(indexname, key1, key2,num)
            dict1 = dict(zip(lc0a, c0a))
            return dict1

        def get2levelData(self,tmindex,key1,key2,num):
            la0=[]
            c0a = self.scrdf(self.tdf,key1, tmindex)[key2].value_counts().index.tolist()[:num]
            c0a.extend('' for _ in range(num-len(c0a)))
            c0a1 = self.scrdf(self.tdf,key1, tmindex)[key2].value_counts().values.tolist()[:num]
            c0a1.extend('' for _ in range(num-len(c0a1)))
            c0a2 = self.getformat2f(self.scrdf(self.tdf,key1, tmindex)[key2].value_counts(normalize=True))
            la0.extend(c0a)
            la0.extend(c0a1)
            la0.extend(c0a2)
            return la0

        def getdict4(self,num1,num2):#num1是上个月的数据
            num=[]
            for i in num2.index.tolist():
                for j in num1.index.tolist():
                    if i==j:
                        if num1.loc[j] > num2.loc[i]:
                            num.append('下降')
                            num.append(int(num1.loc[i] - num2.loc[i]))
                        else:
                            num.append('上升')
                            num.append(int(num2.loc[i] - num1.loc[i]))
            listdd = [f'd0d{n}' for n in range(len(num))]
            return dict(zip(listdd,num))
        def getdict51(self,cbi):
            listdict=[]
            vc1=self.scrdf(self.ldf, '三级问题', cbi)['被反映区'].value_counts()
            vc2=self.scrdf(self.tdf, '三级问题', cbi)['被反映区'].value_counts()

            num0=self.getChangeIndex(vc1,vc2,0)
            num1=self.getChangeIndex(vc1,vc2,1)
            listdict.append(self.list2str(num0))
            listdict.append(self.list2str(num1))
            return listdict
        def getdict5(self,cbi):
            num=[]
            for i in range(len(cbi)):
                num.extend(self.getdict51(cbi[i]))
            listdd = [f'd0a{n}' for n in range(len(num))]
            return dict(zip(listdd, num))
        def getWordmain(self):
            name = '三级问题'
            key1='被反映区'
            cbcounts = self.tdf[name].value_counts()
            cbi=cbcounts.index.tolist()
            cbcountsp = self.tdf[name].value_counts(normalize=True)
            pt = self.getChangeData(self.ldf[name].value_counts(),
                                    self.tdf[name].value_counts())#三级问题获得counts
            self.getpic2(pt)
            num=[0,2,4,6,8,10,12,14,16,18]
            for i in range(10):
                self.getpic(i,num[i],cbi[i])
            new_dict={}
            new_dict.update(self.getdict1())
            new_dict.update(self.getdict2(cbcounts,cbcountsp))
            new_dict.update(self.getdict3())
            new_dict.update(self.get1level(cbi[0],'da',name,key1,5))
            new_dict.update(self.get1level(cbi[1], 'db', name, key1, 5))
            new_dict.update(self.get1level(cbi[2], 'dc', name, key1, 5))
            new_dict.update(self.get1level(cbi[3], 'dd', name, key1, 5))
            new_dict.update(self.get1level(cbi[4], 'de', name, key1, 5))
            new_dict.update(self.get1level(cbi[5], 'df', name, key1, 5))
            new_dict.update(self.get1level(cbi[6], 'dg', name, key1, 5))
            new_dict.update(self.get1level(cbi[7], 'dh', name, key1, 5))
            new_dict.update(self.get1level(cbi[8], 'di', name, key1, 5))
            new_dict.update(self.get1level(cbi[9], 'dj', name, key1, 5))
            new_dict.update(self.getdict4(self.ldf[name].value_counts(),
                                self.tdf[name].value_counts()))
            new_dict.update(self.getdict5(cbi[:10]))
            return new_dict

    def getMonthTime(month):
            time = ['t0', 't1', 't2', 't3']
            month2 = ['1', '19', '2', '18']
            month3 = ['2', '19', '3', '18']
            month4 = ['3', '19', '4', '18']
            month5 = ['4', '19', '5', '18']
            month6 = ['5', '19', '6', '18']
            month7 = ['6', '19', '7', '18']
            month8 = ['7', '19', '8', '18']
            month9 = ['8', '19', '9', '18']
            month10 = ['9', '19', '10', '18']
            month11 = ['10', '19', '11', '18']
            month12 = ['11', '19', '12', '18']
            month1 = ['12', '19', '1', '18']
            if month == 1:
                month_dict = dict(zip(time, month1))
                return month_dict
            elif month == 2:
                month_dict = dict(zip(time, month2))
                return month_dict
            elif month == 3:
                month_dict = dict(zip(time, month3))
                return month_dict
            elif month == 4:
                month_dict = dict(zip(time, month4))
                return month_dict
            elif month == 5:
                month_dict = dict(zip(time, month5))
                return month_dict
            elif month == 6:
                month_dict = dict(zip(time, month6))
                return month_dict
            elif month == 7:
                month_dict = dict(zip(time, month7))
                return month_dict
            elif month == 8:
                month_dict = dict(zip(time, month8))
                return month_dict
            elif month == 9:
                month_dict = dict(zip(time, month9))
                return month_dict
            elif month == 10:
                month_dict = dict(zip(time, month10))
                return month_dict
            elif month == 11:
                month_dict = dict(zip(time, month11))
                return month_dict
            elif month == 12:
                month_dict = dict(zip(time, month12))
                return month_dict
            else:
                return None
                
                # if __name__=='__main__':
    import time
    asset_url = './data/模板.docx'
    context1=excel2word1(Infilename1,Infilename2).getWordmain()
    context2=excel2word2(Infilename3,Infilename4).getWordmain()
    tpl = DocxTemplate(asset_url)
    picture={}
    for i in range(12):
        picture.update({'pic'+str(i): InlineImage(tpl, Outdirname+'\\pic\\pic'+str(i)+'.jpg', width=Mm(130), height=Mm(100))})
    for i in range(5):
        picture.update({'pica' + str(i): InlineImage(tpl, Outdirname+'\\pic\\pica'+str(i) + '.jpg', width=Mm(130), height=Mm(100))})
    for i in range(20):
        picture.update({'picb' + str(i): InlineImage(tpl, Outdirname+'\\pic\\picb'+str(i) + '.png', width=Mm(160), height=Mm(130))})
    
    # print(context2)

    context1.update(getMonthTime(int(Inmonth)))
    context1.update(picture)
    context1.update(context2)
    tpl.render(context1)
    tpl.save(Outdirname+'\\涉水数据' + str(int(Inmonth))+'月份数据分析报告'+time.strftime("%Y%m%d")+'.docx')
