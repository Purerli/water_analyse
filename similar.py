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

if __name__=="__main__":
    infile=r'D:\Users\MR\Desktop\北京市水务局涉水数据分析软件v1.1\直派水务局\2022.4.xlsx'
    outfile=r'D:\Users\MR\Desktop\北京市水务局涉水数据分析软件v1.1'
    similar_ayalyse(Infilename=infile,Outdirname=outfile)