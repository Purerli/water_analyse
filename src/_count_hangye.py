from operator import index
import pandas as pd
import time
index_df="D:\\Users\\MR\\Desktop\\北京市水务局涉水数据分析软件v1.1\\直派乡镇（区县）\\2022.4.xlsx"
index_value_counts=pd.read_excel(index_df)['三级问题'].value_counts()
index_list=pd.read_excel(index_df)['三级问题'].value_counts().index.tolist()

month_list=[]
years=[2020,2021]
for i in range(len(years)):
    for j in range(1,13):
        month_list.append(str(years[i])+"."+str(j))
month_list.extend(['2022.1','2022.2','2022.3','2022.4'])
short_path="D:\\Users\\MR\\Desktop\\北京市水务局涉水数据分析软件v1.1\\直派乡镇（区县）\\"
out_df=pd.DataFrame(index=index_list,columns=month_list)
# print(out_df)
def get_index_values(nums):
    nums_index=nums.index.tolist()
    
    box=[]
    for i in range(len(index_list)):
        if index_list[i] in nums_index:
            box.append(nums.iloc[nums_index.index(index_list[i])])
        else:
            box.append(0)
    return box
for i in range(len(month_list)):
    end_path=short_path+month_list[i]+".xlsx"
    try:
        df_values=pd.read_excel(end_path,sheet_name="区街道乡镇")['三级问题'].value_counts()
        out_df[month_list[i]]=get_index_values(df_values)
    except:
        df_values=pd.read_excel(end_path)['三级问题'].value_counts()
        out_df[month_list[i]]=get_index_values(df_values)
        # out_df=df.groupby(index_df['三级问题']).value_counts()
        
savepath =short_path + "行业诉求类型统计"+time.strftime("%Y%m%d")+".xlsx"
writer = pd.ExcelWriter(savepath)
out_df.to_excel(writer, sheet_name="Sheet1", index=True)
writer.save()
