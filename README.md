# water_analyse 项目来源于市水务局涉水投诉数据
## 开始
```
git clone https://github.com/Purerli/water_analyse.git
cd src && python3 main_gui.py
scp -r b1023@172.25.19.163:/home/b1023/直派水务局/ c:\\你的存放路径
scp -r b1023@172.25.19.163:/home/b1023/直派乡镇（区县）/ c:\\你的存放路径
//scp这两步只能在校园网下，且不消耗流量，密码：实验室号
```
## 开发
+ _attribute_river_lake_village.py实现河湖属性增补
+ _count_draw.py实现统计分析与绘制图形
+ _count_hangye.py实现按月行业统计
+ _similar_predict.py实现相似性分析和趋势分析报告
+ _extract_reason.py实现原因提取
+ _intelligent_classification.py实现智能分类
+ _month_report.py实现每月月报
+ main_gui.py图形化界面接口
## data文件夹下存放依赖
## demo_excel模板化表格
