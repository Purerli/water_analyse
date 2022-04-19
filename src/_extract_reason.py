def Extract_reason(Infilename=None,Outdirname=None):
    import difflib

    import jieba
    jieba.set_dictionary("./data/dict.txt")
    jieba.initialize()
    import pandas as pd

    def get_equal_rate(str1, str2):
        return difflib.SequenceMatcher(None, str1, str2).quick_ratio()

    class Nstr:
        def __init__(self, arg):
            self.x = arg

        def __sub__(self, other):
            c = self.x.replace(other.x, "")
            return c

    def frequency_sort(items):
        # your code here
        lst1 = []
        for i in items:
            if i not in lst1:
                lst1.append(i)
        lst = []
        dic = {}
        # # print(lst1)
        for i in lst1:
            dic[i] = items.count(i)
        dic = dict(sorted(dic.items(), key=lambda x: x[1], reverse=True))
        # #print(dic)
        count = 0
        for k, v in dic.items():
            for i in range(v):
                lst.append(k)
        # # print(lst)
        return lst
    file_name = Infilename
    data = pd.read_excel(file_name, sheet_name='Sheet1')
    jieguo = data["处理结果"]
    mainkey = ["因", "因为", "由", "由于", "原因"]
    listkeyword = []
    resendlist = []
    reasonlist = []
    for i in range(0, len(data)):
        prores = jieguo[i]
        if str(prores) == "None":
            resend = "无数据"
        else:
            seg_list = jieba.cut(prores)
            seg = list(seg_list)
            # print(seg)
            for keyword in mainkey:
                for item in seg:
                    if item == keyword:
                        listkeyword.append(item)
            listkeyword = frequency_sort(listkeyword)
            if len(listkeyword) != 0:
                proresnew = prores
                for keyw in listkeyword:
                    keywordcount = listkeyword.count(keyw)
                    if keywordcount != 1:
                        findk = proresnew.find(keyw)
                        finde = proresnew.find("。", findk)
                        resend = proresnew[findk:finde + 1]
                        resendlist.append(resend)
                        proresnew = proresnew[finde + 1:]
                    else:
                        findkk = prores.find(keyw)
                        findee = prores.find("。", findkk)
                        resendd = prores[findkk:findee + 1]
                        resendlist.append(resendd)
                strresend = ""
                for i in resendlist:
                    strresend = strresend + i
                resend = strresend
            else:
                resend = "未找到"
        reasonlist.append(resend)
        resendlist = []
        listkeyword = []
        # print(listkeyword)
    data['诉求原因'] = reasonlist
    savepath = Outdirname + "\\诉求原因提取.xlsx"
    writer = pd.ExcelWriter(savepath)
    data.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.save()