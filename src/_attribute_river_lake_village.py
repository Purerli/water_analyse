import pandas as pd
import sys
def Extract_river(data_type='直派水务局',Infilename=None,Outdirname=None):
        # Infilename
        # output_filename=OutDirname
    import codecs
    import re
    import jieba
    jieba.set_dictionary("./data/dict.txt")
    jieba.initialize()
    rDict_filename='./data/river.txt' # 河流名称字典
    stopkey_filename= './data/stopWord_river.txt'# 停用词字典
    tDict_filename='./data/total_dict.txt' # 自建水务词库

    def open_txt(dict):
        '''
        打开文本文档
        '''
        dict=[w.strip() for w in codecs.open(dict, 'r',encoding="utf-8").readlines()] 
        return dict
    
    def txt_read(files):
      txt_dict = {}
      count=0
      fopen = open(files,encoding='utf-8')
      for line in fopen.readlines():
        line = str(line).replace("\n","")
        txt_dict[line.split(' ',1)[0]] = line.split(' ',1)[1]
        #split（）函数用法，逗号前面是以什么来分割，后面是分割成n+1个部分，且以数组形式从0开始
      fopen.close()
      return txt_dict


    def get_content(Infilename):
       '''
       读取excel中指定两列的内容
       '''
       df = pd.read_excel(Infilename)
       content1=[]
       content=[]
        # 按行遍历，读入每行中的两列数据，合并成一个list中一个元素
      # for row in range(2,ws.max_row+1):
    #    for row in range(2,ws.max_row+1):
    #        for col in range(17,18):
    #          content1.append((ws.cell(row,col).value))
    #        content.append(''.join(content1))
    #        content1.clear()
       for i in range(len(df)):
           content1.append(str(df['诉求内容'].iloc[i]))
           content1.append(str(df['主办单位'].iloc[i]))
           
       for j in range(len(content1)):
            if j%2==0:
             content.append(''.join(content1[j:j+2]))
            else:
                continue
       return content
    
    def data_clean():
        '''
        去停用词
        '''
        jieba.load_userdict(tDict_filename) # 添加自建词库
        content=get_content(Infilename)
        stopkey=open_txt(stopkey_filename)
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

    def search_river():
     '''
     在字典中查找河流名称
     '''
     l = []
     suggestion=data_clean()
     riverdict=txt_read(rDict_filename)
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
        
    # if __name__=='__main__':
    '''
    输出结果到excel
    '''
    df = pd.read_excel(Infilename)
    df['具体点位']=search_river()
    savepath = Outdirname + "\\直派水务局河湖增补.xlsx"  #导入excel数据，设置表头标题
    writer = pd.ExcelWriter(savepath)   #设置导出的excel文件地址
    df.to_excel(writer, sheet_name="Sheet1", index=False)  
    writer.save()
    sys.exit()
def Extract_village(data_type='直派水务局',Infilename=None,Outdirname=None):
    import jieba
    import pandas as pd
    jieba.set_dictionary("./data/dict.txt")
    jieba.initialize()
    import difflib
    class Nstr:
        def __init__(self, arg):
            self.x = arg

        def __sub__(self, other):
            c = self.x.replace(other.x, "")
            return c

    def get_equal_rate(str1, str2):
        return difflib.SequenceMatcher(None, str1, str2).quick_ratio()

    def getpoint(strindex):
        global resualt
        cunming = seg[strindex]
        scorelist1 = []
        scorelist2 = []
        scorelist3 = []
        scorelist4 = []
        for cun in item_seg:
            scorelist1.append(get_equal_rate(cunming, cun))
        # print(max(scorelist1))
        # print(len(cunming))
        # print(item_seg[scorelist1.index(max(scorelist1))])
        if max(scorelist1) == 1 and len(cunming) != 1:
            cunmingindex = scorelist1.index(max(scorelist1))
            cunming1 = item_seg[cunmingindex]
            resualt = cunming1
        else:
            cunming = seg[strindex - 1] + seg[strindex]
            for cun in item_seg:
                scorelist2.append(get_equal_rate(cunming, cun))
            # print(max(scorelist2))
            if max(scorelist2) > 0.9:
                cunmingindex = scorelist2.index(max(scorelist2))
                cunming2 = item_seg[cunmingindex]
                resualt = cunming2
            else:
                cunming = seg[strindex - 2] + seg[strindex - 1] + seg[strindex]
                for cun in item_seg:
                    scorelist3.append(get_equal_rate(cunming, cun))
                # print(max(scorelist3))
                if max(scorelist3) > 0.8:
                    cunmingindex = scorelist3.index(max(scorelist3))
                    cunming3 = item_seg[cunmingindex]
                    resualt = cunming3
                else:
                    cunming = seg[strindex - 3] + seg[strindex - 2] + seg[strindex - 1] + seg[strindex]
                    for cun in item_seg:
                        scorelist4.append(get_equal_rate(cunming, cun))
                    # print(max(scorelist4))
                    if max(scorelist4) > 0.75:
                        cunmingindex = scorelist4.index(max(scorelist4))
                        cunming4 = item_seg[cunmingindex]
                        resualt = cunming4
                    else:
                        resualt = "未找到"
        return resualt

    atxt = r"./data/北京市全部地名.txt"
    btxt = r'./data/北京市街道.txt'
    ctxt = r"./data/北京市城区.txt"
    zhuizhongres = []
    qulist = []
    listzhen = []
    countlist = []
    xiangzhen = []
    mainkey = ["村", "小区", "胡同", "社区", "里", "路", "园", "院", "楼", "区"]
    keywordlibiao = ["村", "小区", "胡同", "社区", "里", "路", "园", "院", "楼", "区"]
    global seg
    global item_seg
    jieba.load_userdict(atxt)
    item_seg = []
    with open(atxt, encoding='utf-8') as f:
        for line in f:
            item_seg.append(line.strip('\n'))
    with open(btxt, encoding='utf-8') as f:
        for line in f:
            line = line.strip('\n')
            xiangzhen.append(line)
    with open(ctxt, encoding='utf-8') as cleardata:
        for qu in cleardata:
            qulist.append(qu.strip('\n'))
    item_seg = list(set(item_seg) - set(xiangzhen) - set(qulist))
    if data_type == "直派行业":
        file_name = Infilename
        data = pd.read_excel(file_name, sheet_name='Sheet1')
        jiedao = data["被反映街乡镇"]
        dianwei = data["问题点位"]
        neirong = data["主要内容"]
        for i in range(0, len(data)):
            tiquresult = "不详"
            keywordlist = []
            reslist = []
            chun = dianwei[i]
            chengqu = jiedao[i]
            nr = neirong[i]
            seg_list = jieba.cut(nr)
            seg = list(seg_list)
            for ct in range(0, 3):
                chun = Nstr(chun) - Nstr("北京市")
                chun = Nstr(chun) - Nstr(chengqu)
                for item in xiangzhen:
                    chun = Nstr(chun) - Nstr(item)
            reschun = chun
            for key in mainkey:
                r = chun.find(key)
                if r == -1:
                    pass
                else:
                    tiquresult = chun[0:r + len(key)]
                    for biaozhunhua in item_seg:
                        if get_equal_rate(biaozhunhua, tiquresult) > 0.8:
                            tiquresult = biaozhunhua
                            break
                        else:
                            pass
                    break
            if tiquresult == "不详":
                for keyword in keywordlibiao:
                    for segitem in seg:
                        if keyword in segitem:
                            keywordlist.append(seg.index(segitem))
                countkeyword = len(keywordlist)
                if countkeyword == 0:
                    tiquresult = "不详"
                else:
                    for keyworditem in range(0, countkeyword):
                        res = getpoint(keywordlist[keyworditem])
                        reslist.append(res)
                    panduan = 0
                    for resu in reslist:
                        for keyword in keywordlibiao:
                            if keyword in resu:
                                tiquresult = resu
                                panduan = 1
                                break
                            else:
                                pass
                        if panduan == 1:
                            break
                    if panduan == 0:
                        tiquresult = reschun
                        for biaozhunhua in item_seg:
                            if get_equal_rate(biaozhunhua, tiquresult) > 0.8:
                                tiquresult = biaozhunhua
                                break
                            else:
                                pass
            zhuizhongres.append(tiquresult)
        data['小区村庄'] = zhuizhongres
        savepath = Outdirname + "\\直派行业小区村庄增补.xlsx"
        writer = pd.ExcelWriter(savepath)
        data.to_excel(writer, sheet_name="Sheet1", index=False)
        writer.save()
    elif data_type == "直派水务局":
        file_name = Infilename
        data = pd.read_excel(file_name)
        neirong = data["诉求内容"]
        for i in range(0, len(data)):
            tiquresult = "不详"
            keywordlist = []
            reslist = []
            nr = neirong[i]
            seg_list = jieba.cut(nr)
            seg = list(seg_list)
            for keyword in keywordlibiao:
                for segitem in seg:
                    if keyword in segitem:
                        keywordlist.append(seg.index(segitem))
            countkeyword = len(keywordlist)
            if countkeyword == 0:
                tiquresult = "不详"
            else:
                for keyworditem in range(0, countkeyword):
                    res = getpoint(keywordlist[keyworditem])
                    reslist.append(res)
                panduan = 0
                for resu in reslist:
                    for keyword in keywordlibiao:
                        if keyword in resu:
                            tiquresult = resu
                            panduan = 1
                            break
                        else:
                            pass
                    if panduan == 1:
                        break
                if panduan == 0:
                    tiquresult = "不详"
                    for biaozhunhua in item_seg:
                        if get_equal_rate(biaozhunhua, tiquresult) > 0.8:
                            tiquresult = biaozhunhua
                            break
                        else:
                            pass
            zhuizhongres.append(tiquresult)
        data['小区村庄'] = zhuizhongres
        savepath = Outdirname + "\\直派水务局小区村庄增补.xlsx"
        writer = pd.ExcelWriter(savepath)
        data.to_excel(writer, sheet_name="Sheet1", index=False)
        writer.save()
def Extract_attribute(Infilename=None,Outdirname=None):
    import difflib
    import jieba
    jieba.set_dictionary("./data/dict.txt")
    jieba.initialize()
    def get_equal_rate_1(str1, str2):
        return difflib.SequenceMatcher(None, str1, str2).quick_ratio()

    def stop_word(textlist, stopwordlist):
        outstrlist = []
        for word in textlist:
            if word not in stopwordlist:
                if word != "\t":
                    outstrlist.append(word)
        return outstrlist

    df = pd.read_excel(Infilename, sheet_name='区街道乡镇')
    cklist=['市政供水','自备井','物业','自来水公司','小区','村']
    for i in cklist:
        jieba.add_word(i)

    def stopwordslist():
        stopwords=['$', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '?', '_', '“', '”', '、', '。', '《', '》', '一', '一些', '一何', '一切', '一则', '一方面', '版权', '那', '一样', '一般', '一转眼', '万一', '上', '上下', '下', '不', '偏心', '自己', '不光', '不单', '不只', '不外乎', '不如', '面临', '不尽', '不尽然', '不得', '不怕', '不惟', '不成', '不拘', '不料', '是', '不比', '吃的', '不特', '不独', '不屑一顾', '不至于', '不若', '或者', '不过', '不问', '与', '与其', '与其说', '与否', '本体', '且', '且不说', '且说', '重新', '个', '购买', '临', '为', '为了', '为什么', '为何', '直到', '其中', '为着', '乃', '喜爱', '喜爱于', '么', '之', '一个', '休息', '之类', '乌乎', '乎', '乘', '也', '也好', '也罢', '了', '二来', '于', '于是', '于是乎', '云云', '云尔', '些', '亦', '人', '像这样的', '家人', '什么', '不好看', '今', '附近', '仍', '依旧', '从', '回顾', '从而', '他', '生者', '她们', '以', '以上', '以为', '以便', '向往', '以及', '以故', '以期', '以来', '以至', '以至于', '以致', '我们', '任', '任何', '任凭', '似的', '但', '但凡', '但是', '何', '何以', '何况', '在哪里', '什么时候', '余外', '成为', '你', '你们', '使', '过去', '例如', '依', '相比', '根据', '方便', '俺', '俺们', '以后', '就使', '以后或', '然然', '如果若', '借', '假使', '镜像', '假若', '傥然', '像', '儿', '先不先', '光是', '简体中文', '全部', '兮兮', '关于', '其', '其一', '有', '其二', '其他', '其余', '其他', '即将', '一', '具体说来', '兼之', '内', '再', '再其次', '再则', '再有', '再者', '再者说', '那些', ',', '冲', '况且', '几', '几时', '凡', '凡是', '凭直觉', '综合', '用心', '出来', '分别', '则', '则甚', '别', '别人', '别处', '别是', '别的', '别管', '别说', '到', '相对', '前此', '前者', '加之', '许可', '即', '即令', '即使', '轻微的', '即如', '即或', '即若', '却', '去', '又', '又及', '及', '以及', '及至', '反之', '可能性', '回复', '回复说', '受到', '另', '同时', '另外', '另悉', '只', '只当', '只怕', '只是', '只有', '只消', '给', '只限', '叫', '叮咚', '可', '可以', '而', '可见', '各', '每个人', '各位', '各种', '自己', '同', '同时', '后', '后期', '向', '向使', '向着', '吓', '吗', '否则', '吧', '吧哒', '吱', '呀', '呃', '呕吐', '呗', '呜呜', '呜呼', '呢', '呵', '呵呵', '呸', '呼哧', '咋', '和', '咚', '咦', '那', '咱', '我们', '咳', '哇', '哈', '哈哈', '哉', '哎', '哎呀', '哎哟', '哗', '哟', '哦哦', '哩', '哪', '哪个', '哪些', '哪儿', '天哪', '哪年', '相对', '哪样', '哪边', '哪里', '哼', '哼唷', '唉', '唯有', '啊', '啐', '啥', '啦', '啪达', '啷当', '喂', '喏', '喔唷', '喽', '嗡', '嗡嗡', '嗬', '嗯', '嗳', '嘎', '嘎登', '嘘', '嘛嘛', '嘻', '嘿', '嘿嘿', '因', '因为', '因了', '因此', '因着', '减少', '固然', '在', '在下', '在于', '地', '因为', '报价', '多', '怎么样', '多少', '大', '大家', '她', '女', '好', '如', '如上', '那个', '如下', '怎么办', '如其', '王', '如是', '如果如果', '如此', '如若', '始而', '哪个料', '哪个知', '宁', '宁可', '避免', '宁肯', '它', '的', '对', '对于', '对待', '对方', '对比', '将', '小', '尔', '尔后', '尔尔', '尚且', '就', '就是', '就是了', '就是说', '不良', '她', '尽', '虽然', '……', '岂但', '己', '已', '已矣', '巴', '巴巴', '并', '并且', '不是', '庶乎', '庶几', '开外', '开始', '归归', '归齐', '当', '当地', '当然', '当着', '彼', '彼时', '彼此', '往', '待', '很', '得', '得了', '怎么', '怎么', '怎么办', '怎么样', '怎奈', '怎么样', '花朵', '总的来看', '总', '总的说来', '总说之', '恰好相反', '您的', '惟其', '慢说', '我', '我们', '或', '或则', '求你了', '或曰', '或者', '达到', '所', '所以', '属于', '所幸', '所有', '才', '不', '打', '打从', '把', '抑或', '拿', '按', '按照', '水', '想', '据', '集中', '外', '故', '故此', '故而', '旁人', '无', '无宁', '图', '既', '既往', '既是', '既然', '时候', '是', '是', '是的', '曾', '替', '替补', '最', '有', '今人', '有关', '有及', '有时', '有的', '望', '朝', '想', '本', '我', '本地', '本着', '做作', '来', '来着', '来自', '说', '极了', '果然', '果真', '某', '某人', '部分', '某某', '可知', '欤', '正值', '如', '正巧', '正是', '此', '此地', '这里', '另外', '此时', '来源', '此间', '免宁', '每', '颗粒', '比', '比及', '前辈', '比方', '没奈何', '沿', '沿线', '漫说', '焉', '然则', '然后', '然而', '照', '照着', '犹且', '犹自', '甚且', '什么', '甚或', '甚而', '甚至', '甚至于', '用', '解释', '由', '因为', '由是', '本人', '点', '的', '确实', '话', '直到', '相对说', '省得', '看', '眨眼', '着', '着呢', '矣', '矣乎', '矣哉', '离', '竟而', '第', '等', '候诊室', '等等', '简言之', '管', '类如', '紧接着', '纵', '纵令', '纵使', '纵然', '经', '经过', '结果', '给', '继之', '继后', '继而', '综上所述', '罢了', '者', '而', '还有', '而况', '而后', '而外', '笑声', '混合', '说的', '能', '能不能', '腾', '自', '自个儿', '因为', '自各儿', '自后', '自家', '自己', '自打', '自我', '至', '至于', '来自', '至若', '致', '般的', '若', '若夫', '允', '若果', '若非', '莫吃', '莫如', '莫若', '虽', '虽则', '是不是', '虽说', '被', '要', '要不', '要不是', '要吃', '是否', '可以', '暗示', '不愿如', '让', '从', '论', '设使', '设或', '设若', '诚如', '诚然', '该', '说来', '诸', '诸位', '日常', '谁', '谁人', '谁料', '谁知', '贼死', '赖以', '赶', '起', '起见', '趁', '趁着', '越是', '距', '跟', '较', '较之', '边', '过', '还', '还是', '还有', '还有', '这', '这来', '这个', '如此', '这么些', '这么样', '这么点儿', '这', '这会儿', '这里', '这就是说', '那时', '这样', '这次', '这般', '这篇', '这里', '过', '连', '地图', '逐步', '通过', '遵循', '遵照', '那', '那个', '那么', '那些', '那样', '那些', '那会儿', '那边', '那时', '那', '那般', '反', '那里', '都', '其他人', '开始', '针对', '阿', '除', '除了', '除外', '除开', '梦想', '通过', '随', '迎接', '随时', '随着', '难道说', '非但', '非徒', '非特', '非独', '靠', '顺', '顺着', '首先', '！', '，', '：', '；', '？']
        return stopwords
    list1= []
    list1.extend(0 for _ in range(len(df)))
    list2 = []
    list2.extend(0 for _ in range(len(df)))

    for i in range(len(df)):

        cell_context = df['主要内容'].iloc[i]
        cut_content = jieba.lcut(cell_context)

        stopwords = stopwordslist()  # 创建停用词列表
        fin_cut_content = stop_word(cut_content, stopwords)


        list3 = []
        for s in fin_cut_content:

            if get_equal_rate_1("市政供水", s) > 0.8 or get_equal_rate_1("物业", s) > 0.9 or get_equal_rate_1("自来水公司",
                                                                                        s) > 0.8 or get_equal_rate_1(
                    "小区", s) == 1 or get_equal_rate_1("自来水集团",s) > 0.8:
                list3.append("市政供水")
            if get_equal_rate_1("物业", s) > 0.9 :
                list3.append("物业")
            if get_equal_rate_1("自来水公司", s) > 0.8:
                list3.append("自来水公司")
            if get_equal_rate_1("自来水集团", s) > 0.8:
                list3.append("自来水公司")
            if get_equal_rate_1("自备井", s) == 1 :
                list3.append(s)
            if get_equal_rate_1("井",s) == 1 or get_equal_rate_1("井水",s) > 0.9 or get_equal_rate_1("乡村",s) > 0.9 or get_equal_rate_1("村民",s) > 0.9:
                list3.append("自备井")

        if "自备井" in list3 :

            list1[i] = "自备井"
            list2[i] = "其他产权单位"
        elif"市政供水" in list3 or "物业" in list3 or "自来水公司" in list3:
            list1[i] = "市政供水"
            if "物业" in list3:
                list2[i] = "物业"
            elif "自来水公司" in list3:

                list2[i] = "自来水公司"
            else:
                list2[i] = "其他产权单位"
        else :
            list1[i] = "未知或其他产权单位"
            list2[i] = "未知或其他产权单位"
        #     list2.append("不详")

    df['供水类型'] = list1
    df['负责单位'] = list2
    savepath=Outdirname+"\\供水属性标注.xlsx"
    writer = pd.ExcelWriter(savepath)
    df.to_excel(writer, sheet_name="区街道乡镇", index=False)
    writer.save()
    sys.exit()
def Sum_Extract_time(data_type='直派水务局',Infilename=None,Outdirname=None):
    import pandas as pd
    if data_type=='直派水务局':
        zs=pd.read_excel(Infilename)
        zsdatadf=zs['登记时间']
        zs['年']=zsdatadf.str[2:4]
        zs['月']=zsdatadf.str[5:7]
        zs['日']=zsdatadf.str[8:10]
        zs['时']=zsdatadf.str[11:13]
        old_zs=pd.read_excel('./data/直派水务局（总）.xlsx')
        all_zs=old_zs.append(zs,ignore_index=True)
        savepath1='./data/直派水务局（总）.xlsx'
        writer1 = pd.ExcelWriter(savepath1)   #设置导出的excel文件地址
        writer3 = pd.ExcelWriter(Outdirname+'\\直派水务局（总）.xlsx')
        all_zs.to_excel(writer1,index=False)
        all_zs.to_excel(writer3,index=False)
        writer1.save()
        writer3.save()
    elif data_type=='直派行业':
        hy=pd.read_excel(Infilename)
        hydatadf=hy['创建时间']
        hy['年']=hydatadf.str[2:4]
        hy['月']=hydatadf.str[5:7]
        hy['日']=hydatadf.str[8:10]
        hy['时']=hydatadf.str[11:13]
        old_hy=pd.read_excel('./data/直派行业（总）.xlsx')
        all_hy=old_hy.append(hy,ignore_index=True)
        savepath2='./data/直派行业（总）.xlsx'
        writer2 = pd.ExcelWriter(savepath2)   #设置导出的excel文件地址
        writer4 = pd.ExcelWriter(Outdirname+'\\直派行业（总）.xlsx')
        all_hy.to_excel(writer2,index=False)
        all_hy.to_excel(writer4,index=False)
        writer2.save()
        writer4.save()