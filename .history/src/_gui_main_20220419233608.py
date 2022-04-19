from gooey import Gooey, GooeyParser
import matplotlib
import pandas.plotting._matplotlib
import matplotlib.pyplot as plt
import matplotlib.backends.backend_tkagg
matplotlib.use('tkagg')
import 
@Gooey(
richtext_controls=True,                 # 打开终端对颜色支持
program_name="市水务局涉水诉求工单数据分析",        # 程序名称
encoding="utf-8",                  # 设置编码格式，打包的时候遇到问题
progress_regex=r"^progress: (\d+)%$",   # 正则，用于模式化运行时进度信息
default_size=(800,600)
)
def start():
    parser = GooeyParser(description='')

    subs = parser.add_subparsers(help='commands', dest='command')

    Fir_parser = subs.add_parser('属性增补')
    
    Fir_parser.add_argument (
            "data_type", 
    metavar='数据源格式',
    help=   "请选择数据源格式",
    choices=['直派水务局','直派行业'], 
    default='直派水务局')
    Fir_parser.add_argument (
            "device_type",
    metavar='增补属性',
    help=   "请选择需要增补的数据",
    choices=['河湖','小区（村）','供水属性标注','汇总数据以及时间分列'],
    default='河湖')

    Fir_parser.add_argument (
            "FirInFilename",
    metavar='选择文件',
    widget='MultiFileChooser',
    default=None,
    help='请选择源数据')
    
    Fir_parser.add_argument(
            "FirOutDirname", 
    metavar='输出文件目录',
    widget='DirChooser',
    default=None,
    help='请选择输出文件目录')
    

    Sec_parser = subs.add_parser('智能分类和原因提取')
    Sec_parser.add_argument(
            "classify_reason", 
    metavar='智能分类和原因提取',
    choices=['诉求智能分类（直派水务局）','诉求原因提取（直派行业）'], 
    default=None)
    Sec_parser.add_argument (
            "SecInFilename",
    metavar='选择输入文件',
    widget='MultiFileChooser',
    default=None,
    help='请选择源数据')
    Sec_parser.add_argument (
            "SecOutDirname",
    metavar='保存文件',
    widget='DirChooser',
    default=None,
    help='请选择保存文件目录')
    
    third_parser=subs.add_parser('统计分析')
    third_parser.add_argument (
            "data_type", 
    metavar='数据源格式',
    help=   "请选择数据源格式",
    choices=['直派水务局','直派行业'], 
    default='直派水务局')
    third_parser.add_argument(
        "YearMonth",
    metavar='输入当前年月',
    widget='TextField',
    default='202101',
    help='')
    third_parser.add_argument (
            "Count",
    metavar='统计',
    help=   "请选择需要统计的数据",
    choices=['单位解决类型','诉求变化趋势','水利工程运行维护','各区街道集团','只绘图'],
    default='单位解决类型')
    third_parser.add_argument (
            "Draw",
    metavar='绘图',
    help=   "请选择需要绘制的数据",
    choices=['水利工程运行维护变化趋势-折线图','水环境维护变化趋势-折线图','雨水与海绵城市变化趋势-折线图','饼状图','柱状图（直派水务局一级诉求）','柱状图（直派行业三级问题）','地理分布图（区）','只统计'],
    default='水利工程运行维护变化趋势-折线图')
    third_parser.add_argument (
            "ThirdInFilename",
    metavar='输入文件',
    widget='MultiFileChooser',
    default=None,
    help='请选择源数据')
    third_parser.add_argument (
            "ThirdOutDirname",
    metavar='保存文件',
    widget='DirChooser',
    default=None,
    help='请选择保存文件目录')
    fouth_parser=subs.add_parser('相似性分析、趋势预测和突变点标识')
    fouth_parser.add_argument (
            "demand", 
    metavar='需求',
    help=   "请选择需求",
    choices=['相似性分析','趋势预测','突变点标识'], 
    default='相似性分析')
    fouth_parser.add_argument(
        "YearMonth",
    metavar='输入当前年月',
    widget='TextField',
    default='202001',
    help='')
    fouth_parser.add_argument(
        "FouthInFilename",
    metavar='输入Excel文件',
    widget='MultiFileChooser',
    default=None,
    help='请选择输入文件（汇总excel）')
    
    fouth_parser.add_argument(
            "FouthOutDirname",
    metavar='保存文件',
    widget='DirChooser',
    default=None,
    help='请选择保存文件目录')

    fifth_parser = subs.add_parser('月报')
    fifth_parser.add_argument(
        "Yearmonth",
    metavar='输入当前年月',
    widget='TextField',
    default='202001',
    help='')
    fifth_parser.add_argument(
        "FifthInFilename1",
    metavar='输入文件',
    widget='MultiFileChooser',
    default=None,
    help='请输入上月直派水务局数据')
    fifth_parser.add_argument(
        "FifthInFilename2",
    metavar='输入文件',
    widget='MultiFileChooser',
    default=None,
    help='请输入该月直派水务局数据')
    fifth_parser.add_argument(
        "FifthInFilename3",
    metavar='输入文件',
    widget='MultiFileChooser',
    default=None,
    help='请输入上月直派行业数据')
    fifth_parser.add_argument(
        "FifthInFilename4",
    metavar='输入文件',
    widget='MultiFileChooser',
    default=None,
    help='请输入该月直派行业数据')
    fifth_parser.add_argument(
            "FifthOutDirname",
    metavar='保存文件',
    widget='DirChooser',
    default=None,
    help='请选择保存文件目录')

    args = parser.parse_args()
    print(args,flush=True)    # 坑点：flush=True在打包的时候会用到
   	# 将界面收集的参数进行处理
    return args
if __name__ == '__main__':

    args=start()#InFilename='输入单文件',InDirname='输入文件目录',OutDirname='输出文件目录'
    if args.command=='属性增补':
        if args.device_type=='河湖':
            Extract_river(args.data_type,args.FirInFilename,args.FirOutDirname)
        elif args.device_type=='小区（村）':
            Extract_village(args.data_type,args.FirInFilename,args.FirOutDirname)
        elif args.device_type=='供水属性标注':
            Extract_attribute(args.FirInFilename,args.FirOutDirname)
        elif args.device_type=='汇总数据以及时间分列':
            Sum_Extract_time(args.data_type,args.FirInFilename,args.FirOutDirname)

    
    elif args.command=='智能分类和原因提取':
        if args.classify_reason=='直派水务局':
            Classification(args.SecInFilename,args.SecOutDirname)

        elif args.classify_reason=='直派行业':
            Extract_reason(args.SecInFilename,args.SecOutDirname)

    elif args.command=='统计分析':
        if args.Draw=='只统计':
            Count(args.data_type,args.YearMonth,args.Count,args.ThirdInFilename,args.ThirdOutDirname)

        elif args.Count=='只绘图':
            Draw(args.data_type,args.Draw,args.ThirdInFilename,args.ThirdOutDirname)

    elif args.command=='相似性分析、趋势预测和突变点标识':
        if args.demand=='相似性分析':
            similar_ayalyse(args.YearMonth,args.FouthInFilename,args.FouthOutDirname)
        else:
            predict_plitti(args.demand,args.YearMonth,args.FouthInFilename,args.FouthOutDirname)

    elif args.command=='月报':
        month_report(args.Yearmonth,args.FifthInFilename1,args.FifthInFilename2,args.FifthInFilename3,args.FifthInFilename4,args.FifthOutDirname)

