import pandas,numpy,matplotlib
from scipy import stats
#from statsmodels.stats.multicomp import pairwise_tukeyhsd
#from statsmodels.stats.anova import anova_lm
# 方差分析依赖的库，现不需要
from statsmodels.formula.api import ols
from matplotlib import pyplot,ticker
import seaborn


base = pandas.read_excel('xxxxx.xlsx')
# 原始数据所在的文档


'''
# 检验方差齐，弃用
def dopps_stest(cate,var):
    cut = base[[cate,var]].dropna()
    values = base[cate].unique()
    s_list = []
    for value in values:
        s_list.append(cut[cut[cate]==value][var])
    ts = stats.levene(*s_list)
    return ts

# 手动生成哑变量的回归，弃用
def dopps_Regols_manual(var):
    cut = base[['Region',var]].dropna()
    cut[['isShanghai','isGuangzhou']] = 0
    cut.loc[cut['Region']=='Shanghai','isShanghai'] = 1
    cut.loc[cut['Region']=='Guangzhou','isGuangzhou'] = 1
    formu = var+"~isShanghai+isGuangzhou"
    model = ols(formula=formu, data=cut).fit()
    return model.summary()
    
# 检验分组后各组的正态性，用不上
def dopps_ntest(cate,var):
    cut = base[[cate,var]].dropna()
    values = cut[cate].unique()
    n_list = []
    for value in values:
        ts = stats.shapiro(cut[cut[cate]==value][var])
        print(ts)
    return None
'''    

# 检验正态性，简单封装
def dopps_ntest(var, data=base):
    cut = data[var].dropna()
    return stats.shapiro(cut)
    #P值在pvalue属性中

# 定量数据：回归
def dopps_ols(x, y, summary=True, data=base):
    cut = data[[x,y]].dropna()
    formu = f"{y}~C({x})"
    model = ols(formula=formu, data=cut).fit()
    if not summary:
        return model
    #P值在f_pvalue属性中
    return model.summary()

# 分类数据：卡方检验
def dopps_chi2(var1,var2, data=base):
    cut = data[[var1,var2]].dropna()
    table = pandas.crosstab(cut[var1],cut[var2])
    return stats.chi2_contingency(table)
    #P值在pvalue属性中

# 列定量数据的表格，正态数据->均值±标准差； 非正态数据->中位数、上下四分位数
def table_quant(var, cate=['Region','Gender'], data=base, transform_age=True):
    if not isinstance(cate, list):
        cate = [cate]
    cut = data[[var,*cate]].dropna()
    if 'Age' in cate and transform_age:
        cut['Age'] = pandas.cut(cut['Age'], bins=[0,45,55,65,75,95], right=False, labels=['<45','45-54','55-64','65-74','≥75'])
    if stats.shapiro(cut[var]).pvalue>0.05:
        mean = round(cut[var].mean(), 1)
        std = round(cut[var].std(), 1)
        all_stat = f"{mean}±{std}"
        table = pandas.DataFrame({'Variables':[var], 'All':[all_stat]})
        for group in cate:
            series_mean = round(cut.groupby(group, observed=False)[var].mean(), 1)
            series_std = round(cut.groupby(group, observed=False)[var].std(), 1)
            group_names = cut[group].value_counts().sort_index().index
            #value_counts()配合sort_index()完成的排序可以把<和>两个类别放在纯数字范围类别两侧，sorted()配合unique()做不到。
            for subgroup in group_names:
                curr_mean = series_mean[subgroup]
                curr_std = series_std[subgroup]
                table[subgroup] = f"{curr_mean}±{curr_std}"
            pvalue = dopps_ols(group, var, summary=False, data=cut).f_pvalue
            table[f"P_{group}"] = round(pvalue, 4) if pvalue>0.0001 else '<0.0001'
    else:
        median = round(cut[var].quantile(), 1)
        q1 = round(cut[var].quantile(0.25), 1)
        q3 = round(cut[var].quantile(0.75), 1)
        all_stat = f"{median} ({q1}, {q3})"
        table = pandas.DataFrame({'Variables':[var], 'All':[all_stat]})
        for group in cate:
            series_median = round(cut.groupby(group, observed=False)[var].quantile(), 1)
            series_q1 = round(cut.groupby(group, observed=False)[var].quantile(0.25), 1)
            series_q3 = round(cut.groupby(group, observed=False)[var].quantile(0.75), 1)
            group_names = cut[group].value_counts().sort_index().index
            for subgroup in group_names:
                curr_median = series_median[subgroup]
                curr_q1 = series_q1[subgroup]
                curr_q3 = series_q3[subgroup]
                table[subgroup] = f"{curr_median} ({curr_q1}, {curr_q3})"
            pvalue = dopps_ols(group, var, summary=False, data=cut).f_pvalue
            table[f"P_{group}"] = round(pvalue, 4) if pvalue>0.0001 else '<0.0001'
    return table

# 对定量数据绘图，提琴图
def plot_quant(var, cate=['Region','Gender'], data=base, transform_age=True):
    if not isinstance(cate, list):
        cate = [cate]
    cut = data[[var,*cate]].dropna()
    if 'Age' in cate and transform_age:
        cut['Age'] = pandas.cut(cut['Age'], bins=[0,45,55,65,75,95], right=False, labels=['<45','45-54','55-64','65-74','≥75'])
    plot_ratio = [1]
    chart_width = 12.8/11
    chart_num = 1
    for group in cate:
        plot_ratio.append(len(cut[group].unique()))
        chart_num += len(cut[group].unique())
    color_map = ['Greens', 'YlOrBr', 'Reds', 'Greys', 'Purples', 'Blues', 
                 'Oranges', 'YlOrRd', 'OrRd', 'PuRd', 'RdPu', 'BuPu', 
                 'GnBu', 'PuBu', 'YlGnBu', 'PuBuGn', 'BuGn', 'YlGn']
    if len(cate)>len(color_map):
        print('无法对过多的分组变量上不同的色彩')
    fig,axes = pyplot.subplots(
        1, 
        1+len(cate), 
        gridspec_kw={"width_ratios":plot_ratio}, 
        figsize=(chart_width*chart_num,4.8)
    )
    seaborn.violinplot(
        data=cut[var], 
        cut=0, 
        ax=axes[0]
    )
    axes[0].set_xlabel('All')
    axes[0].spines[['top','right']].set_visible(False)
    for index,ax in enumerate(axes[1:]):
        seaborn.violinplot(
            x=cut[cate[index]], 
            y=cut[var], 
            cut=0, 
            palette=color_map[index], 
            hue=cut[cate[index]], 
            ax=axes[index+1]
        )
        #seaborn.despine()
        ax.yaxis.set_visible(False)
        ax.spines[['left','right','top']].set_visible(False)
    pyplot.show()

# 列定序数据的表格
def table_class(var, cate=['Region','Gender'], data=base, show_percent=True, transform_age=True):
    if not isinstance(cate, list):
        cate = [cate]
    cut = data[[var,*cate]].dropna()
    if 'Age' in cate and transform_age:
        cut['Age'] = pandas.cut(cut['Age'], bins=[0,45,55,65,75,95], right=False, labels=['<45','45-54','55-64','65-74','≥75'])
    percent = cut[var].value_counts(normalize=True).sort_index()
    if show_percent:
        percent = round(percent*100, 1)
    table = pandas.DataFrame({'Variables':[var, *percent.index]})
    table['All'] = ["", *percent]
    if show_percent:
        table.loc[1:,'All'] = table.loc[1:,'All'].map(lambda per: f"{per}%")
    for group in cate:
        series_percent = cut.groupby(group, observed=False)[var].value_counts(normalize=True).sort_index()
        if show_percent:
            series_percent = round(series_percent*100, 1)
        subgroups = list(dict.fromkeys([t[0] for t in series_percent.index]))
        #转换成字典再转回来，以清除重复项同时保持顺序
        for subgroup in subgroups:
            #table[subgroup] = ["", *series_percent[subgroup]]
            #这样有时会出现某些类别占比为0，不计入value_count()的输出，无法将表中对应列填满。
            #报错：Length of values does not match length of index
            series_copy = list(series_percent[subgroup].index)
            for i,indexname in enumerate(percent.index):
                if indexname not in series_copy:
                    series_copy.insert(i,0)
                else:
                    series_copy[i] = series_percent[subgroup][indexname]
            table.loc[:,subgroup] = ["", *series_copy]
            if show_percent:
                table.loc[1:,subgroup] = table.loc[1:,subgroup].map(lambda per: f"{per}%")
        pvalue = dopps_chi2(group, var, data=cut).pvalue
        pvalue = round(pvalue, 4) if pvalue>0.0001 else '<0.0001'
        table[f"P_{group}"] = ''
        table.loc[0,f"P_{group}"] = pvalue
    return table

# 对定序数据绘图，百分比堆积柱状图
def plot_class(var, cate=['Region','Gender'], data=base, transform_age=True, colormap='Set3'):
    if not isinstance(cate, list):
        cate = [cate]
    cut = data[[var,*cate]].dropna()
    if 'Age' in cate and transform_age:
        cut['Age'] = pandas.cut(cut['Age'], bins=[0,45,55,65,75,95], right=False, labels=['<45','45-54','55-64','65-74','≥75'])
    table = table_class(var, cate=cate, data=cut, show_percent=False, transform_age=False)
    plot_ratio = [1]
    chart_width = 12.8/11
    chart_num = 1
    for group in cate:
        plot_ratio.append(len(cut[group].unique()))
        chart_num += len(cut[group].unique())
    cmap = matplotlib.colormaps[colormap].colors
    fig,axes = pyplot.subplots(
        1, 
        1+len(cate), 
        gridspec_kw={"width_ratios":plot_ratio}, 
        figsize=(chart_width*chart_num,4.8)
    )
    axes[0].yaxis.set_major_formatter(ticker.PercentFormatter(xmax=100))
    table.loc[1:,'All'] *= 100
    table.loc[1:,'All'] = table.loc[1:,'All'].map(lambda x: round(x, 1))
    bottom_0 = 0
    for C in range(1,len(table)):
        bar = axes[0].bar(
            x='All', 
            height=table.loc[C,'All'], 
            bottom=bottom_0, 
            color=cmap[C-1]
        )
        axes[0].bar_label(bar, fmt='%g%%', label_type='center')
        bottom_0 += table.loc[C,'All']
    axes[0].spines[['top','right']].set_visible(False)
    for i,group in enumerate(cate, start=1):
        series_index = cut.groupby(group, observed=False)[var].value_counts().sort_index()
        subgroups = list(dict.fromkeys([t[0] for t in series_index.index]))
        table.loc[1:,subgroups] *= 100
        table.loc[1:,subgroups] = table.loc[1:,subgroups].map(lambda x: round(x, 1))
        axes[i].yaxis.set_major_formatter(ticker.PercentFormatter(xmax=100))
        bottom_i = numpy.zeros(len(subgroups))
        for C in range(1,len(table)):
            bar = axes[i].bar(
                x=subgroups, 
                height=table.loc[C,subgroups], 
                bottom=bottom_i, 
                color=cmap[C-1], 
                label=table.iloc[C,0]
            )
            b_label = axes[i].bar_label(bar, fmt='%g%%', label_type='center')
            for bl in b_label:
                if bl.get_text()=='0%':
                    bl.set_visible(False)
            bottom_i += table.loc[C,subgroups]
        axes[i].yaxis.set_visible(False)
        axes[i].spines[['left','right','top']].set_visible(False)
    pyplot.legend(bbox_to_anchor=(1,1))
    pyplot.show()
    
# 输出整合的表格和图
# multi_var是嵌套的列表，形如[[],[],[]]。内层列表包含相同种类的数据。
# seq是形如[1,0,0]或[True,False,False]的列表，指示multi_var中每个内层列表是何种数据。
# 1和True代表定序数据，0和False代表定量数据。
def multichart(multi_var, seq, cate=['Region','Gender'], data=base, ifplot=True, transform_age=True):
    if not isinstance(multi_var, list):
        multi_var = [[multi_var]]
    if not isinstance(multi_var[0], list):
        multi_var = [multi_var]
    if not isinstance(seq, list):
        seq = [seq]
    if not isinstance(cate, list):
        cate = [cate]
    plain_var = []
    for var in multi_var:
        plain_var += var
    cut = data[[*plain_var,*cate]].dropna()
    if 'Age' in cate and transform_age:
        cut['Age'] = pandas.cut(cut['Age'], bins=[0,45,55,65,75,95], right=False, labels=['<45','45-54','55-64','65-74','≥75'])
    chart_list=[]
    for index,var in enumerate(multi_var):
        subchart_list = []
        if seq[index]:
            for single_var in var:
                subchart_list.append(table_class(single_var, cate=cate, data=cut, transform_age=False))
                if ifplot:
                    plot_class(single_var, cate=cate, data=cut, transform_age=False)
        else:
            for single_var in var:
                subchart_list.append(table_quant(single_var, cate=cate, data=cut, transform_age=False))
                if ifplot:
                    plot_quant(single_var, cate=cate, data=cut, transform_age=False)
        subchart = pandas.concat(subchart_list, ignore_index=True)
        chart_list.append(subchart)
    chart = pandas.concat(chart_list, ignore_index=True)
    return chart
