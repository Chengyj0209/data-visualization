#构建网络数据清洗，提取三列并删除空
invest_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments.xlsx',usecols = [6,15,27])
l = 0
while type(invest_df.投资时间[l]) != float:
    l += 1
invest_df = invest_df.head(l)

for i in range(l):
    if invest_df.企业名称[i] == '不披露' or invest_df.机构名称[i] == '不公开的投资者':
        invest_df = invest_df.drop([i,i],axis = 0)

invest_df.to_excel('D:/课件/数据可视化/风险投资实验/investments_clear1.xlsx')


#清洗投资退出时间表
ipo_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/MA_IPO.xlsx',usecols = [1,3,6])

for i in range(len(ipo_df)):
    if type(ipo_df.退出机构[i]) == float:
        ipo_df = ipo_df.drop([i,i],axis = 0)

ipo_df.to_excel('D:/课件/数据可视化/风险投资实验/investments_clear2.xlsx')




#清洗1.4表
invest_df = invest_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments.xlsx',usecols = [15,27,32,30,31,33])

for i in range(len(invest_df)):
    if invest_df.二级代码[i] == '--' or invest_df.一级字母代码[i] == '--' or invest_df.机构名称[i] == '不公开的投资者' or invest_df.投资时间[i] == '--' or type(invest_df.投资时间[i]) == float:
        invest_df = invest_df.drop([i, i], axis=0)

invest_df.to_excel('D:/课件/数据可视化/风险投资实验/investments_clear3.xlsx')


#分类4
def check(time1,time2):
    if time2 < '2011-01-01':
        return False
    year = int(time2[0:4])-int(time1[0:4]) - 5
    if year < 0:
        return False
    if year == 0:
        if time2[4:10] < time1[4:10]:
            return False
    return True

in_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments_clear3.xlsx')
all_invester = list(set(in_df.机构名称))
kinds_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/分类/分类4/分类4.xlsx')

for x in all_invester:
    time = list(in_df.投资时间[in_df.机构名称 == x])
    time1 = time[0]
    time2 = time[-1]
    issuccess = check(time1,time2)
    if issuccess:
        issuccess = '是'
    else:
        issuccess = "否"
    data = {"机构名称":x,"是否成功":issuccess}
    kinds_df = kinds_df.append(data,ignore_index=True)

kinds_df.to_excel('D:/课件/数据可视化/风险投资实验/分类/分类4/分类4.xlsx')


##一。分类四
kinds_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/分类/分类4/分类4.xlsx')
invest_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments_clear3.xlsx')
first_1 = {}#某个第一产业一共有多少成功的投资者
first_2 = {}#某个第一产业一共有多少个失败投资者
second_1 = {}
second_2 = {}
tick_1 = {}#第一产业字母对应名称
tick_2 = {}#第二产业数字对应名称
for i in invest_df.itertuples():
    invester = getattr(i,'机构名称')
    first = getattr(i,'一级字母代码')
    second = getattr(i,'二级代码')
    now = list(kinds_df.是否成功[kinds_df.机构名称 == invester])
    tick_1[first] = getattr(i,'一级名称')
    tick_2[second] = getattr(i,'二级名称')
    if list(kinds_df.是否成功[kinds_df.机构名称 == invester])[0] == '是':
        if first not in first_1.keys():
            first_1[first] = 1.0
        else:
            first_1[first] += 1
        if second not in second_1.keys():
            second_1[second] = 1.0
        else:
            second_1[second] += 1
    else:
        if first not in first_2.keys():
            first_2[first] = 1.0
        else:
            first_2[first] += 1
        if second not in second_2.keys():
            second_2[second] = 1.0
        else:
            second_2[second] += 1
#全部一级产业
first_1 = sorted(first_1.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
coor_y = 0.5
draw_y = []#y轴工业名称
number_y = []#需要替换的y坐标
height = 0.85#柱子宽度
for items in range(len(first_1)):
    industry = first_1[items][0]
    quantity_1 = first_1[items][1]
    quantity_2 = first_2[industry]
    if items == 0:
        plt.barh(y = coor_y + height/2,width = quantity_1,color = '#333399',height = height,label = '成功风投机构')
        plt.barh(y = coor_y - height/2,width = quantity_2, color='#FF6600',height = height,label = '失败风投机构')
    else:
        plt.barh(y=coor_y + height / 2, width=quantity_1, color='#333399', height=height)
        plt.barh(y=coor_y - height / 2, width=quantity_2, color='#FF6600', height=height)
    draw_y.append(tick_1[industry])
    number_y.append(coor_y)
    coor_y += 2.3
plt.yticks(number_y,draw_y)
plt.xlabel('机构数量',size = 16)
plt.ylabel('一级名称',size = 16)
plt.title('投资机构投资一级产业',size = 19)
plt.legend()
plt.savefig('my.png')
print('ok')


#成功二级产业
plt.figure(figsize= (1000,2000))
second_1 = sorted(second_1.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
coor_y = 0.5
draw_y = []#y轴工业名称
number_y = []#需要替换的y坐标
height = 1#柱子宽度

for items in range(len(second_1)):
    industry = second_1[items][0]
    quantity_1 = second_1[items][1]
    plt.barh(y = coor_y,width = quantity_1,color = '#333399',height = height)
    #plt.barh(y = coor_y - height/2,width = quantity_2, color='#FF6600',height = height,label = '失败风投机构')
    draw_y.append(tick_2[industry])
    number_y.append(coor_y)
    coor_y += 2
plt.yticks(number_y,draw_y)
plt.xlabel('机构数量',size = 11)
plt.ylabel('二级名称',size = 11)
plt.title('成功投资机构投资二级产业',size = 14)
plt.yticks(fontsize=7)
#plt.legend(frameon=False)
plt.show()

#失败二级产业
fig,ax = plt.subplots()
plt.figure(figsize= (1000,2000))
second_2 = sorted(second_2.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
coor_y = 0.5
draw_y = []#y轴工业名称
number_y = []#需要替换的y坐标
height = 1#柱子宽度

for items in range(len(second_2)):
    industry = second_2[items][0]
    quantity_2 = second_2[items][1]
    #plt.barh(y = coor_y,width = quantity_2,color = '#333399',height = height)
    plt.barh(y = coor_y - height/2,width = quantity_2, color='#FF6600',height = height)
    draw_y.append(tick_2[industry])
    number_y.append(coor_y)
    coor_y += 2
plt.yticks(number_y,draw_y)
plt.xlabel('机构数量',size = 11)
plt.ylabel('二级名称',size = 11)
plt.title('失败投资机构投资二级产业',size = 14)
plt.yticks(fontsize=7)
#plt.legend(frameon=False)
plt.show()



#五类一级产业
kinds_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/分类/分类2/分为五类.xls')
invest_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments_clear3.xlsx')
first_1 = {}#某个第一产业一共有第一类的投资者
first_2 = {}
first_3 = {}
first_4 = {}
first_5 = {}

second_1 = {}
second_2 = {}
second_3 = {}
second_4 = {}
second_5 = {}

tick_1 = {}#第一产业字母对应名称
tick_2 = {}#第二产业数字对应名称
for i in invest_df.itertuples():
    invester = getattr(i,'机构名称')
    first = getattr(i,'一级字母代码')
    second = getattr(i,'二级代码')
    tick_1[first] = getattr(i,'一级名称')
    tick_2[second] = getattr(i,'二级名称')
    if invester not in list(kinds_df.投资机构):
        continue
    if list(kinds_df.分类类别[kinds_df.投资机构 == invester])[0] == 1:
        if first not in first_1.keys():
            first_1[first] = 1.0
        else:
            first_1[first] += 1
        if second not in second_1.keys():
            second_1[second] = 1.0
        else:
            second_1[second] += 1
    elif list(kinds_df.分类类别[kinds_df.投资机构 == invester])[0] == 2:
        if first not in first_2.keys():
            first_2[first] = 1.0
        else:
            first_2[first] += 1
        if second not in second_2.keys():
            second_2[second] = 1.0
        else:
            second_2[second] += 1
    elif list(kinds_df.分类类别[kinds_df.投资机构 == invester])[0] == 3:
        if first not in first_3.keys():
            first_3[first] = 1.0
        else:
            first_3[first] += 1
        if second not in second_3.keys():
            second_3[second] = 1.0
        else:
            second_3[second] += 1
    elif list(kinds_df.分类类别[kinds_df.投资机构 == invester])[0] == 4:
        if first not in first_4.keys():
            first_4[first] = 1.0
        else:
            first_4[first] += 1
        if second not in second_4.keys():
            second_4[second] = 1.0
        else:
            second_4[second] += 1
    else:
        if first not in first_5.keys():
            first_5[first] = 1.0
        else:
            first_5[first] += 1
        if second not in second_5.keys():
            second_5[second] = 1.0
        else:
            second_5[second] += 1

#first_2 = sorted(first_2.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
# first_1 = sorted(first_2.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
first_3 = sorted(first_3.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
#first_4 = sorted(first_4.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
#first_5 = sorted(first_5.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
coor_y = 2.3
draw_y = []#y轴工业名称
number_y = []#需要替换的y坐标
height = 1#柱子宽度
for items in range(len(first_3)):
    industry = first_3[items][0]
    # quantity_1 = first_1[items][1]
    #quantity_2 = first_2[items][1]
    quantity_3 = first_3[items][1]
    #quantity_4 = first_4[items][1]
    #quantity_5 = first_5[items][1]
    #plt.barh(y=coor_y , width=quantity_1, color='navy', height=height)
    #plt.barh(y=coor_y , width=quantity_2, color='navy', height=height)
    plt.barh(y=coor_y, width=quantity_3, color='navy', height=height)
    #plt.barh(y=coor_y , width=quantity_4, color='navy', height=height,)
    #plt.barh(y=coor_y , width=quantity_5, color='navy', height=height)
    draw_y.append(tick_1[industry])
    number_y.append(coor_y)
    coor_y += 1.9
plt.yticks(number_y,draw_y)
plt.yticks(fontsize=10)
plt.xlabel('机构数量',size = 16)
plt.ylabel('一级名称',size = 16)
plt.title('第三类投资机构投资一级产业',size = 19)
plt.legend(loc = 'lower right',frameon=False,fontsize = 10)
plt.show()

#五类二级产业
#second_2 = sorted(second_2.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
#second_1 = sorted(second_1.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
#second_3 = sorted(second_3.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
#second_4 = sorted(second_4.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
second_5 = sorted(second_5.items(),key = lambda kv:(kv[1],kv[0]))#按照value值排序
coor_y = 2.3
draw_y = []#y轴工业名称
number_y = []#需要替换的y坐标
height = 1#柱子宽度
for items in range(len(second_5)):
    industry = second_5[items][0]
    #quantity_1 = second_1[items][1]
    #quantity_2 = second_2[items][1]
    #quantity_3 = second_3[items][1]
    #quantity_4 = second_4[items][1]
    quantity_5 = second_5[items][1]
    #plt.barh(y=coor_y , width=quantity_1, color='navy', height=height)
    #plt.barh(y=coor_y , width=quantity_2, color='navy', height=height)
    #plt.barh(y=coor_y, width=quantity_3, color='navy', height=height)
    #plt.barh(y=coor_y , width=quantity_4, color='navy', height=height,)
    plt.barh(y=coor_y , width=quantity_5, color='navy', height=height)
    draw_y.append(tick_2[industry])
    number_y.append(coor_y)
    coor_y += 1.9
plt.yticks(number_y,draw_y)
plt.yticks(fontsize=7)
plt.xlabel('机构数量',size = 11)
plt.ylabel('二级名称',size = 11)
plt.title('第五类投资机构投资二级产业',size = 14)
#plt.legend(loc = 'lower right',frameon=False,fontsize = 10)
plt.show()



#计算vc活跃指数建表
file = op.load_workbook(r'D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')
sheet = file['Sheet1']
invest_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments_clear3.xlsx')

year = []#存入年份
for i in invest_df.itertuples():
    time = getattr(i, '投资时间')[0:4]
    year.append(time)
year = sorted(list(set(year)))

for i in range(len(year)):
    sheet.cell(1,i+2,year[i])

invester = []#机构
for i in invest_df.itertuples():
    name = getattr(i, '机构名称')
    invester.append(name)
invester = sorted(list(set(invester)))

for i in range(len(invester)):
    sheet.cell(i+2,1,invester[i])

file.save('D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')

#活跃指数表
active_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')
invest_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments_clear3.xlsx')

invester = list(set(invest_df.机构名称))
year_num = {}
year = list(active_df.columns.values)
year.remove('机构名称')

for i in invest_df.itertuples():   #统计每年的投资总次数
    now = getattr(i, '投资时间')[0:4]
    if now not in year_num.keys():
        year_num[now] = 0
    year_num[now] += 1

file = op.load_workbook(r'D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')
sheet = file['Sheet1']
for j in range(len(active_df)):
    name = active_df.机构名称[j]
    now_df = invest_df[invest_df.机构名称 == name]
    each_year = {}
    for x in year:
        each_year[x] = 0
    for i in now_df.itertuples():
        now = getattr(i, '投资时间')[0:4]
        if now not in each_year.keys():
            each_year[now] = 0
        each_year[now] += 1

    for time,num in each_year.items():
        percentage = num / year_num[time]
        percentage = format(percentage,'.6f')
        sheet.cell(j+2, int(time) - 1991 + 2, percentage)
        #print(active_df[active_df.机构名称 == name].time)

file.save('D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')



#近四年vc活跃指数变化图
active_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')
file = op.load_workbook(r'D:/课件/数据可视化/风险投资实验/分类/分类1/风险投资领袖.xlsx')
sheet = file['Sheet1']

elite = []#大佬名单
for i in range(1,sheet.max_row+1):
    elite.append(sheet.cell(row = i,column = 1).value)

follow = list(active_df.机构名称)#小弟名单
for x in follow:
    if x in elite:
        follow.remove(x)

file = op.load_workbook(r'D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')
sheet = file['Sheet1']

year = list(active_df.columns.values)#横坐标
year.remove('机构名称')
index_x = 2
ticks_x = []
show_x = []
tick_y = []
ok1 = True
ok2 = True
for i in range(len(year)-4,len(year)):
    now = int(year[i])
    for j in range(len(active_df)):
        invester = active_df.机构名称[j]
        offset = np.random.rand() * 3.5
        if j % 2 == 0:
            offset *= -1
        draw_y = float(sheet.cell(row = j+2,column = i+2).value)
        if draw_y == 0:
            continue
        if invester in elite:
            if ok1:
                plt.scatter(x = index_x + offset, y = draw_y,c = 'red',s = 12,label = 'elite')
                ok1 = False
            else:
                plt.scatter(x=index_x + offset, y=draw_y, c='red', s=12)
        else:
            if ok2:
                plt.scatter(x=index_x + offset, y=draw_y, c='blue',s = 12,label = 'follow')
                ok2 = False
            else:
                plt.scatter(x=index_x + offset, y=draw_y, c='blue', s=12)

    ticks_x.append(index_x)
    show_x.append(now)
    index_x += 8
    print(year[i])

plt.xticks(ticks_x,show_x,fontsize = 12)
plt.yticks([0.00,0.005,0.01,0.015,0.02,0.03,0.04,0.05,0.06],[0.00,0.005,0.01,0.015,0.02,0.03,0.04,0.05,0.06],fontsize = 12)
plt.title('近四年VC活跃指数变化',fontsize = 16)
plt.legend(loc = 'upper right',frameon=False,fontsize = 12)
plt.ylabel('活跃指数',fontsize = 10)
plt.show()



#历年活跃指数
active_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')
file = op.load_workbook(r'D:/课件/数据可视化/风险投资实验/分类/分类1/风险投资领袖.xlsx')
sheet = file['Sheet1']

elite = []#大佬名单
for i in range(1,sheet.max_row+1):
    elite.append(sheet.cell(row = i,column = 1).value)

follow = list(active_df.机构名称)#小弟名单
for x in follow:
    if x in elite:
        follow.remove(x)

file = op.load_workbook(r'D:/课件/数据可视化/风险投资实验/活跃指数.xlsx')
sheet = file['Sheet1']

year = list(active_df.columns.values)#横坐标
year.remove('机构名称')
index_x = 2
ticks_x = []
show_x = []
tick_y = []
width = 1

fig,ax = plt.subplots()

for i in range(0,len(year)):
    now = int(year[i])
    num_1 = {}
    num_2 = {}
    for j in range(len(active_df)):
        invester = active_df.机构名称[j]

        draw_y = float(sheet.cell(row = j+2,column = i+2).value)
        draw_y = format(draw_y,'.4f')
        if invester in elite:
            if draw_y not in num_1.keys():
                num_1[draw_y] = 0
            num_1[draw_y] += 1
        else:
            if draw_y not in num_2.keys():
                num_2[draw_y] = 0
            num_2[draw_y] += 1

    num_1 = sorted(num_1.items(),key = lambda kv:(kv[1],kv[0]))
    num_2 = sorted(num_2.items(),key = lambda kv:(kv[1],kv[0]))
    h_1 = float(num_1[len(num_1)-1][0])
    h_2 = float(num_2[len(num_2) - 1][0])
    if h_1 == 0 and len(num_1) > 1:
        h_1 = float(num_1[len(num_1)-2][0])
    if h_2 == 0 and len(num_2) > 1:
        h_2 = float(num_2[len(num_2)- 2][0])

    if i == 0:
        plt.bar(x=index_x - width / 2, height=h_1, width=width,color = 'brown',label = 'elite')
        plt.bar(x=index_x + width / 2, height=h_2, width=width,color = 'mediumblue',label = 'follow')
    else:
        plt.bar(x=index_x - width / 2, height=h_1, width=width, color='brown')
        plt.bar(x=index_x + width / 2, height=h_2, width=width, color='mediumblue')

    ticks_x.append(index_x)
    show_x.append(now)
    index_x += 4
    print(year[i])

plt.xticks(ticks_x,show_x,fontsize = 12)
#plt.yticks([0.00,0.005,0.01,0.015,0.02,0.03,0.04,0.05,0.06,0.07,0.08,0.09,0.1],[0.00,0.005,0.01,0.015,0.02,0.03,0.04,0.05,0.06,0.07,0.08,0.09,0.1],fontsize = 12)
plt.title('历年两类VC活跃指数变化',fontsize = 16)
plt.legend(loc = 'upper right',frameon=False,fontsize = 12)
plt.ylabel('活跃指数',fontsize = 12)
plt.show()

#五类对于第一产业投资的转变
invest_df = pd.read_excel('D:/课件/数据可视化/风险投资实验/investments_clear3.xlsx')
file = op.load_workbook(r'D:/课件/数据可视化/风险投资实验/分类/分类2/分为五类.xlsx')
sheet = file['Sheet1']

excellent = []
good = []
fair = []
bad = []
failure = []

for i in range(2,39):
    excellent.append(sheet.cell(row = i,column = 1).value)

for i in range(39,159):
    good.append(sheet.cell(row = i,column = 1).value)

for i in range(159,317):
    fair.append(sheet.cell(row = i,column = 1).value)

for i in range(317,505):
    bad.append(sheet.cell(row = i,column = 1).value)

for i in range(505,6221):
    failure.append(sheet.cell(row = i,column = 1).value)

year = list(range(2005,2016))
industry = ['采矿业', '租赁和商务服务业 ', '金融业', '教育', '农/林/牧/渔业', '住宿和餐饮业', '科学研究和技术服务业', '建筑业', '卫生和社会工作', '交通运输、仓储和邮政业 ', '农、林、牧、渔业', '房地产业', '批发和零售业', '制造业', '居民服务、修理和其他服务业', '文化、体育和娱乐业', '信息传输、软件和信息技术服务业']
address_x = 2.5
address_y = 1.5
year_x = {}
industry_y = {}
draw_x = []
draw_y =[]

for i in year:
    year_x[i] = address_x
    draw_x.append(address_x)
    address_x += 2.5

for i in industry:
    industry_y[i] = address_y
    draw_y.append(address_y)
    plt.plot([0, address_x], [address_y, address_y],c = 'black',linestyle = ':',alpha = 0.5)
    address_y += 2

for i in invest_df.itertuples():
    invester = getattr(i,'机构名称')
    now = int(getattr(i,'投资时间')[0:4])
    indust = getattr(i,'一级名称')

    if invester not in failure:
        continue
    if now not in year:
        continue

    offset_x = np.random.rand()
    sign_x = offset_x
    while sign_x < 1:
        sign_x *= 10
    sign_x = int(sign_x)
    if sign_x % 2 == 0:
        offset_x *= -1

    offset_y = np.random.rand()
    sign_y = offset_y
    while sign_y < 1:
        sign_y *= 10
    sign_y = int(sign_y)
    if sign_y % 2 == 0:
        offset_y *= -1
    offset_y *= 0.5
    if ((industry_y[indust] - 1.5)/2)%2 == 0:
        plt.scatter(x = year_x[now] + offset_x,y = industry_y[indust] + offset_y,c='b',s=6,marker = 'v')
    else:
        plt.scatter(x=year_x[now] + offset_x, y=industry_y[indust] + offset_y, c='r', s=6, marker='^')

plt.title('failure对一级产业投资的转变',fontsize = 16)
plt.xticks(draw_x,year,fontsize = 12)
plt.yticks(draw_y,industry,fontsize = 12)



