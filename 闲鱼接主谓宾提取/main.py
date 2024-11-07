import pandas as pd

excel_file_path = 'D:\副本数据标注解析版（2014-2023）.xlsx'
sheet_name = 'Sheet1'
year = []
name = []
w_num = []
num = []
w_type = []
type=[]
ju = []
guanxi = []
df = pd.read_excel(excel_file_path, sheet_name=sheet_name, dtype=str)
# not_name2 = ['research-problem','dataset','code','baseline','Baseline','result','ablation-analysis']
not_name2 = ['code','ablation-analysis']

for index, row in df.iterrows():
    num.append(row['序号'])
    year.append(row['年份'])
    w_num.append(row['文章编号'])
    name.append(row['篇名'])
    w_type.append(row['文章类型'])
    type.append(row['贡献类型'])
    ju.append(row['贡献句'])
    guanxi.append(row['实体及关系'])

year_ = []
name_ = []
w_num_ = []
num_ = []
w_type_ = []
type_ =[]
ju_ = []
guanxi_ = []
for i in range(len(guanxi)):
    if type[i] not in not_name2:
        year_.append(year[i])
        name_.append(name[i])
        w_num_.append(w_num[i])
        num_.append(num[i])
        w_type_.append(w_type[i])
        type_.append(type[i])
        ju_.append(ju[i])
        guanxi_.append(guanxi[i])

l1 = []
a1 = []
a2 = []
a3 = []
a4 = []
a5 = []
a6 = []
a7 = []

def split_string(input_string):
    sentences = input_string.split(';')
    l1 = []  # 主语
    l2 = []  # 谓语
    l3 = []  # 宾语
    l4 = []  # 补语谓语（针对主语）
    l5 = []  # 补语宾语（针对主语）
    l6 = []  # 补语谓语（针对宾语）
    l7 = []  # 补语宾语（针对宾语）
    result = []
    for sentence in sentences:
        parts = sentence.split(':')
        if len(parts) > 3:
            l1.append(parts[0])
            l2.append(parts[1])
            if '@' in parts[2] :
                if '@' not in parts[3] and '$' not in parts[3]:
                    parts_mid = str(parts[2])
                    before_at = parts_mid.split('@', 1)[0]
                    after_at = parts_mid.split('@', 1)[1]
                    l3.append(before_at)
                    l4.append(after_at)
                    l5.append(parts[3])
                    l6.append(None)
                    l7.append(None)
                elif '@' in parts[3] :
                    parts_mid = str(parts[2])
                    before_at_mid1 = parts_mid.split('@', 1)[0]
                    after_at_mid1 = parts_mid.split('@', 1)[1]
                    parts_mid = str(parts[3])
                    before_at_mid2 = parts_mid.split('@', 1)[0]
                    after_at_mid2 = parts_mid.split('@', 1)[1]
                    before_at = before_at_mid1 + ','+before_at_mid2
                    after_at = after_at_mid1
                    l3.append(before_at)
                    l4.append(after_at)
                    l5.append(after_at_mid2+','+parts[4])
                    l6.append(None)
                    l7.append(None)
                elif '$' in parts[3]:
                    parts_mid = str(parts[2])
                    before_at_mid1 = parts_mid.split('@', 1)[0]
                    after_at_mid1 = parts_mid.split('@', 1)[1]
                    parts_mid = str(parts[3])
                    before_dollar_mid2 = parts_mid.split('$', 1)[0]
                    after_dollar_mid2 = parts_mid.split('$', 1)[1]
                    l3.append(before_at_mid1)
                    l4.append(after_at_mid1)
                    l5.append(before_dollar_mid2)
                    l6.append(after_dollar_mid2)
                    l7.append(parts[4])
            elif '$' in parts[2]:
                if '$' not in parts[3] and '@' not in parts[3]:
                    before_dollar = parts[2].split('$', 1)[0]
                    after_dollar = parts[2].split('$', 1)[1]
                    l3.append(before_dollar)
                    l4.append(None)
                    l5.append(None)
                    l6.append(after_dollar)
                    l7.append(parts[3])
                elif '$' in parts[3]:
                    parts_mid = str(parts[2])
                    before_dollar_mid1 = parts_mid.split('$', 1)[0]
                    after_dollar_mid1 = parts_mid.split('$', 1)[1]
                    parts_mid = str(parts[3])
                    before_dollar_mid2 = parts_mid.split('$', 1)[0]
                    after_dollar_mid2 = parts_mid.split('$', 1)[1]
                    before_dollar = before_dollar_mid1
                    after_dollar = after_dollar_mid1 +','+ after_dollar_mid2
                    l3.append(before_dollar)
                    l4.append(None)
                    l5.append(None)
                    l6.append(after_dollar)
                    l7.append(before_dollar_mid2+ ','+parts[4])
                elif '@' in parts[3]:
                    parts_mid = str(parts[2])
                    before_dollar_mid1 = parts_mid.split('$', 1)[0]
                    after_dollar_mid1 = parts_mid.split('$', 1)[1]
                    parts_mid = str(parts[3])
                    before_at_mid2 = parts_mid.split('@', 1)[0]
                    after_at_mid2 = parts_mid.split('@', 1)[1]
                    l3.append(before_dollar_mid1)
                    l4.append(after_dollar_mid1)
                    l5.append(parts[4])
                    l6.append(before_at_mid2)
                    l7.append(after_at_mid2)
        elif len(parts) == 3:
            l1.append(parts[0])
            l2.append(parts[1])
            l3.append(parts[2])
            l4.append(None)
            l5.append(None)
            l6.append(None)
            l7.append(None)
        # else:
        #     l1.append(sentence)
        #     l2.append(None)
        #     l3.append(None)
        #     l4.append(None)
        #     l5.append(None)
        #     l6.append(None)
        #     l7.append(None)
    # for i in range(len(l1)):
    #     y=i
    #     result.append('主语:'+str(l1[y])+" 谓语:"+str(l2[y])+' 宾语:'+str(l3[y])+' 补语谓语（针对主语）:'+str(l4[y])+' 补语宾语（针对主语）:'+str(l5[y])+' 补语谓语（针对宾语）:'+str(l6[y])+' 补语宾语（针对宾语）:'+str(l7[y]))
    # return result
    return l1,l2,l3,l4,l5,l6,l7

# for i in range(len(guanxi_)):
#     a1,a2,a3,a4,a5,a6,a7 = split_string(str(guanxi[i]))
#     for j in range(len(a1)):
#         l1.append(a1[j])
# print(len(l1))

year__ = []
name__ = []
w_num__ = []
num__ = []
w_type__ = []
type__ =[]
ju__ = []
guanxi__ = []
for i in range(len(num_)):
    sentences =str(guanxi_[i]).split(';')
    count = 0
    for sentence in sentences:
        if sentence.count(':') > 1:
            count+=1
    for j in range(count):
        year__.append(year_[i])
        name__.append(name_[i])
        w_num__.append(w_num_[i])
        num__.append(num_[i])
        w_type__.append(w_type_[i])
        type__.append(type_[i])
        ju__.append(ju_[i])
        guanxi__.append(guanxi_[i])
        
        


# for i in range(len(guanxi)):
#     print(split_string(str(guanxi[i])))
# r_num = []
# r_year = []
# r_name = []
# r_w_num = []
# r_type = []
# r_w_type = []
# r_ju = []
# r_guanxi = []
# r_result = []
r_zhu = []
r_wei = []
r_bin = []
r_1 =[]
r_2 = []
r_3 = []
r_4 = []
for i in guanxi_:
    a1,a2,a3,a4,a5,a6,a7 = split_string(str(i))
    for j in range(len(a1)):
        r_zhu.append(a1[j])
        r_wei.append(a2[j])
        r_bin.append(a3[j])
        r_1.append(a4[j])
        r_2.append(a5[j])
        r_3.append(a6[j])
        r_4.append(a7[j])
# for i in range(len(type)):
#     if type[i] not in not_name2:
#         # print(guanxi[i])
#         # r_num.append(num_[i])
#         # r_year.append(year_[i])
#         # r_name.append(name_[i])
#         # r_w_num.append(w_num_[i])
#         # r_w_type.append(w_type_[i])
#         # r_type.append(type_[i])
#         # r_ju.append(ju_[i])
#         # r_guanxi.append(guanxi_[i])
#         # r_result.append(split_string(str(guanxi[i])))
#         a1,a2,a3,a4,a5,a6,a7 = split_string(str(guanxi[i]))
#         for i in range(len(a1)):
#             r_zhu.append(a1[i])
#             r_wei.append(a2[i])
#             r_bin.append(a3[i])
#             r_1.append(a4[i])
#             r_2.append(a5[i])
#             r_3.append(a6[i])
#             r_4.append(a7[i])


data = {
    '序号': num__,
    '年份': year__,
    '篇名': name__,
    '贡献类型': type__,
    '贡献句': ju__,
    '实体及关系': guanxi__,
    # '解析实体及关系': r_result,
    '主语' : r_zhu,
    '谓语' : r_wei,
    '宾语' : r_bin,
    '补语谓语（针对主语）' : r_1,
    '补语宾语（针对主语）' : r_2,
    '补语谓语（针对宾语）' : r_3,
    '补语宾语（针对宾语）' : r_4

}

# 创建 DataFrame
df1 = pd.DataFrame(data)


# 将 DataFrame 保存到 Excel 文件
output_file = 'D://output.xlsx'
df1.to_excel(output_file, index=False)
print(f'Data saved to {output_file}')

