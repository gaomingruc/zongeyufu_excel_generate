import os
import pandas as pd


def get_data_address_lists():
    """得到当前文件夹下的数据源路径"""
    data_address_yibao_list = list()
    data_address_zonge_list = list()
    for root, dirs, files in os.walk("."):
        for file in files:
#            print(file)
            if file[:2] == "医保":
#                print("医保:", file)
                data_address_yibao_list.append(file)
            elif file[:2] == "11":
#                print("总额:", file)
                data_address_zonge_list.append(file)
    return data_address_yibao_list, data_address_zonge_list


def get_result_yibao(data_address_list):
    """输入想合并的文件的路径，返回各科医保结果"""
    data_list = list()
    for data_address in data_address_list:
        print(data_address)
        data = pd.read_excel(data_address, header = 8)
        data["是否是手术科室"] = data["是否是手术科室"].fillna(method="ffill")
        data["申请科室"] = data["申请科室"].fillna(value="缺失值")
        data["类型"] = data_address.split(" ")[1]
        data["时间"] = data_address.split(" ")[2].split(".")[0]
        data_list.append(data)
        
    department_set = set()
    for data in data_list:
        department_set = department_set | set(data["申请科室"])
    department_list = list(department_set)
    department_list.remove("缺失值")
    
    for department in department_list:
        print(department)
        rows = list()
        for data in data_list:
            row = data[data["申请科室"] == department]
            rows.append(row)
        result = pd.concat(rows).reset_index(drop = True)
#        print(result.columns)
        result = result[['申请科室', '类型', '时间', '是否是手术科室', '总费用', '今年累计总费用', '去年总费用', '去年累计总费用', '当月纵比',
       '累计增幅总费用', '基金支付费用', '基金申报额今年累计', '基金当月纵比', '去年基金支付费用', '去年基金支付费用累计',
       '基金支付费用累计增幅', '人次', '人次当月纵比', '人次累计', '去年人次', '去年人次累计', '人次累计增幅',
       '医保次均费用', '次均当月纵比', '今年累计次均费用', '去年次均费用', '去年累计次均费用', '次均费用累计增幅', '人数',
       '累计人数', '医保人数当月纵比', '上年人数', '上年累计人数', '人数累计纵比', '次头比', '累计人次人头比',
       '次头当月纵比', '上年人次人头比', '上年累计人次人头比', '人次人头比累计增幅', '自费费用', '自费金额累计',
       '去年自费金额', '自费当月纵比', '去年自费金额累计', '自费金额累计增幅', '药品费用', '药品费用累计',
       '药品费用当月纵比', '去年药品费用', '去年药品费用累计', '药品费用累计增幅', '现金支付费用', '自付金额累计',
       '现金支付当月纵比', '去年现金自付', '去年现金自付累计', '现金自付累计增幅', '检查治疗费', '检查治疗当月纵比',
       '今年检查治疗费累计', '去年检查治疗费', '去年检查治疗费累计', '检查治疗费累计增幅', '卫生材料费用', '今年卫生材料费累计',
       '卫生材料费当月纵比', '去年卫生材料费', '去年卫生材料费累计', '卫生材料费累计增幅', '其他费用', '今年其他费用累计',
       '其他费用当月纵比 ', '去年其他费用', '去年其他费用累计', '其他费用累计增幅', '药占比', '今年累计药占比',
       '去年药占比', '药占比当月纵比', '去年累计药占比', '药占比累计增幅', '医保人均费用', '人均费用今年累计',
       '去年人均费用', '人均当月纵比', '人均累计增幅', '去年人均累计', '药品耗材占比', '累计药品耗材比',
       '去年药品累计耗材比', '药品耗材当月纵比', '去年药品耗材比', '药品耗材累计增幅']]
        print(result)
        print("\n")
        result_name = department + "_" + list(set(result["时间"]))[0] + "_" + "医保" + ".xlsx"
        print(result_name)
        result.to_excel(result_name, index=False)
        print("-" * 20)
        

def get_result_zonge(data_address_list):
    """输入想合并的文件的路径，返回各科总额结果"""
    data_list = list()
    for data_address in data_address_list:
        print(data_address)
        data = pd.read_excel(data_address, header = 8)
        data["是否是手术科室"] = data["是否是手术科室"].fillna(method="ffill")
        data["申请科室"] = data["申请科室"].fillna(value="缺失值")
        data["类型"] = data_address.split(" ")[-2]
        data["时间"] = data_address.split(" ")[-1].split(".")[0]
        data_list.append(data)
        
    department_set = set()
    for data in data_list:
        department_set = department_set | set(data["申请科室"])
    department_list = list(department_set)
    department_list.remove("缺失值")
    
    for department in department_list:
        print(department)
        rows = list()
        for data in data_list:
            row = data[data["申请科室"] == department]
            rows.append(row)
        result = pd.concat(rows).reset_index(drop = True)
#        print(result.columns)
        result = result[['申请科室', '类型', '时间', '是否是手术科室', '当期总费用', '当期累计总费用', '上期总费用', '上期累计总费用', '当期纵比',
       '累计纵比', '当期基金支付费用', '基金申报额当期累计', '上期基金支付费用', '上期基金支付费用累计', '基金当月纵比',
       '基金支付费用累计纵比', '当前人次', '当期人次累计', '上期人次', '上期人次累计', '人次当月纵比', '人次累计纵比',
       '当期医保次均费用', '当期累计次均费用', '上期次均费用', '上期累计次均费用', '次均当月纵比', '次均费用累计纵比',
       '当期人数', '当期累计人数', '上期人数', '上期累计人数', '医保人数当月纵比', '人数累计纵比', '当期次头比',
       '当期累计人次人头比', '上期人次人头比', '上期累计人次人头比', '次头当月纵比', '人次人头累计纵比', '当期自费费用',
       '当期自费金额累计', '上期自费金额', '上期自费金额累计', '自费当月纵比', '自费金额累计纵比', '当期药品费用',
       '当期药品费用累计', '上期药品费用', '上期药品费用累计', '药品费用当月纵比', '药品费用累计纵比', '当期现金支付费用',
       '当期自付金额累计', '上期现金自付', '上期现金自付累计', '现金支付当月纵比', '现金自付累计纵比', '当期检查治疗费',
       '当期检查治疗费累计', '上期检查治疗费', '上期检查治疗费累计', '检查治疗当月纵比', '检查治疗费累计纵比',
       '当期卫生材料费用', '当期卫生材料费累计', '上期卫生材料费', '上期卫生材料费累计', '卫生材料费当月纵比',
       '卫生材料费累计纵比', '当期其他费用', '当期其他费用累计', '上期其他费用', '上期其他费用累计', '其他费用当月纵比 ',
       '其他费用累计纵比', '当期药占比', '当期累计药占比', '上期药占比', '上期累计药占比', '药占比当月纵比',
       '药占比累计纵比', '当期医保人均费用', '人均费用当期累计', '上期人均费用', '上期人均累计', '人均当月纵比',
       '人均累计纵比', '当期药品耗材占比', '当期累计药品耗材比', '上期药品耗材比', '上期药品累计耗材比', '药品耗材当月纵比',
       '药品耗材累计纵比']]
        print(result)
        print("\n")
        result_name = department + "_" + list(set(result["时间"]))[0] + "_" + "总额" + ".xlsx"
        print(result_name)
        result.to_excel(result_name, index=False)
        print("-" * 20)
        

if __name__ == "__main__":
    data_address_yibao_list, data_address_zonge_list = get_data_address_lists()
#    print(data_address_yibao_list)
#    print(data_address_zonge_list)
    get_result_yibao(data_address_yibao_list)
    get_result_zonge(data_address_zonge_list)
    print("DONE")