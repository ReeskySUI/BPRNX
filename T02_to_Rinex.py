import os
import glob
import re
import pandas as pd
from collections import defaultdict
import subprocess
import time as t

# 获取非默认安装convertToRINEX软件所在文件夹路径
rf = input('请输入convertToRINEX软件所在文件夹路径(默认为C:\Program Files (x86)\Trimble\convertToRINEX,可按回车跳过):')
# 读取xlsx文件，跳过前两行，并只读取B, C, D, E列
def read_xlsx(file_path):
    # 使用pandas读取文件，skiprows跳过前两行，usecols指定需要的列
    df = pd.read_excel(file_path ,skiprows=2 ,usecols="B:E",header=None)
    # 去除所有列都为NaN的行
    df = df.dropna(how='any')
    df[1]=df[1].astype(int)
    df[3]=df[3].astype(int)
    # 将数据转换为二维列表
    data = df.values.tolist()
    return data
# 读取xlsx文件到data列表
file_path = input('请输入静态外业测量xlsx文件路径:').replace('\\','/')
data = read_xlsx(file_path)

# 打印测试
# print(data)

# 获取指定文件夹中所有扩展名为 .T02 的文件的名称和路径
def get_T02_files(directory):
    # 使用 glob 模块查找所有 .T02 文件
    file_paths = glob.glob(os.path.join(directory, "*.T02"))
    
    # 提取文件名（不含路径）
    file_names = [os.path.basename(file_path) for file_path in file_paths]
    
    # 返回文件名和对应的路径
    return file_names, file_paths

# 读取原始数据
directory = input('请输入原始数据文件夹路径：')
file_names, file_paths = get_T02_files(directory)
dzb = dict(zip(file_names, file_paths))
# 打印文件名和路径测试
# print("文件名:", file_names)
# print("文件路径:", file_paths)

# 从文件名中提取仪器号和时间号（假设文件名格式为：4位仪器号+4位时间号.T02）
def extract_instrument_time(file_name):
    # 使用正则表达式匹配4位的仪器号和4位的时间号
    match = re.match(r'(\d{4})(\d{4})\.T02', file_name)
    if match:
        instrument_num = match.group(1)  # 仪器号
        time_num = match.group(2)        # 时间号
        return instrument_num, time_num
    return None, None

# 匹配文件名和data列表
def match_files_to_data(file_names, data):
    # 提取文件名中的仪器号和时间号
    files_info = [(file_name, *extract_instrument_time(file_name)) for file_name in file_names]
    # print(files_info)

    # 创建一个字典，按仪器号后分组
    file_dict = defaultdict(list)
    for file_name, file_instrument_num, file_time_num in files_info:
        if file_instrument_num:  # 确保仪器号不为空
            file_dict[file_instrument_num].append((file_name, file_time_num))

    # 按顺序对data列表进行匹配
    matched_results = []
    matched_files = set()  # 用于记录已经匹配的文件
    # 解包每行数据，按照表格行顺序进行匹配
    for row in data:
        serial_num, point_name, instrument_num, instrument_height = row
        # 将仪器号转为字符串以便处理
        instrument_num_str = str(instrument_num)
        
        # 查找与仪器号后2/3/4位匹配的文件信息
        matched = False
        for file_instrument_num in file_dict:
            # 文件名中的仪器号应该是完整的 4 位
            if file_instrument_num.endswith(instrument_num_str):
                for file_name, file_time_num in file_dict[file_instrument_num]:
                    if file_name not in matched_files:
                        matched_results.append({
                            "file_name": file_name,
                            "data_row": row,
                            "file_time_num": file_time_num
                        })
                        matched_files.add(file_name)  # 标记文件为已匹配
                        matched = True
                        break  # 只匹配第一个找到的文件，跳出循环
            if matched:
                break  # 如果已经找到匹配的文件，跳出外层循环
        
        if not matched:
            print(f"未找到与仪器号 {instrument_num} 匹配的文件")

    return matched_results

# 匹配文件名和测量数据
matched_files = match_files_to_data(file_names, data)

# 自动获取输出文件夹
output_path = input('请输入标准数据输出文件夹：')
# 测试输出路径
# print(output_path)

start = t.time()
mf = 'C:\Program Files (x86)\Trimble\convertToRINEX'
if rf != '':
    mf = rf
# 生成匹配结果日志文件，核准数据
with open("log.txt",'a',encoding='utf-8') as f:
    f.truncate(0)
    a = len(matched_files)
    f.write('文件总数： %d \n\n' %a)
    f.write('输出位置： %s \n\n' %output_path)
    for match in matched_files:
        fn = dzb[match['file_name']]# 原始文件
        h  = match['data_row'][-1]# 仪器高
        mo = match['data_row'][1]# 点名MAKER NAME
        f.write(f"文件名: {match['file_name']}, 对应数据: {match['data_row']}\n")

        # cmd = f"convertToRinex {fn} -p {output_path} -h {h} -mo {mo}"

        # 批处理T02文件
        cmd1 = subprocess.run(['convertToRinex',fn,'-p',output_path,'-h',str(h),'-mo',str(mo)],shell=True,cwd=mf)
        # print(cmd1.returncode) # 返回码
        # print(cmd1.stdout) # 标准输出
        # print(cmd1.stderr) # 标准错误
        print('当前进度：',matched_files.index(match)+1,'/',len(matched_files),'|',int((matched_files.index(match)+1)/len(matched_files)*100),'%')
end = t.time()
time_sum = int(end - start)
# print(time_sum,int(time_sum))
print("已完成，程序用时%ss\nlog.txt日志文件已生成\n标准文件请查看如下目录：\n%s" %(time_sum,output_path))
input('按任意键退出程序...')