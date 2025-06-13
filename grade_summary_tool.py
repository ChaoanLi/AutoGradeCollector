#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学生成绩汇总工具
================

一个用于自动汇总学生成绩数据的工具，支持从多个Excel文件中提取学生信息和成绩数据。

主要功能：
- 自动发现和识别Excel文件
- 智能提取学生基本信息（学号、姓名、班级等）
- 汇总多种类型的成绩数据（作业、实验、讨论等）
- 生成完整的成绩汇总表

作者：Chaoan
许可：MIT License
"""

import pandas as pd
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')


class StudentDataExtractor:
    """
    学生数据提取器
    自动识别和提取学生名单中的关键信息
    """
    
    def __init__(self, data_directory='.'):
        """
        初始化数据提取器
        
        Args:
            data_directory (str): 数据文件所在目录，默认为当前目录
        """
        self.data_directory = data_directory
        self.all_files = []
        
    def discover_files(self):
        """
        发现并分析目录中的所有Excel文件
        
        Returns:
            list: 发现的Excel文件列表
        """
        files = []
        for ext in ['*.xlsx', '*.xls']:
            files.extend(Path(self.data_directory).glob(ext))
        
        self.all_files = [str(f) for f in files]
        print(f"发现 {len(self.all_files)} 个数据文件:")
        for i, file in enumerate(self.all_files, 1):
            print(f"  {i}. {Path(file).name}")
        return self.all_files


def create_grade_summary(data_directory='.'):
    """
    创建学生成绩汇总表
    
    功能说明：
    - 从CL文件提取学生信息（原序号、学号、姓名、班级、性别）
    - 汇总SA、LA、TL成绩
    - 智能学号识别：先找到"学号"列，然后向下搜索指定行数
    
    Args:
        data_directory (str): 数据文件所在目录
    
    Returns:
        pandas.DataFrame: 成绩汇总表
    """
    print("开始创建成绩汇总...")
    
    # 初始化数据提取器
    extractor = StudentDataExtractor(data_directory)
    extractor.discover_files()
    
    # 第一步：从CL文件提取学生信息
    print("\n第一步：从CL文件提取学生信息...")
    cl_students = []
    cl_files = [f for f in extractor.all_files if 'CL' in f and f.endswith('.xls')]
    
    for cl_file in cl_files:
        print(f"   处理 {Path(cl_file).name}...")
        
        try:
            # 读取整个文件 - 智能选择引擎
            df = None
            engines = ['openpyxl', 'xlrd'] if cl_file.endswith('.xls') else ['openpyxl']
            
            for engine in engines:
                try:
                    df = pd.read_excel(cl_file, header=None, engine=engine)
                    break
                except Exception as e:
                    if engine == engines[-1]:  # 最后一个引擎也失败了
                        raise e
                    continue
            
            # 方法一：智能表头识别
            header_row_idx = None
            col_mapping = {}
            
            # 在前20行中查找表头
            for i in range(min(20, len(df))):
                row_data = df.iloc[i].astype(str).tolist()
                
                # 检查这一行是否包含表头关键词
                found_headers = 0
                temp_mapping = {}
                
                for j, cell_value in enumerate(row_data):
                    cell_str = str(cell_value).strip().lower()
                    
                    if '学号' in cell_str or 'student' in cell_str or 'id' in cell_str:
                        temp_mapping['学号'] = j
                        found_headers += 1
                    elif '姓名' in cell_str or 'name' in cell_str or '名字' in cell_str:
                        temp_mapping['姓名'] = j
                        found_headers += 1
                    elif '班级' in cell_str or 'class' in cell_str or '专业' in cell_str:
                        temp_mapping['班级'] = j
                        found_headers += 1
                    elif '性别' in cell_str or 'gender' in cell_str or '男女' in cell_str:
                        temp_mapping['性别'] = j
                        found_headers += 1
                    elif '序号' in cell_str or 'no' in cell_str or '编号' in cell_str:
                        temp_mapping['序号'] = j
                        found_headers += 1
                
                # 如果找到至少2个表头，认为这是表头行
                if found_headers >= 2:
                    header_row_idx = i
                    col_mapping = temp_mapping
                    print(f"     找到表头行: 第{i+1}行，包含列: {list(col_mapping.keys())}")
                    break
            
            # 如果找到了表头，按列提取数据
            if header_row_idx is not None and '学号' in col_mapping:
                student_id_col = col_mapping['学号']
                
                # 从表头行之后开始，向下搜索30行查找学号数据
                search_end = min(header_row_idx + 31, len(df))
                for i in range(header_row_idx + 1, search_end):
                    row_data = df.iloc[i].astype(str).tolist()
                    
                    # 检查学号列是否包含学号数据
                    if student_id_col < len(row_data):
                        cell_value = str(row_data[student_id_col]).strip()
                        
                        # 灵活的学号识别：检查是否为有效学号格式
                        if (cell_value and cell_value != 'nan' and cell_value != '' and
                            len(cell_value) >= 8 and 
                            any(char.isdigit() for char in cell_value) and
                            not cell_value.replace('.', '').replace('-', '').isdigit()):
                            
                            student_id = cell_value
                            
                            # 根据列映射提取其他信息
                            student_name = ''
                            original_seq = ''
                            class_info = ''
                            gender_info = ''
                            
                            if '姓名' in col_mapping and col_mapping['姓名'] < len(row_data):
                                name_value = str(row_data[col_mapping['姓名']]).strip()
                                if name_value and name_value != 'nan':
                                    student_name = name_value
                            
                            if '班级' in col_mapping and col_mapping['班级'] < len(row_data):
                                class_value = str(row_data[col_mapping['班级']]).strip()
                                if class_value and class_value != 'nan':
                                    class_info = class_value
                            
                            if '性别' in col_mapping and col_mapping['性别'] < len(row_data):
                                gender_value = str(row_data[col_mapping['性别']]).strip()
                                if gender_value and gender_value != 'nan':
                                    gender_info = gender_value
                            
                            if '序号' in col_mapping and col_mapping['序号'] < len(row_data):
                                seq_value = str(row_data[col_mapping['序号']]).strip()
                                if seq_value and seq_value != 'nan' and seq_value.isdigit():
                                    original_seq = seq_value
                            
                            # 添加学生信息
                            student_data = {
                                '学号': student_id,
                                '姓名': student_name,
                                '原序号': original_seq,
                                '班级': class_info,
                                '性别': gender_info,
                                '来源文件': Path(cl_file).name
                            }
                            cl_students.append(student_data)
                            print(f"     找到学生: {student_id} {student_name} (原序号:{original_seq}, 班级:{class_info}, 性别:{gender_info})")
            
            # 方法二：如果没找到表头，使用改进的逐行扫描方法作为备用
            else:
                print(f"     未找到标准表头，使用改进的逐行扫描方法...")
                
                # 先查找包含"学号"字符的单元格位置
                student_id_col = None
                header_row = None
                
                for i in range(min(20, len(df))):  # 在前20行中查找"学号"
                    row_data = df.iloc[i].astype(str).tolist()
                    for j, cell_value in enumerate(row_data):
                        cell_str = str(cell_value).strip().lower()
                        if '学号' in cell_str or 'student' in cell_str or 'id' in cell_str:
                            student_id_col = j
                            header_row = i
                            print(f"     找到学号列: 第{i+1}行第{j+1}列")
                            break
                    if student_id_col is not None:
                        break
                
                # 如果找到了学号列，在该列向下搜索30行
                if student_id_col is not None and header_row is not None:
                    search_end = min(header_row + 31, len(df))
                    for i in range(header_row + 1, search_end):
                        if i < len(df):
                            row_data = df.iloc[i].astype(str).tolist()
                            if student_id_col < len(row_data):
                                cell_str = str(row_data[student_id_col]).strip()
                                # 检查是否为有效学号格式
                                if (cell_str and cell_str != 'nan' and cell_str != '' and
                                    len(cell_str) >= 8 and 
                                    any(char.isdigit() for char in cell_str) and
                                    not cell_str.replace('.', '').replace('-', '').isdigit()):
                                    
                                    student_id = cell_str
                                    
                                    # 尝试从同一行找到其他信息
                                    student_name = ''
                                    original_seq = ''
                                    class_info = ''
                                    gender_info = ''
                                    
                                    # 检查同一行的其他列
                                    for k in range(len(row_data)):
                                        if k != student_id_col and row_data[k] and str(row_data[k]) != 'nan':
                                            potential_value = str(row_data[k]).strip()
                                            
                                            # 识别姓名（支持各种姓名格式）
                                            is_student_id = (len(potential_value) >= 8 and 
                                                           any(char.isdigit() for char in potential_value) and
                                                           not potential_value.replace('.', '').replace('-', '').isdigit())
                                            if not student_name and not (is_student_id or potential_value.replace('.', '').isdigit()):
                                                if len(potential_value) >= 2 and len(potential_value) <= 15:
                                                    # 排除明显不是姓名的内容
                                                    if not any(keyword in potential_value.lower() for keyword in ['class', 'grade', '班', '级', '系', '院', '专业']):
                                                        student_name = potential_value
                                            
                                            # 识别序号（纯数字且较短）
                                            elif not original_seq and potential_value.isdigit() and len(potential_value) <= 3:
                                                original_seq = potential_value
                                            
                                            # 识别性别
                                            elif not gender_info and potential_value in ['男', '女']:
                                                gender_info = potential_value
                                            
                                            # 识别班级信息
                                            elif not class_info:
                                                # 检查是否包含班级特征
                                                if (any(keyword in potential_value for keyword in ['T', '工商', '管理', '经济', '金融', '会计']) and 
                                                    len(potential_value) <= 20 and len(potential_value) >= 3):
                                                    class_info = potential_value
                                    
                                    # 添加学生信息
                                    if student_id:
                                        student_data = {
                                            '学号': student_id,
                                            '姓名': student_name,
                                            '原序号': original_seq,
                                            '班级': class_info,
                                            '性别': gender_info,
                                            '来源文件': Path(cl_file).name
                                        }
                                        cl_students.append(student_data)
                                        print(f"     找到学生: {student_id} {student_name} (原序号:{original_seq}, 班级:{class_info}, 性别:{gender_info})")
                else:
                    print(f"     未找到学号列，跳过此文件")
        
        except Exception as e:
            print(f"     处理失败: {e}")
    
    print(f"   共找到 {len(cl_students)} 名学生")
    
    # 第二步：收集SA成绩
    print("\n第二步：收集SA表得分...")
    sa_scores = {}
    sa_files = [f for f in extractor.all_files if 'SA' in f and f.endswith('.xls')]
    
    for sa_file in sa_files:
        try:
            # 智能选择引擎
            df = None
            engines = ['openpyxl', 'xlrd'] if sa_file.endswith('.xls') else ['openpyxl']
            
            for engine in engines:
                try:
                    df = pd.read_excel(sa_file, engine=engine)
                    break
                except Exception as e:
                    if engine == engines[-1]:
                        raise e
                    continue
            score_col = None
            
            for col in df.columns:
                if '得分' in str(col):
                    score_col = col
                    break
            
            if score_col:
                for _, row in df.iterrows():
                    if pd.notna(row['学号']) and pd.notna(row[score_col]):
                        student_id = str(row['学号']).strip()
                        score = row[score_col]
                        
                        if student_id not in sa_scores:
                            sa_scores[student_id] = {}
                        sa_scores[student_id][Path(sa_file).name.replace('.xls', '')] = score
        except Exception as e:
            print(f"     处理 {Path(sa_file).name} 失败: {e}")
    
    print(f"   收集到 {len(sa_scores)} 名学生的SA成绩")
    
    # 第三步：收集LA成绩
    print("\n第三步：收集LA表得分...")
    la_scores = {}
    la_files = [f for f in extractor.all_files if 'LA' in f and f.endswith('.xls')]
    
    for la_file in la_files:
        try:
            # 智能选择引擎
            df = None
            engines = ['openpyxl', 'xlrd'] if la_file.endswith('.xls') else ['openpyxl']
            
            for engine in engines:
                try:
                    df = pd.read_excel(la_file, engine=engine)
                    break
                except Exception as e:
                    if engine == engines[-1]:
                        raise e
                    continue
            score_col = None
            
            for col in df.columns:
                if '得分' in str(col):
                    score_col = col
                    break
            
            if score_col:
                for _, row in df.iterrows():
                    if pd.notna(row['学号']) and pd.notna(row[score_col]):
                        student_id = str(row['学号']).strip()
                        score = row[score_col]
                        
                        if student_id not in la_scores:
                            la_scores[student_id] = {}
                        la_scores[student_id][Path(la_file).name.replace('.xls', '')] = score
        except Exception as e:
            print(f"     处理 {Path(la_file).name} 失败: {e}")
    
    print(f"   收集到 {len(la_scores)} 名学生的LA成绩")
    
    # 第四步：收集TL讨论成绩
    print("\n第四步：收集TL表讨论成绩...")
    tl_scores = {}
    tl_files = [f for f in extractor.all_files if 'TL' in f]
    
    for tl_file in tl_files:
        try:
            # 智能选择引擎
            df = None
            engines = ['openpyxl', 'xlrd'] if tl_file.endswith('.xls') else ['openpyxl']
            
            for engine in engines:
                try:
                    df = pd.read_excel(tl_file, engine=engine)
                    break
                except Exception as e:
                    if engine == engines[-1]:
                        raise e
                    continue
            
            if '讨论/' in df.columns:
                for _, row in df.iterrows():
                    if pd.notna(row['学号']) and pd.notna(row['讨论/']):
                        student_id = str(row['学号']).strip()
                        score = row['讨论/']
                        tl_scores[student_id] = score
        except Exception as e:
            print(f"     处理 {Path(tl_file).name} 失败: {e}")
    
    print(f"   收集到 {len(tl_scores)} 名学生的TL讨论成绩")
    
    # 第五步：合并数据
    print("\n第五步：合并所有成绩数据...")
    
    # 构建列名
    columns = ['序号', '原序号', '学号', '姓名', '班级', '性别', '来源文件']
    
    # 添加所有SA列
    all_sa_files = sorted(list(set([k for scores in sa_scores.values() for k in scores.keys()])))
    columns.extend(all_sa_files)
    
    # 添加所有LA列
    all_la_files = sorted(list(set([k for scores in la_scores.values() for k in scores.keys()])))
    columns.extend(all_la_files)
    
    # 添加TL讨论列
    columns.append('TL讨论')
    
    # 创建汇总数据
    summary_data = []
    
    for i, student in enumerate(cl_students, 1):
        student_id = student['学号']
        row = [
            i,  # 新序号
            student.get('原序号', ''),  # 原序号
            student_id,  # 学号
            student.get('姓名', ''),  # 姓名
            student.get('班级', ''),  # 班级
            student.get('性别', ''),  # 性别
            student['来源文件']  # 来源文件
        ]
        
        # 添加SA成绩（空成绩填0）
        for sa_file in all_sa_files:
            score = 0
            if student_id in sa_scores and sa_file in sa_scores[student_id]:
                score = sa_scores[student_id][sa_file]
            row.append(score)
        
        # 添加LA成绩（空成绩填0）
        for la_file in all_la_files:
            score = 0
            if student_id in la_scores and la_file in la_scores[student_id]:
                score = la_scores[student_id][la_file]
            row.append(score)
        
        # 添加TL讨论成绩（空成绩填0）
        tl_score = tl_scores.get(student_id, 0)
        row.append(tl_score)
        
        summary_data.append(row)
    
    # 创建DataFrame
    summary_df = pd.DataFrame(summary_data, columns=columns)
    
    print(f"\n成绩汇总完成")
    print(f"   总计 {len(summary_df)} 名学生")
    print(f"   包含 {len(all_sa_files)} 个SA成绩项")
    print(f"   包含 {len(all_la_files)} 个LA成绩项")
    print(f"   包含 1 个TL讨论成绩")
    
    return summary_df


def export_summary_with_stats(summary_df, output_filename="学生成绩汇总表.xlsx"):
    """
    导出成绩汇总表并显示统计信息
    
    Args:
        summary_df (pandas.DataFrame): 成绩汇总表
        output_filename (str): 输出文件名
    """
    if summary_df is None or len(summary_df) == 0:
        print("没有找到任何学生数据")
        return
    
    print(f"\n=== 成绩汇总结果 ===")
    print(f"汇总表包含 {len(summary_df)} 名学生的成绩")
    print(f"汇总表包含 {len(summary_df.columns)} 列数据")
    
    print(f"\n列名清单:")
    for i, col in enumerate(summary_df.columns, 1):
        print(f"  {i:2d}. {col}")
    
    print(f"\n前5名学生的成绩预览:")
    print(summary_df.head().to_string(index=False))
    
    # 导出汇总结果
    summary_df.to_excel(output_filename, index=False)
    print(f"\n成绩汇总表已导出至: {output_filename}")
    
    # 检查数据完整性
    print(f"\n数据完整性检查:")
    for col in summary_df.columns:
        if col in ['SA1', 'SA2', 'SA3', 'SA4', 'SA5', 'SA6', 'SA7', 'LA1', 'LA2', 'LA3', 'LA4', 'LA5', 'LA6', 'TL讨论']:
            empty_count = summary_df[col].apply(lambda x: x == '' or pd.isna(x)).sum()
            zero_count = summary_df[col].apply(lambda x: x == 0).sum()
            non_zero_count = summary_df[col].apply(lambda x: x != 0 and x != '' and pd.notna(x)).sum()
            print(f"  {col}: {empty_count} 个空值, {zero_count} 个0值, {non_zero_count} 个有效成绩")
    
    # 基本统计信息
    print(f"\n统计信息:")
    
    # 统计基本信息完整性
    basic_info_stats = {}
    for col in ['原序号', '姓名', '班级', '性别']:
        if col in summary_df.columns:
            non_empty = summary_df[col].apply(lambda x: x != '' and pd.notna(x)).sum()
            basic_info_stats[col] = non_empty
    
    print(f"  基本信息完整性:")
    for info, count in basic_info_stats.items():
        print(f"    {info}: {count} 人有数据")
    
    # 按来源文件统计
    if '来源文件' in summary_df.columns:
        file_counts = summary_df['来源文件'].value_counts()
        print(f"  按来源文件统计:")
        for file_name, count in file_counts.items():
            print(f"    {file_name}: {count} 人")
    
    # 性别分布统计
    if '性别' in summary_df.columns:
        gender_counts = summary_df['性别'].value_counts()
        if len(gender_counts) > 0:
            print(f"  性别分布:")
            for gender, count in gender_counts.items():
                if gender and str(gender) != '':
                    print(f"    {gender}: {count} 人")


def main():
    """
    主函数 - 执行完整的成绩汇总流程
    """
    print("=== 学生成绩汇总工具 ===")
    print("开始执行成绩汇总...")
    
    # 创建成绩汇总
    summary_df = create_grade_summary()
    
    # 导出结果并显示统计
    export_summary_with_stats(summary_df)
    
    print("\n=== 处理完成 ===")


if __name__ == "__main__":
    main() 