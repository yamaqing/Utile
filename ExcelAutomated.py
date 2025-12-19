import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os
import math
import logging

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ==========================================
# 用户配置区域 (可以在这里修改时间设置)
# ==========================================

# 1. 设置起止日期 (格式: YYYY-MM-DD)
START_DATE = '2024-12-12'
END_DATE = '2025-12-11'

# 2. 调休设置：即使是周末也需要工作的日期 (格式: YYYY-MM-DD)
# 例如: '2024-02-04' (周日), '2024-02-18' (周日)
SPECIAL_WORK_DAYS = [

]

# 目标生成行数设置
# 如果设置为 0，则默认使用 Excel 文件中原有的行数
# 如果设置为 > 0 的整数 (例如 1000)，则会忽略 Excel 原有数据量，强制生成指定数量的数据
TARGET_ROW_COUNT = 0 

# 3. 调休设置：即使是周一到周五也是休息的日期 (节假日) (格式: YYYY-MM-DD)
# 例如: '2024-01-01' (元旦), '2024-02-10' 到 '2024-02-17' (春节)
SPECIAL_OFF_DAYS = [
    
]

# 4. 区域分布设置 (需求2)
# 格式: {'区域名称': 数量}
# 注意：总数量最好与Excel行数匹配，或者让程序自动根据比例填充
REGION_CONFIG = {
    '区域1': 500,
    '区域2': 92,
    '区域3': 150,
    '区域4': 91,
    '区域5': 101,
    '区域6': 70
}
# 模式选择: 
# 'exact': 严格按照具体数量生成 (如果总数不够Excel行数，会报错或循环)
# 'ratio': 按照比例自动分配给所有行 (推荐)
REGION_MODE = 'ratio' 

# 5. 区域与问题类型联动设置 (需求3)
# 格式: {'区域名称': {'问题类型A': 数量, '问题类型B': 数量}}
# 这里的数量可以是具体数值，也可以是比例。
# 程序会自动根据该区域的实际总行数，按这里设定的比例进行分配。
REGION_PROBLEM_CONFIG = {
    '区域1': {'类型1': 105, '类型2': 26, '类型3': 35,'类型4': 135,'类型5': 186 ,'类型6': 13, '类型7': 0}, # 如果区域1总共有100行，则严格按此数量分配
    '区域2': {'类型1': 23,'类型2': 0, '类型3': 0,'类型4': 36,'类型5': 32 ,'类型6': 1, '类型7': 0},
    '区域3': {'类型1': 12, '类型2': 50, '类型3': 3,'类型4': 41,'类型5': 41 ,'类型6': 3, '类型7': 0},
    '区域4': {'类型1': 24, '类型2': 0, '类型3': 1,'类型4': 29,'类型5': 35 ,'类型6': 2, '类型7': 0},
    '区域5': {'类型1': 6, '类型2': 0, '类型3': 0,'类型4': 35,'类型5': 0 ,'类型6': 60, '类型7': 0},
    '区域6': {'类型1': 0, '类型2': 20, '类型3': 20,'类型4': 0,'类型5': 10 ,'类型6': 0, '类型7': 20},
}
# 默认问题类型 (如果区域没有在配置中找到)
DEFAULT_PROBLEM_TYPE = '默认类型'

# 6. 问题与解决的匹配设置 (需求4)
# 控制开关: True = 严格匹配 (C和D必须成对出现), False = 宽松匹配 (同类型下随机组合)
STRICT_MATCH = True

# 格式: {'问题类型': [('具体问题', '解决方案'), ...]}
# 注意：这里的 key 必须与上面的 REGION_PROBLEM_CONFIG 中的类型名称一致
PROBLEM_SOLUTION_CONFIG = {
    '类型1': [
        ('类型1-问题A', '类型1-解决A'), 
        ('类型1-问题B', '类型1-解决B'),
        ('类型1-问题C', '类型1-解决C')
    ],
    '类型2': [
        ('类型2-问题A', '类型2-解决A'), 
        ('类型2-问题B', '类型2-解决B')
    ],
    '类型3': [
        ('类型3-问题A', '类型3-解决A')
    ],
    '类型4': [('类型4-问题A', '类型4-解决A')],
    '类型5': [('类型5-问题A', '类型5-解决A')],
    '类型6': [('类型6-问题A', '类型6-解决A')],
    '类型7': [('类型7-问题A', '类型7-解决A')],
    '默认类型': [('未知问题', '待定解决')]
}
#设置生的Excel表格的存储路径和名字
DEFAULT_DIR = r'e:\Python'
EXCEL_PATH = os.path.join(DEFAULT_DIR, '生成数据表_text2.xlsx')
# ==========================================
# 核心逻辑代码
# ==========================================

def get_valid_workdays(start_str, end_str, special_work, special_off):
    """
    生成指定范围内的所有有效工作日
    """
    start = datetime.strptime(start_str, '%Y-%m-%d')
    end = datetime.strptime(end_str, '%Y-%m-%d')
    
    valid_days = []
    current = start
    
    # 转换为集合提高查找效率
    work_set = set(special_work)
    off_set = set(special_off)
    
    while current <= end:
        date_str = current.strftime('%Y-%m-%d')
        weekday = current.weekday() # 0=周一, 6=周日
        
        is_weekend = weekday >= 5
        
        # 判定逻辑：
        # 1. 如果在“特殊上班日”里，不论星期几，都是工作日
        # 2. 否则，如果在“特殊休息日”里，不论星期几，都是休息日
        # 3. 否则，如果是周末，休息；如果是平日，工作
        
        is_workday = False
        
        if date_str in work_set:
            is_workday = True
        elif date_str in off_set:
            is_workday = False
        else:
            is_workday = not is_weekend
            
        if is_workday:
            valid_days.append(date_str)
            
        current += timedelta(days=1)
        
    return valid_days

def process_excel(file_path):
    if not os.path.exists(file_path):
        logger.error(f"错误: 文件 {file_path} 不存在")
        return

    logger.info(f"正在读取文件: {file_path}")
    try:
        # 读取 Excel
        df = pd.read_excel(file_path)
        
        # 确定总行数
        original_count = len(df)
        
        if TARGET_ROW_COUNT > 0:
            logger.info(f"检测到目标行数设置: {TARGET_ROW_COUNT}")
            original_count = TARGET_ROW_COUNT
            # 如果目标行数大于当前行数，或者为了清理旧数据，我们可以重建一个空的 DataFrame
            # 这里我们直接创建一个全新的 DataFrame，长度为 TARGET_ROW_COUNT
            # 这样可以保证数据的绝对纯净和长度一致
            df = pd.DataFrame(index=range(original_count))
            logger.info(f"已重置数据表，将生成 {original_count} 行数据")
        else:
            logger.info(f"读取成功，使用原有行数: {original_count}")
        
        if REGION_MODE == 'exact':
            total_regions = sum(REGION_CONFIG.values())
            if TARGET_ROW_COUNT == 0 and original_count != total_regions:
                logger.info(f"检测到 exact 模式，自动设置行数为区域总数: {total_regions}")
                original_count = total_regions
                df = pd.DataFrame(index=range(original_count))
        
        # 生成有效工作日池
        logger.info("正在计算有效工作日...")
        valid_workdays = get_valid_workdays(START_DATE, END_DATE, SPECIAL_WORK_DAYS, SPECIAL_OFF_DAYS)
        logger.info(f"在 {START_DATE} 到 {END_DATE} 期间，共有 {len(valid_workdays)} 个有效工作日")
        
        if not valid_workdays:
            logger.error("错误: 没有找到有效的工作日，请检查日期范围和调休设置")
            return

        # --- 需求1：时间分配 ---
        logger.info("正在分配日期...")
        random_dates = np.random.choice(valid_workdays, size=original_count)
        df['时间'] = random_dates
        
        # 按时间排序 (从小到大)
        df = df.sort_values('时间')
        
        # --- 需求2：区域分配 ---
        logger.info(f"正在分配区域 (模式: {REGION_MODE})...")
        
        regions_list = []
        if REGION_MODE == 'ratio':
            # 按比例生成
            total_weight = sum(REGION_CONFIG.values())
            
            # 使用 numpy 的 choice 进行加权随机选择
            # p 是概率分布
            probs = [count / total_weight for count in REGION_CONFIG.values()]
            names = list(REGION_CONFIG.keys())
            
            regions_list = np.random.choice(names, size=original_count, p=probs)
            
        else: # 'exact' 模式
            # 严格按照数量生成，如果不足则循环，如果多了则截断
            temp_list = []
            for name, count in REGION_CONFIG.items():
                temp_list.extend([name] * count)
            
            # 如果配置数量少于行数，需要循环补充
            if len(temp_list) < original_count:
                logger.warning(f"警告: 配置的区域总数 ({len(temp_list)}) 少于 Excel 行数 ({original_count})，将循环使用。")
                # 计算需要重复多少次
                repeats = math.ceil(original_count / len(temp_list))
                temp_list = temp_list * repeats
            
            # 截取到需要的长度
            regions_list = temp_list[:original_count]
            # 打乱顺序
            np.random.shuffle(regions_list)
            
        df['区域'] = regions_list
        
        # --- 需求3：区域与问题类型联动 (精确数量控制) ---
        logger.info("正在根据区域分配问题类型 (严格按照配置数量)...")
        
        # 初始化 '问题类型' 列
        df['问题类型'] = DEFAULT_PROBLEM_TYPE
        
        # 按区域分组处理
        for region_name, group_indices in df.groupby('区域').groups.items():
            # 获取该区域在当前Excel中的实际行数
            actual_count = len(group_indices)
            
            if actual_count == 0:
                continue
                
            # 获取该区域的配置
            if region_name in REGION_PROBLEM_CONFIG:
                type_config = REGION_PROBLEM_CONFIG[region_name]
                
                # 计算配置中的总数
                config_total = sum(type_config.values())
                
                if config_total == 0:
                    continue
                
                # 检查配置总数与实际行数是否一致
                if config_total != actual_count:
                    logger.warning(f"警告: 区域 {region_name} 的问题类型配置总数 ({config_total}) 与实际行数 ({actual_count}) 不一致！")
                    logger.warning(f"将按照配置的问题类型比例重新计算数量")
                
                # 生成问题类型列表
                types_list = []
                for t_name, t_val in type_config.items():
                    if t_val > 0:  # 只处理数量大于0的类型
                        types_list.extend([t_name] * t_val)
                
                # 检查生成的类型列表长度是否与实际行数一致
                if len(types_list) != actual_count:
                    logger.warning(f"警告: 区域 {region_name} 生成的问题类型数量 ({len(types_list)}) 与实际行数 ({actual_count}) 不一致！")
                    
                    # 重新计算问题类型数量
                    types_list = []
                    for t_name, t_val in type_config.items():
                        if config_total > 0:
                            # 按比例计算数量
                            count = int(round((t_val / config_total) * actual_count))
                            if count > 0:
                                types_list.extend([t_name] * count)
                    
                    # 再次检查长度
                    if len(types_list) != actual_count:
                        logger.warning(f"警告: 区域 {region_name} 重新计算后的问题类型数量 ({len(types_list)}) 仍与实际行数 ({actual_count}) 不一致！")
                        logger.warning("将使用随机方式填充剩余位置")
                        
                        # 计算差异
                        diff = actual_count - len(types_list)
                        if diff > 0:
                            # 随机选择问题类型填充
                            available_types = list(type_config.keys())
                            additional_types = random.choices(available_types, k=diff)
                            types_list.extend(additional_types)
                        elif diff < 0:
                            # 截断多余的类型
                            types_list = types_list[:actual_count]
                
                # 打乱顺序
                random.shuffle(types_list)
                
                # 填回 DataFrame
                df.loc[group_indices, '问题类型'] = types_list
            
            else:
                # 如果配置里没有这个区域，保持默认值
                logger.info(f"提示: 区域 '{region_name}' 未在问题类型配置中找到，使用默认值。")

        # 重置索引
        df = df.reset_index(drop=True)
        
        # --- 需求4：问题与解决联动 (严格/宽松模式) ---
        logger.info(f"正在生成问题与解决 (严格模式: {STRICT_MATCH})...")
        
        problems = []
        solutions = []
        
        # 预处理：将配置字典转换为便于随机选取的结构
        # strict_pool: {'类型1': [('P1', 'S1'), ('P2', 'S2')]}
        # loose_pool: {'类型1': {'p': ['P1', 'P2'], 's': ['S1', 'S2']}}
        loose_pool = {}
        for t, pairs in PROBLEM_SOLUTION_CONFIG.items():
            if pairs:
                p_list = [p[0] for p in pairs]
                s_list = [p[1] for p in pairs]
                loose_pool[t] = {'p': p_list, 's': s_list}
        
        for idx, row in df.iterrows():
            p_type = row['问题类型']
            
            # 获取该类型下的可选池
            # 如果类型不在配置中，尝试用默认
            if p_type not in PROBLEM_SOLUTION_CONFIG:
                p_type = '默认类型'
                
            if p_type in PROBLEM_SOLUTION_CONFIG and PROBLEM_SOLUTION_CONFIG[p_type]:
                pairs = PROBLEM_SOLUTION_CONFIG[p_type]
                
                if STRICT_MATCH:
                    # 严格模式：随机选一个 Pair，必须成对
                    # random.choice 随机选一个元组
                    selected_pair = random.choice(pairs)
                    problems.append(selected_pair[0])
                    solutions.append(selected_pair[1])
                else:
                    # 宽松模式：问题和解决分别独立随机选取
                    # 只要是该类型下的即可
                    pool = loose_pool[p_type]
                    problems.append(random.choice(pool['p']))
                    solutions.append(random.choice(pool['s']))
            else:
                # 实在找不到配置，填空或者默认
                problems.append('未配置问题')
                solutions.append('未配置解决')
                
        df['问题'] = problems
        df['解决'] = solutions

        # --- 最终数据一致性校验 (需求5) ---
        logger.info("正在进行最终数据校验...")
        
        # 只保留需要的5列
        target_columns = ['时间', '区域', '问题', '解决', '问题类型']
        
        # 检查列是否存在，如果不存在则创建空列 (理论上上面都已经生成了)
        for col in target_columns:
            if col not in df.columns:
                df[col] = ''
                
        # 筛选列
        df = df[target_columns]
        
        # 检查每一列的非空数量
        column_counts = df.count()
        logger.info("各列数据量统计:")
        logger.info("%s", column_counts)
        
        if column_counts.min() == column_counts.max() == original_count:
             logger.info(f"校验通过！所有 5 列的数据量均一致且为 {original_count}。")
        else:
             logger.warning("警告：数据量存在不一致！")
             # 如果有空值，填充默认值以保证一致性
             df = df.fillna('未知')
             logger.info("已自动填充空值以修复一致性。")

        # 输出预览
        logger.info("结果预览 (前10行):")
        logger.info("\n%s", df.head(10).to_string())
        
        # 保存结果
        output_path = file_path.replace('.xlsx', '_已排期.xlsx')
        df.to_excel(output_path, index=False)
        logger.info(f"成功! 结果已保存至: {output_path}")
        
    except Exception:
        logger.exception("处理过程中发生错误")

def create_new_excel(output_path):
    """
    直接创建新的Excel文件，不依赖现有文件
    """
    try:
        # 确定总行数
        if TARGET_ROW_COUNT > 0:
            logger.info(f"检测到目标行数设置: {TARGET_ROW_COUNT}")
            original_count = TARGET_ROW_COUNT
        else:
            # 如果没有设置目标行数，使用区域配置的总数
            total_region_count = sum(REGION_CONFIG.values())
            original_count = total_region_count
            logger.info(f"使用区域配置总数作为行数: {original_count}")
        
        # 创建空的DataFrame
        df = pd.DataFrame(index=range(original_count))
        logger.info(f"已创建空数据表，将生成 {original_count} 行数据")
        
        # 生成有效工作日池
        logger.info("正在计算有效工作日...")
        valid_workdays = get_valid_workdays(START_DATE, END_DATE, SPECIAL_WORK_DAYS, SPECIAL_OFF_DAYS)
        logger.info(f"在 {START_DATE} 到 {END_DATE} 期间，共有 {len(valid_workdays)} 个有效工作日")
        
        if not valid_workdays:
            logger.error("错误: 没有找到有效的工作日，请检查日期范围和调休设置")
            return

        # --- 需求1：时间分配 ---
        logger.info("正在分配日期...")
        random_dates = np.random.choice(valid_workdays, size=original_count)
        df['时间'] = random_dates
        
        # 按时间排序 (从小到大)
        df = df.sort_values('时间')
        
        # --- 需求2：区域分配 ---
        # 严格按照配置的数量生成区域
        logger.info("正在分配区域 (严格按照配置数量生成)...")
        
        regions_list = []
        for name, count in REGION_CONFIG.items():
            regions_list.extend([name] * count)
        
        # 验证区域总数与目标行数是否一致
        if len(regions_list) != original_count:
            logger.warning(f"警告: 配置的区域总数 ({len(regions_list)}) 与目标行数 ({original_count}) 不一致！")
            logger.warning(f"将使用配置的区域总数 ({len(regions_list)}) 作为实际行数")
            original_count = len(regions_list)
            df = pd.DataFrame(index=range(original_count))
        
        # 打乱顺序
        np.random.shuffle(regions_list)
        df['区域'] = regions_list
        
        # --- 需求3：区域与问题类型联动 (精确数量控制) ---
        logger.info("正在根据区域分配问题类型 (严格按照配置数量)...")
        
        # 初始化 '问题类型' 列，默认为 DEFAULT_PROBLEM_TYPE
        df['问题类型'] = DEFAULT_PROBLEM_TYPE
        
        # 按区域分组处理
        # df.groupby('区域') 会将数据按区域拆分
        for region_name, group_indices in df.groupby('区域').groups.items():
            # 获取该区域在当前Excel中的实际行数
            actual_count = len(group_indices)
            
            if actual_count == 0:
                continue
                
            # 获取该区域的配置
            if region_name in REGION_PROBLEM_CONFIG:
                type_config = REGION_PROBLEM_CONFIG[region_name]
                
                # 计算配置中的总数
                config_total = sum(type_config.values())
                
                if config_total == 0:
                    continue
                
                # 检查配置总数与实际行数是否一致
                if config_total != actual_count:
                    logger.warning(f"警告: 区域 {region_name} 的问题类型配置总数 ({config_total}) 与实际行数 ({actual_count}) 不一致！")
                    logger.warning(f"将按照配置的问题类型比例重新计算数量")
                
                # 生成问题类型列表
                types_list = []
                for t_name, t_val in type_config.items():
                    if t_val > 0:  # 只处理数量大于0的类型
                        types_list.extend([t_name] * t_val)
                
                # 检查生成的类型列表长度是否与实际行数一致
                if len(types_list) != actual_count:
                    logger.warning(f"警告: 区域 {region_name} 生成的问题类型数量 ({len(types_list)}) 与实际行数 ({actual_count}) 不一致！")
                    
                    # 重新计算问题类型数量
                    types_list = []
                    for t_name, t_val in type_config.items():
                        if config_total > 0:
                            # 按比例计算数量
                            count = int(round((t_val / config_total) * actual_count))
                            if count > 0:
                                types_list.extend([t_name] * count)
                    
                    # 再次检查长度
                    if len(types_list) != actual_count:
                        logger.warning(f"警告: 区域 {region_name} 重新计算后的问题类型数量 ({len(types_list)}) 仍与实际行数 ({actual_count}) 不一致！")
                        logger.warning("将使用随机方式填充剩余位置")
                        
                        # 计算差异
                        diff = actual_count - len(types_list)
                        if diff > 0:
                            # 随机选择问题类型填充
                            available_types = list(type_config.keys())
                            additional_types = random.choices(available_types, k=diff)
                            types_list.extend(additional_types)
                        elif diff < 0:
                            # 截断多余的类型
                            types_list = types_list[:actual_count]
                
                # 打乱顺序
                random.shuffle(types_list)
                
                # 填回 DataFrame
                df.loc[group_indices, '问题类型'] = types_list
            
            else:
                # 如果配置里没有这个区域，保持默认值或设为特定值
                logger.info(f"提示: 区域 '{region_name}' 未在问题类型配置中找到，使用默认值。")

        # 重置索引
        df = df.reset_index(drop=True)
        
        # --- 需求4：问题与解决联动 (严格/宽松模式) ---
        logger.info(f"正在生成问题与解决 (严格模式: {STRICT_MATCH})...")
        
        problems = []
        solutions = []
        
        # 预处理：将配置字典转换为便于随机选取的结构
        # strict_pool: {'类型1': [('P1', 'S1'), ('P2', 'S2')]}
        # loose_pool: {'类型1': {'p': ['P1', 'P2'], 's': ['S1', 'S2']}}
        loose_pool = {}
        for t, pairs in PROBLEM_SOLUTION_CONFIG.items():
            if pairs:
                p_list = [p[0] for p in pairs]
                s_list = [p[1] for p in pairs]
                loose_pool[t] = {'p': p_list, 's': s_list}
        
        for idx, row in df.iterrows():
            p_type = row['问题类型']
            
            # 获取该类型下的可选池
            # 如果类型不在配置中，尝试用默认
            if p_type not in PROBLEM_SOLUTION_CONFIG:
                p_type = '默认类型'
                
            if p_type in PROBLEM_SOLUTION_CONFIG and PROBLEM_SOLUTION_CONFIG[p_type]:
                pairs = PROBLEM_SOLUTION_CONFIG[p_type]
                
                if STRICT_MATCH:
                    # 严格模式：随机选一个 Pair，必须成对
                    # random.choice 随机选一个元组
                    selected_pair = random.choice(pairs)
                    problems.append(selected_pair[0])
                    solutions.append(selected_pair[1])
                else:
                    # 宽松模式：问题和解决分别独立随机选取
                    # 只要是该类型下的即可
                    pool = loose_pool[p_type]
                    problems.append(random.choice(pool['p']))
                    solutions.append(random.choice(pool['s']))
            else:
                # 实在找不到配置，填空或者默认
                problems.append('未配置问题')
                solutions.append('未配置解决')
                
        df['问题'] = problems
        df['解决'] = solutions

        # --- 最终数据一致性校验 (需求5) ---
        logger.info("正在进行最终数据校验...")
        
        # 只保留需要的5列
        target_columns = ['时间', '区域', '问题', '解决', '问题类型']
        
        # 检查列是否存在，如果不存在则创建空列 (理论上上面都已经生成了)
        for col in target_columns:
            if col not in df.columns:
                df[col] = ''
                
        # 筛选列
        df = df[target_columns]
        
        # 检查每一列的非空数量
        column_counts = df.count()
        logger.info("各列数据量统计:")
        logger.info("%s", column_counts)
        
        if column_counts.min() == column_counts.max() == original_count:
             logger.info(f"校验通过！所有 5 列的数据量均一致且为 {original_count}。")
        else:
             logger.warning("警告：数据量存在不一致！")
             # 如果有空值，填充默认值以保证一致性
             df = df.fillna('未知')
             logger.info("已自动填充空值以修复一致性。")

        # 输出预览
        logger.info("结果预览 (前10行):")
        logger.info("\n%s", df.head(10).to_string())
        
        # 保存结果
        df.to_excel(output_path, index=False)
        logger.info(f"成功! 结果已保存至: {output_path}")
        
    except Exception:
        logger.exception("处理过程中发生错误")


if __name__ == "__main__":
    logger.info("默认目录: %s", DEFAULT_DIR)
    output_path = EXCEL_PATH
    logger.info("开始生成新表格...")
    create_new_excel(output_path)
