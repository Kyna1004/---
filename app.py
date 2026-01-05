import pandas as pd
import os

def process_overall_data(file_path):
    # 1. 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return None

    try:
        # 2. 读取【整体数据】Sheet页
        # 核心逻辑修正：直接加载该Sheet页的数据，不通过日志表计算
        df_overall = pd.read_excel(file_path, sheet_name='整体数据')
        
        # 预处理：去除列名的前后空格（防止 Excel 表头有由于空格导致的读取错误）
        df_overall.columns = df_overall.columns.str.strip()

        # 3. 检查关键列是否存在
        required_col = '加入购物车'
        if required_col not in df_overall.columns:
            print(f"错误：在【整体数据】表中找不到列名为 [{required_col}] 的列，请检查Excel表头。")
            print(f"现有列名：{df_overall.columns.tolist()}")
            return None

        # ==========================================
        # 核心逻辑修正区域 - START
        # ==========================================
        
        # 逻辑：直接取值
        # 不做 sum() 聚合（除非你需要汇总所有行），不做 count() 计数
        # 假设每一行代表一天或一个维度的统计，这里直接保留原始数值
        df_overall['加购次数'] = df_overall['加入购物车']
        
        # 如果数据中有空值(NaN)，填充为0，保证后续计算不出错
        df_overall['加购次数'] = df_overall['加购次数'].fillna(0).astype(int)

        # ==========================================
        # 核心逻辑修正区域 - END
        # ==========================================

        # 4. (可选) 展示结果或进行后续计算
        # 例如：如果你需要计算 加购率 = 加购次数 / 访客数
        if '访客数' in df_overall.columns:
            df_overall['加购率'] = df_overall.apply(
                lambda x: x['加购次数'] / x['访客数'] if x['访客数'] > 0 else 0, axis=1
            )
            # 格式化加购率为百分比
            df_overall['加购率_显示'] = df_overall['加购率'].apply(lambda x: format(x, '.2%'))

        print("数据读取及修正成功！前5行数据如下：")
        print(df_overall[['加入购物车', '加购次数']].head())
        
        return df_overall

    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        return None

# ==========================================
# 执行代码
# ==========================================
file_path = '周期性复盘报告.xlsx'  # 请确保文件名一致
df_result = process_overall_data(file_path)
