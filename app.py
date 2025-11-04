import streamlit as st
import pandas as pd
import tempfile
import os
from collections import defaultdict
import sys
import re
import io

# =============================================================================
# SB 商品集生成函数 (from test SB 集合版.py)
# =============================================================================
def generate_header_from_brand_survey(brand_survey_file, output_file, sheet_name=0):
    try:
        # 读取整个文件，使用第一行作为列名（用于关键词和否定列）
        df_survey = pd.read_excel(brand_survey_file, sheet_name=sheet_name)
        print(f"成功读取文件：{brand_survey_file}，数据形状：{df_survey.shape}")
        print(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        print(f"错误：未找到文件 {brand_survey_file}。请确保文件在同一目录下。")
        return None
    except Exception as e:
        print(f"读取文件时出错：{e}")
        return None

    # 新加：动态区域检测函数
    def find_region_start_end(df, target_theme):
        """扫描A列找到主题行，返回 (header_row, end_row) (0-based索引)"""
        theme_row = None
        next_theme_row = None
        for idx, val in enumerate(df.iloc[:, 0]):  # A列 (index 0)
            if pd.notna(val) and target_theme in str(val).strip():
                theme_row = idx
                break
        if theme_row is None:
            print(f"错误：未找到主题 '{target_theme}' 在A列")
            return None, None
        
        # 找下一个主题
        next_themes = []  # 商品集无下一个主题，到文件末尾
        for idx in range(theme_row + 1, len(df)):
            val = str(df.iloc[idx, 0]).strip()
            if any(nt in val for nt in next_themes):
                next_theme_row = idx
                break
        end_row = next_theme_row - 1 if next_theme_row else len(df) - 1  # 到文件末尾
        header_row = theme_row + 1  # header在主题行下一行
        print(f"找到 '{target_theme}' 区域: 主题行 {theme_row+1}, header行 {header_row+1}, 数据到行 {end_row+1}")
        return header_row, end_row

    # 先找主题行，用于限全局设置范围
    temp_result = find_region_start_end(df_survey, 'SB落地页：商品集')  # target_theme
    if temp_result[0] is None:
        return None
    global_limit = temp_result[0]  # 用header_row限全局设置提取范围
    
    # 提取全局设置：从row2-20的A:B列（标签在A，值在B）
    global_settings = {}
    for i in range(0, min(20, global_limit)):  # 从 iloc[0] (文档 row2) 开始
        if i >= len(df_survey):
            break
        label = str(df_survey.iloc[i, 0]).strip() if pd.notna(df_survey.iloc[i, 0]) else ''
        value = str(df_survey.iloc[i, 1]).strip() if pd.notna(df_survey.iloc[i, 1]) and len(df_survey.columns) > 1 else ''
        print(f"Row {i+1}: label='{label}', value='{value}'")  # 调试打印：检查提取
        
        # 更robust匹配：使用 in 或 startswith，避免空格或变体问题
        if '品牌实体编号' in label:
            global_settings['brand_entity_id'] = value
            print(f"匹配品牌实体编号: {value}")  # 确认匹配
        elif '品牌名称' in label:
            global_settings['brand_name'] = value
        elif '竞价优化' in label:
            global_settings['bidding_optimization'] = value if value else '手动'
        elif '预算类型' in label:
            global_settings['budget_type'] = value if value else '每日'
        elif 'SB广告格式' in label:
            global_settings['ad_format'] = value if value else '商品集'
        elif '创意素材标题' in label:
            global_settings['creative_title'] = value
        elif '落地页 URL' in label:
            global_settings['landing_url'] = value
    
    print(f"全局设置: {global_settings}")
    
    # 动态读取活动区域
    header_row, end_row = temp_result  # 复用
    if header_row is None:
        return None

    # 读取header行作为列名
    header_data = pd.read_excel(brand_survey_file, sheet_name=sheet_name, skiprows=header_row, nrows=1)
    col_names = header_data.iloc[0].tolist()  # 获取列名

    # 读取数据行 (从header下一行到end_row)
    activity_df = pd.DataFrame()
    if end_row > header_row:
        activity_df = pd.read_excel(brand_survey_file, sheet_name=sheet_name, skiprows=header_row + 1, nrows=end_row - header_row)
        activity_df.columns = col_names  # 设置列名
        print(f"活动数据形状: {activity_df.shape}")
        print(f"活动列名: {list(activity_df.columns)}")
    else:
        print("无活动数据行")
        return None
    
    # 先收集 ASIN 数据，使用原始 activity_df（在清理前）
    activity_to_asins = {}
    for idx, row in activity_df.iterrows():
        campaign_name = str(row.iloc[1]).strip() if len(row) > 1 else ''  # B列 '广告活动名称'
        if pd.isna(campaign_name) or not campaign_name or campaign_name.lower() == 'nan':
            continue
        # 收集创意素材 ASIN：从D列（index 3）开始，到F列（index 5）结束，收集非空值，并过滤非 ASIN（如数字）
        creative_asins = []
        start_idx = 3  # D列 (0-based index 3)
        end_idx = 6   # F列结束 (exclusive, so up to index 5)
        for j in range(start_idx, end_idx):
            if j < len(row):
                val = row.iloc[j]
                if pd.notna(val) and str(val).strip():
                    val_str = str(val).strip()
                    # 过滤：假设 ASIN 是字符串，非纯数字/浮点
                    if not val_str.replace('.', '').replace('-', '').isdigit():
                        creative_asins.append(val_str)
        activity_to_asins[campaign_name] = creative_asins
        print(f"活动 {campaign_name} ASIN: {creative_asins}")
    
    # 现在清理Unnamed列：先转str处理NaN列名
    activity_df.columns = activity_df.columns.astype(str)
    activity_df = activity_df.loc[:, ~activity_df.columns.str.contains('^Unnamed')]
    
    unique_campaigns = [name for name in activity_df.iloc[:, 1].dropna() if str(name).strip() and str(name).lower() != 'nan']  # B列
    unique_campaigns = list(dict.fromkeys(unique_campaigns))  # 去重
    print(f"独特活动名称数量: {len(unique_campaigns)}: {unique_campaigns}")
    
    # 活动到值的映射（CPC, 预算, 自定义竞价调整百分比, 自定义图片素材编号, 创意素材 ASIN）
    required_cols = ['CPC', '预算', '自定义竞价调整百分比', '自定义图片素材编号']
    activity_to_values = {}
    for idx, row in activity_df.iterrows():
        campaign_name = str(row.iloc[1]).strip() if len(row) > 1 else ''  # B列
        if pd.isna(campaign_name) or not campaign_name or campaign_name.lower() == 'nan':
            continue
        vals = {}
        # 动态列索引匹配（因为列名可能Unnamed，转str后用in）
        cpc_idx = None
        budget_idx = None
        bid_adjust_idx = None
        image_idx = None
        for j, col in enumerate(activity_df.columns):
            col_str = str(col).strip()
            if 'CPC' in col_str:
                cpc_idx = j
            elif '预算' in col_str:
                budget_idx = j
            elif '自定义竞价调整' in col_str:
                bid_adjust_idx = j
            elif '自定义图片素材编号' in col_str:
                image_idx = j
        vals['CPC'] = row.iloc[cpc_idx] if cpc_idx is not None else ''
        vals['预算'] = row.iloc[budget_idx] if budget_idx is not None else 10
        vals['自定义竞价调整百分比'] = row.iloc[bid_adjust_idx] if bid_adjust_idx is not None else 0
        vals['自定义图片素材编号'] = row.iloc[image_idx] if image_idx is not None else ''
        
        # 使用预收集的 ASIN
        vals['creative_asin_list'] = activity_to_asins.get(campaign_name, [])
        
        # 新增：提取每个活动的 '品牌徽标素材编号' 从 J 列（假设列名为 '品牌徽标素材编号' 或位置 index 9）
        logo_col_name = '品牌徽标素材编号'  # 如果列名不同，调整这里
        logo_value = ''
        if logo_col_name in activity_df.columns:
            logo_value = str(row[logo_col_name]).strip() if pd.notna(row[logo_col_name]) else ''
        else:
            # 备选：用固定位置 J 列 (index 9, 清理后列数可能变，用iloc[8] if 0-based)
            logo_idx = 8  # 假设J是第9列 (0-based 8)
            if len(row) > logo_idx:
                logo_value = str(row.iloc[logo_idx]).strip() if pd.notna(row.iloc[logo_idx]) else ''
        vals['logo_asset_id'] = logo_value
        print(f"活动 {campaign_name} 的品牌徽标素材编号: {logo_value}")  # 调试打印
        
        activity_to_values[campaign_name] = vals
    
    print(f"生成的活动字典（有 {len(activity_to_values)} 个活动）: {list(activity_to_values.keys())}")
    
    # 关键词列：从df_survey的列7:17（与script-C.py一致）
    keyword_columns = df_survey.columns[7:17]
    # 过滤掉Unnamed列，避免活动数据列被误认为是关键词列
    keyword_columns = [col for col in keyword_columns if not str(col).startswith('Unnamed:')]
    print(f"关键词列: {list(keyword_columns)}")
    
    # 检查关键词重复
    duplicates_found = False
    print("\n=== 检查关键词重复 ===")
    for col in keyword_columns:
        if col in df_survey.columns:
            kw_list = [kw for kw in df_survey[col].dropna() if str(kw).strip()]
            if len(kw_list) > len(set(kw_list)):
                duplicates_found = True
                print(f"警告：{col} 列有重复关键词")
    
    # 否定关键词列：从df_survey找（仅用于重复检查）
    neg_exact = []
    neg_phrase = []
    suzhu_extra_neg_exact = []
    suzhu_extra_neg_phrase = []
    neg_asin = []
    neg_brand = []  # 针对“否品牌”列，生成 brand="XXX"
    
    neg_cols = {
        '否定精准': neg_exact,
        '否定词组': neg_phrase,
        '宿主额外否精准': suzhu_extra_neg_exact,
        '宿主额外否词组': suzhu_extra_neg_phrase,
        '否定ASIN': neg_asin,
        '否品牌': neg_brand  # 提取“否品牌”列
    }
    
    for col_name, lst in neg_cols.items():
        col_idx = None
        for idx, col in enumerate(df_survey.columns):
            if col_name in str(col):
                col_idx = idx
                break
        if col_idx is not None:
            if col_name == '否品牌':           
                col_data = [str(int(kw)).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
            else:    
                col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]       
            lst.extend(col_data)
            lst[:] = list(dict.fromkeys(lst))  # 去重
            if len(col_data) > len(set(col_data)):
                duplicates_found = True
                print(f"警告：'{col_name}' 列有重复关键词")

    # 分别处理否定ASIN和否品牌（不合并）
    print(f"否定ASIN列表: {neg_asin}")
    print(f"否品牌列表: {neg_brand}")
    
    if duplicates_found:
        print("\n提示：由于检测到关键词重复，本次不生成表格。请清理重复后重试。")
        return None
    
    print("关键词无重复，继续生成...")
    
    # 识别关键词类别（与script-C.py一致）
    keyword_categories = set()
    for col in keyword_columns:
        col_lower = str(col).lower()
        if '/' in col_lower:
            parts = col_lower.split('/')
            if len(parts) > 0 and parts[0]:
                keyword_categories.add(parts[0].strip())
            if len(parts) > 1 and parts[1]:
                chinese_part = parts[1].split('-')[0].strip() if '-' in parts[1] else parts[1].strip()
                keyword_categories.add(chinese_part)
        else:
            for suffix in ['精准词', '广泛词', '精准', '广泛']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    if prefix:
                        keyword_categories.add(prefix)
                        break
    
    keyword_categories.update(['suzhu', '宿主', 'host', 'case', '包', '对手', 'tape'])  # 与script-C.py一致，移除'xxx'
    print(f"识别到的关键词类别: {keyword_categories}")
    
    # Header列定义（32列）
    columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告活动草稿编号', '广告组合编号', '广告组编号', '关键词编号', '商品投放 ID',
        '广告活动名称', '开始日期', '结束日期', '状态', '预算类型', '预算', '竞价优化', '自定义竞价调整百分比', '竞价',
        '关键词文本', '匹配类型', '拓展商品投放编号', '广告格式', '落地页 URL', '落地页 ASIN', '品牌实体编号',
        '品牌名称', '品牌徽标素材编号', '自定义图片素材编号', '创意素材标题', '创意素材 ASIN', '视频媒体编号', '创意素材类型'
    ]
    
    product = '品牌推广'
    operation = 'Create'
    status = '已启用'
    rows = []
    
    for campaign_name, campaign_values in activity_to_values.items():
        cpc = float(campaign_values.get('CPC', 0)) if campaign_values.get('CPC') else 0
        budget = float(campaign_values.get('预算', 10))
        bid_adjust = float(campaign_values.get('自定义竞价调整百分比', 0))
        image_asset = campaign_values.get('自定义图片素材编号', '')
        creative_asins = campaign_values.get('creative_asin_list', [])
        creative_asin = ', '.join(creative_asins) if creative_asins else ''
        
        # 新增：每个活动的 logo_asset_id 从 campaign_values 获取
        logo_asset_id = campaign_values.get('logo_asset_id', '')
        
        print(f"处理活动: {campaign_name}")
        print(f"  创意 ASIN: {creative_asin}")
        print(f"  品牌实体编号: {global_settings.get('brand_entity_id', '未提取')}")
        print(f"  品牌徽标素材编号: {logo_asset_id}")  # 新增调试打印
        
        campaign_name_normalized = str(campaign_name).lower()
        
        # 检测是否 ASIN 活动
        is_asin = 'asin' in campaign_name_normalized
        
        # 检测匹配类别
        matched_category = None
        for cat in keyword_categories:
            if cat in campaign_name_normalized:
                matched_category = cat
                break
        
        # 检测匹配类型（精准/广泛）
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        match_type = '精准' if is_exact else '广泛' if is_broad else '精准'  # 默认 exact
        
        # 生成广告活动行
        row_campaign = [
            product, '广告活动', operation, campaign_name, '', '', '', '', '',
            campaign_name, '', '', status, global_settings.get('budget_type', '每日'), budget,
            global_settings.get('bidding_optimization', '手动'), bid_adjust, '', '', '',
            '', global_settings.get('ad_format', '商品集'), global_settings.get('landing_url', ''), '',  # 添加 '' for '落地页 ASIN'
            global_settings.get('brand_entity_id', ''), global_settings.get('brand_name', ''),
            logo_asset_id,  # 用每个活动的 logo_asset_id
            image_asset, global_settings.get('creative_title', ''),
            creative_asin, '', ''  # 无视频
        ]
        rows.append(row_campaign)
        
        # 关键词：从df_survey的关键词列拉取（与script-C.py一致的收集逻辑）
        keywords = []
        matched_columns = []
        if matched_category and (is_exact or is_broad):
            for col in keyword_columns:
                if col in df_survey.columns:
                    col_lower = str(col).lower()
                    if matched_category in col_lower:
                        if (is_exact and any(x in col_lower for x in ['精准', 'exact'])) or \
                           (is_broad and any(x in col_lower for x in ['广泛', 'broad'])):
                            matched_columns.append(col)
                            keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
            keywords = list(dict.fromkeys(keywords))
            print(f"  匹配的列: {matched_columns}")
            print(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
        else:
            print("  无匹配的关键词列，关键词为空")
        
        if keywords:
            for kw in keywords:
                row_keyword = [
                    product, '关键词', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', cpc, kw, match_type,
                    '', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_keyword)
        
        # 否定关键词（主要针对广泛，且非ASIN组）
        if not is_asin and matched_category:
            # 新规则：固定列索引（使用实际列名映射）
            col_indices = {
                'Y': df_survey.columns.get_loc('case精准-否精准') if 'case精准-否精准' in df_survey.columns else None,
                'Z': df_survey.columns.get_loc('case精准-否词组') if 'case精准-否词组' in df_survey.columns else None,
                'AC': df_survey.columns.get_loc('case广泛-否精准') if 'case广泛-否精准' in df_survey.columns else None,
                'AD': df_survey.columns.get_loc('case广泛-否词组') if 'case广泛-否词组' in df_survey.columns else None,
                'W': df_survey.columns.get_loc('宿主精准-否精准') if '宿主精准-否精准' in df_survey.columns else None,
                'X': df_survey.columns.get_loc('宿主精准-否词组') if '宿主精准-否词组' in df_survey.columns else None,
                'AA': df_survey.columns.get_loc('宿主广泛-否精准') if '宿主广泛-否精准' in df_survey.columns else None,
                'AB': df_survey.columns.get_loc('宿主广泛-否词组') if '宿主广泛-否词组' in df_survey.columns else None,
            }
            
            # 列名映射，用于日志（使用实际列名）
            col_names = {
                'Y': 'case精准-否精准',
                'Z': 'case精准-否词组',
                'AC': 'case广泛-否精准',
                'AD': 'case广泛-否词组',
                'W': '宿主精准-否精准',
                'X': '宿主精准-否词组',
                'AA': '宿主广泛-否精准',
                'AB': '宿主广泛-否词组'
            }
            
            # 选择列
            selected_cols = []
            if matched_category in ['case', '包']:
                if is_exact:
                    selected_cols = ['Y', 'Z']
                elif is_broad:
                    selected_cols = ['AC', 'AD']
            elif matched_category in ['suzhu', '宿主', 'host']:
                if is_exact:
                    selected_cols = ['W', 'X']
                elif is_broad:
                    selected_cols = ['AA', 'AB']
            
            # 收集数据，按类型分组，使用 defaultdict(list) 跟踪来源
            neg_data_sources = {
                '否定精准匹配': defaultdict(list),  # kw -> [col_keys]
                '否定词组': defaultdict(list)
            }
            for col_key in selected_cols:
                if col_indices.get(col_key) is not None:
                    col_idx = col_indices[col_key]
                    col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
                    col_data = list(dict.fromkeys(col_data))  # 列内去重
                    m_type = '否定精准匹配' if col_key in ['Y', 'AC', 'W', 'AA'] else '否定词组'
                    for kw in col_data:
                        neg_data_sources[m_type][kw].append(col_key)
            
            # 检查重复：如果 kw 有多个来源
            duplicates_detected = False
            for m_type, kw_sources in neg_data_sources.items():
                for kw, sources in kw_sources.items():
                    if len(sources) > 1:
                        duplicates_detected = True
                        source_names = [col_names.get(s, s) for s in sources]
                        print(f"\n=== 检测到重复否定关键词 ===")
                        print(f"活动: {campaign_name}")
                        print(f"类型: {m_type}")
                        print(f"重复关键词: '{kw}'")
                        print(f"来源列: {', '.join(source_names)}")
                        print(f"原因: 该关键词在多个否定列中出现，导致生成重复行。请检查 survey 文件的这些列并清理重复值。")
                        print("暂停生成 header 表。")
                        return None  # 暂停生成
            
            print("\n=== 重复检测完成（无重复）===")
            
            # 生成行：使用去重后的列表
            for m_type, kw_sources in neg_data_sources.items():
                kws = list(kw_sources.keys())  # 已去重
                if kws:
                    print(f"  {m_type} 否定关键词数量: {len(kws)}")
                for kw in kws:
                    row_neg = [
                        product, '否定关键词', operation, campaign_name, '', '', campaign_name, '', '',
                        campaign_name, '', '', status, '', '', '', '', '', kw, m_type,
                        '', '', '', '', '', '', '', '', '', '', ''
                    ]
                    rows.append(row_neg)
        
        # 商品定向（ASIN）：从df_survey的ASIN列（列名匹配活动名称）
        asin_targets = []
        if is_asin:
            # 精确匹配列名
            for col in df_survey.columns:
                if str(col).strip() == str(campaign_name):
                    col_idx = df_survey.columns.get_loc(col)
                    if col_idx is not None:
                        asin_targets = [str(asin).strip() for asin in df_survey.iloc[:, col_idx].dropna() if str(asin).strip()]
                        asin_targets = list(dict.fromkeys(asin_targets))
                        print(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
                        break
        
        if asin_targets:
            for asin in asin_targets:
                row_asin = [
                    product, '商品定向', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', cpc, '', '',
                    f'asin="{asin}"', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_asin)
        
        # 否定商品定向（全局，针对所有ASIN活动）
        if is_asin:
            # 否定ASIN：使用 asin="XXX"
            for asin in neg_asin:
                row_neg_asin = [
                    product, '否定商品定向', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', '', '', '',
                    f'asin="{asin}"', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_neg_asin)
            
            # 否品牌：使用 brand="XXX"
            for brand in neg_brand:
                row_neg_brand = [
                    product, '否定商品定向', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', '', '', '',
                    f'brand="{brand}"', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_neg_brand)
    
    df_header = pd.DataFrame(rows, columns=columns)
    try:
        df_header.to_excel(output_file, index=False, engine='openpyxl')
        print(f"生成完成！输出文件：{output_file}，总行数：{len(rows)}")
        return output_file
    except Exception as e:
        print(f"保存出错：{e}")
        return None

# =============================================================================
# SBV 品牌旗舰店生成函数 (from test SBV 品牌旗舰店 集合版.py)
# =============================================================================
def generate_header_for_sbv_brand_store(input_file, output_file, sheet_name='品牌广告'):
    try:
        # Read the entire file, header=0
        df_survey = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
        print(f"成功读取文件：{input_file}，数据形状：{df_survey.shape}")
        print(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        print(f"错误：未找到文件 {input_file}。请确保文件在同一目录下。")
        return None
    except Exception as e:
        print(f"读取文件时出错：{e}")
        return None
    
    # Fill NaN with empty string
    df_survey = df_survey.fillna('')
    
    # 新加：动态区域检测函数
    def find_region_start_end(df, target_theme):
        """扫描A列找到主题行，返回 (header_row, end_row) (0-based索引)"""
        theme_row = None
        next_theme_row = None
        for idx, val in enumerate(df.iloc[:, 0]):  # A列 (index 0)
            if pd.notna(val) and target_theme in str(val).strip():
                theme_row = idx
                break
        if theme_row is None:
            print(f"错误：未找到主题 '{target_theme}' 在A列")
            return None, None
        
        # 找下一个主题（顺序：详情页 → 旗舰店 → 商品集）
        next_themes = ["SBV落地页：品牌旗舰店", "落地页：商品集"]  # 从当前开始找下一个
        for idx in range(theme_row + 1, len(df)):
            val = str(df.iloc[idx, 0]).strip()
            if any(nt in val for nt in next_themes if nt != target_theme):
                next_theme_row = idx
                break
        end_row = next_theme_row - 1 if next_theme_row else len(df) - 1  # 到文件末尾
        header_row = theme_row + 1  # header在主题行下一行
        print(f"找到 '{target_theme}' 区域: 主题行 {theme_row+1}, header行 {header_row+1}, 数据到行 {end_row+1}")
        return header_row, end_row

    # 先找主题行，用于限全局设置范围
    temp_result = find_region_start_end(df_survey, 'SBV落地页：品牌旗舰店')
    if temp_result[0] is None:
        return None
    global_limit = temp_result[0]  # 用 [0] 是 header_row，即主题前

    # Extract global settings: from rows 0-20, column A (0) labels, B (1) values
    global_settings = {}
    for i in range(0, min(20, global_limit)):
        if i >= len(df_survey):
            break
        label = str(df_survey.iloc[i, 0]).strip() if pd.notna(df_survey.iloc[i, 0]) else ''
        value = str(df_survey.iloc[i, 1]).strip() if pd.notna(df_survey.iloc[i, 1]) and len(df_survey.columns) > 1 else ''
        print(f"Row {i+1}: label='{label}', value='{value}'")
        
        # Robust matching similar to test SB.py
        if '品牌实体编号' in label or 'ENTITY' in label.upper():
            global_settings['entity_id'] = value
        elif '品牌名称' in label:
            global_settings['brand_name'] = value
        elif '预算类型' in label:
            global_settings['budget_type'] = value if value else '每日'
        elif '创意素材标题' in label:
            global_settings['creative_title'] = value
        elif '落地页 URL' in label:
            global_settings['landing_url'] = value
    
    print(f"全局设置: {global_settings}")
    
    # 动态读取活动区域（用已算的）
    header_row, end_row = temp_result  # 复用
    if header_row is None:
        return None

    # 读取header行作为列名
    header_data = pd.read_excel(input_file, sheet_name=sheet_name, skiprows=header_row, nrows=1)
    col_names = header_data.iloc[0].tolist()  # 获取列名
    
    # 读取数据行 (从header下一行到end_row)
    activity_df = pd.DataFrame()
    if end_row > header_row:
        activity_df = pd.read_excel(input_file, sheet_name=sheet_name, skiprows=header_row + 1, nrows=end_row - header_row)
        activity_df.columns = col_names  # 设置列名
        print(f"活动数据形状: {activity_df.shape}")
        print(f"活动列名: {list(activity_df.columns)}")
    else:
        print("无活动数据行")
        return None

    # 用activity_df构建activity_rows列表
    activity_rows = []
    for idx, row in activity_df.iterrows():
        campaign_name = str(row.iloc[1]).strip() if len(row) > 1 else ''  # B列
        if campaign_name and campaign_name.lower() != 'nan' and not any(global_key in campaign_name for global_key in ['品牌实体编号', '品牌名称', '预算类型', '创意素材标题', '落地页 URL']):
            # Assume this is activity row
            activity_rows.append({
                'index': idx,
                'campaign_name': campaign_name,
                'cpc': row.iloc[2] if len(row) > 2 else '',
                'asin1': str(row.iloc[3]) if len(row) > 3 else '',
                'asin2': str(row.iloc[4]) if len(row) > 4 else '',
                'asin3': str(row.iloc[5]) if len(row) > 5 else '',
                'budget': row.iloc[6] if len(row) > 6 else 10,
                'video_asset': row.iloc[8] if len(row) > 8 else '',
                'logo_asset': row.iloc[9] if len(row) > 9 else ''
            })

    # 加填充 NaN
    activity_df = activity_df.fillna('')
    
    print(f"Found {len(activity_rows)} activity rows: {[r['campaign_name'] for r in activity_rows]}")
    
    # Keyword columns: from header row (iloc[0]), but dynamic like test SB.py
    header_row_full = df_survey.iloc[0].tolist()
    keyword_columns = [col for col in header_row_full if isinstance(col, str) and ('精准词' in col or '广泛词' in col or '否' in col)]
    print(f"关键词相关列: {keyword_columns}")
    
    # Identify keyword categories like in test SB.py
    keyword_categories = set()
    for col in keyword_columns:
        col_lower = str(col).lower()
        if '/' in col_lower:
            parts = col_lower.split('/')
            if len(parts) > 0 and parts[0]:
                keyword_categories.add(parts[0].strip())
            if len(parts) > 1 and parts[1]:
                chinese_part = parts[1].split('-')[0].strip() if '-' in parts[1] else parts[1].strip()
                keyword_categories.add(chinese_part)
        else:
            for suffix in ['精准词', '广泛词', '精准', '广泛']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    if prefix:
                        keyword_categories.add(prefix)
                        break
    keyword_categories.update(['suzhu', '宿主', 'host', 'case', '包', '对手', 'tape'])
    print(f"识别到的关键词类别: {keyword_categories}")
    
    # Negative keywords extraction: map to specific columns like test SB.py
    # Col indices mapping
    col_indices = {
        'W': df_survey.columns.get_loc('宿主精准-否精准') if '宿主精准-否精准' in df_survey.columns else None,
        'X': df_survey.columns.get_loc('宿主精准-否词组') if '宿主精准-否词组' in df_survey.columns else None,
        'AA': df_survey.columns.get_loc('宿主广泛-否精准') if '宿主广泛-否精准' in df_survey.columns else None,
        'AB': df_survey.columns.get_loc('宿主广泛-否词组') if '宿主广泛-否词组' in df_survey.columns else None,
        'Y': df_survey.columns.get_loc('case精准-否精准') if 'case精准-否精准' in df_survey.columns else None,
        'Z': df_survey.columns.get_loc('case精准-否词组') if 'case精准-否词组' in df_survey.columns else None,
        'AC': df_survey.columns.get_loc('case广泛-否精准') if 'case广泛-否精准' in df_survey.columns else None,
        'AD': df_survey.columns.get_loc('case广泛-否词组') if 'case广泛-否词组' in df_survey.columns else None,
    }
    
    # Col names for logging
    col_names_dict = {
        'W': '宿主精准-否精准',
        'X': '宿主精准-否词组',
        'AA': '宿主广泛-否精准',
        'AB': '宿主广泛-否词组',
        'Y': 'case精准-否精准',
        'Z': 'case精准-否词组',
        'AC': 'case广泛-否精准',
        'AD': 'case广泛-否词组'
    }
    
    # Extract neg_asin and neg_brand from specific columns
    neg_asin = []
    neg_brand = []
    neg_asin_col = None
    neg_brand_col = None
    for col_idx, col_name in enumerate(df_survey.columns):
        if '否定asin' in str(col_name).lower():
            neg_asin_col = col_idx
        elif '否品牌' in str(col_name).lower():
            neg_brand_col = col_idx
    if neg_asin_col is not None:
        neg_asin = [str(x).strip() for x in df_survey.iloc[:, neg_asin_col].dropna() if str(x).strip()]
        neg_asin = list(dict.fromkeys(neg_asin))
    if neg_brand_col is not None:
        neg_brand = [str(int(x)).strip() for x in df_survey.iloc[:, neg_brand_col].dropna() if str(x).strip()]
        neg_brand = list(dict.fromkeys(neg_brand))
    print(f"否定ASIN: {neg_asin}")
    print(f"否品牌: {neg_brand}")
    
    # Output columns from original test.py (26 columns)
    output_columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告编号', 
        '广告活动名称', '广告组名称', '广告名称', '状态', '品牌实体编号', 
        '预算类型', '预算', '商品位置', '竞价', '关键词文本', '匹配类型', '拓展商品投放编号', 
        '落地页 URL', '落地页类型', '品牌名称', '同意翻译', '品牌徽标素材编号', 
        '创意素材标题', '创意素材 ASIN', '视频素材编号'
    ]
    
    product = '品牌推广'
    operation = 'Create'
    status = '已启用'
    landing_type = '品牌旗舰店'  # Fixed for SBV
    
    rows = []
    
    default_bid = 0.6
    
    for activity in activity_rows:
        campaign_name = activity['campaign_name']
        print(f"处理活动: {campaign_name}")
        
        cpc = float(activity['cpc']) if pd.notna(activity['cpc']) and activity['cpc'] != '' else default_bid
        budget = float(activity['budget']) if pd.notna(activity['budget']) and activity['budget'] != '' else 10
        asins = [asin for asin in [activity['asin1'], activity['asin2'], activity['asin3']] if asin.strip()]
        asins_str = ', '.join(asins)
        video_asset = activity['video_asset']
        logo_asset = activity['logo_asset']
        
        campaign_name_normalized = str(campaign_name).lower()
        
        # Detect category and match type like test SB.py
        matched_category = None
        for cat in keyword_categories:
            if cat in campaign_name_normalized:
                matched_category = cat
                break
        
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        match_type = '精准' if is_exact else '广泛' if is_broad else '精准'  # Default exact/精准
        
        is_asin = 'asin' in campaign_name_normalized
        
        # Generate rows similar to original, but dynamic keywords
        entity_id = global_settings.get('entity_id', '')
        brand_name = global_settings.get('brand_name', '')
        budget_type = global_settings.get('budget_type', '每日')
        creative_title = global_settings.get('creative_title', '')
        landing_url = global_settings.get('landing_url', '')
        
        # Row1: 广告活动 - set specific columns empty (S=18:落地页 URL, U=21:同意翻译, V=22:品牌徽标, X=24:创意 ASIN)
        row1 = [product, '广告活动', operation, campaign_name, '', '', campaign_name, '', '', status, 
                entity_id, budget_type, budget, '在亚马逊上出售', '', '', '', '', '', '', '', '', '', '', '']
        rows.append(row1)
        
        # Row2: 广告组
        row2 = [product, '广告组', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
        rows.append(row2)
        
        # Row3: 品牌视频广告
        row3 = [product, '品牌视频广告', operation, campaign_name, campaign_name, campaign_name, '', '', campaign_name, status, 
                '', '', '', '', '', '', '', '', landing_url, landing_type, brand_name, 'False', logo_asset, creative_title, asins_str, video_asset]
        rows.append(row3)
        
        # Keywords: fixed column match (skip if ASIN)
        if not is_asin:
            keywords = []
            matched_columns = []
            keyword_col_idx = None
            if matched_category in ['suzhu', '宿主', 'host']:
                if is_exact:
                    keyword_col_idx = 11  # L列: suzhu/宿主/host-精准词
                elif is_broad:
                    keyword_col_idx = 12  # M列: suzhu/宿主/host-广泛词
            elif matched_category == 'case':
                if is_exact:
                    keyword_col_idx = 13  # N列: case/包-精准词
                elif is_broad:
                    keyword_col_idx = 14  # O列: case/包-广泛词
            
            if keyword_col_idx is not None and keyword_col_idx < len(df_survey.columns):
                col_data = [str(kw).strip() for kw in df_survey.iloc[:, keyword_col_idx].dropna() if str(kw).strip()]
                keywords = list(dict.fromkeys(col_data))
                matched_columns = [df_survey.columns[keyword_col_idx]]
                print(f"  匹配的列: {matched_columns}")
                print(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
            else:
                # Fallback to original hardcode style
                try:
                    positive_kw_index = header_row_full.index('case/包-精准词')
                    keywords = [header_row_full[positive_kw_index]]  # Single kw?
                except:
                    keywords = []
            
            if keywords:
                for kw in keywords:
                    row_keyword = [product, '关键词', operation, campaign_name, campaign_name, '', '', '', '', status, 
                                   '', '', '', '', cpc, kw, match_type, '', '', '', '', '', '', '', '', '']
                    rows.append(row_keyword)
            else:
                # Original single row
                row4 = [product, '关键词', operation, campaign_name, campaign_name, '', '', '', '', status, 
                        '', '', '', '', cpc, 'default_kw', match_type, '', '', '', '', '', '', '', '', '']
                rows.append(row4)
            
            # Negative keywords: dynamic like test SB.py, with specific column selection
            if matched_category:
                # Select columns based on category and type
                selected_cols = []
                if matched_category in ['suzhu', '宿主', 'host']:
                    if is_exact:
                        selected_cols = ['W', 'X']
                    elif is_broad:
                        selected_cols = ['AA', 'AB']
                elif matched_category in ['case', '包']:
                    if is_exact:
                        selected_cols = ['Y', 'Z']
                    elif is_broad:
                        selected_cols = ['AC', 'AD']
                
                # Collect data, track sources for duplicates
                neg_data_sources = {
                    '否定精准匹配': defaultdict(list),  # kw -> [col_keys]
                    '否定词组': defaultdict(list)
                }
                for col_key in selected_cols:
                    if col_indices.get(col_key) is not None:
                        col_idx = col_indices[col_key]
                        col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
                        col_data = list(dict.fromkeys(col_data))  # column dedup
                        m_type = '否定精准匹配' if col_key in ['W', 'AA', 'Y', 'AC'] else '否定词组'
                        for kw in col_data:
                            neg_data_sources[m_type][kw].append(col_key)
                
                # Check duplicates: kw with multiple sources
                duplicates_detected = False
                for m_type, kw_sources in neg_data_sources.items():
                    for kw, sources in kw_sources.items():
                        if len(sources) > 1:
                            duplicates_detected = True
                            source_names = [col_names_dict.get(s, s) for s in sources]
                            print(f"\n=== 检测到重复否定关键词 ===")
                            print(f"活动: {campaign_name}")
                            print(f"类型: {m_type}")
                            print(f"重复关键词: '{kw}'")
                            print(f"来源列: {', '.join(source_names)}")
                            print(f"原因: 该关键词在多个否定列中出现，导致生成重复行。请检查 survey 文件的这些列并清理重复值。")
                            print("暂停生成 header 表。")
                            return None  # Pause generation
                
                print("\n=== 重复检测完成（无重复）===")
                
                # Generate rows: deduped kws
                for m_type, kw_sources in neg_data_sources.items():
                    kws = list(kw_sources.keys())
                    if kws:
                        print(f"  {m_type} 否定关键词数量: {len(kws)}")
                    for kw in kws:
                        row_neg = [product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '', status, 
                                   '', '', '', '', '', kw, m_type, '', '', '', '', '', '', '', '', '']
                        rows.append(row_neg)
        
        # ASIN group: generate 商品定向 and 否定商品定向
        if is_asin:
            # 商品定向: exact column match to campaign_name
            asin_targets = []
            for col in df_survey.columns:
                if str(col).strip() == str(campaign_name):
                    col_idx = df_survey.columns.get_loc(col)
                    if col_idx is not None:
                        asin_targets = [str(asin).strip() for asin in df_survey.iloc[:, col_idx].dropna() if str(asin).strip()]
                        asin_targets = list(dict.fromkeys(asin_targets))
                        print(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
                        break
            
            if asin_targets:
                for asin in asin_targets:
                    row_product_target = [product, '商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                          '', '', '', '', cpc, '', '', f'asin="{asin}"', '', '', '', '', '', '', '', '']
                    rows.append(row_product_target)
            
            # 否定商品定向: from global neg_asin and neg_brand
            for neg in neg_asin:
                row_neg_product = [product, '否定商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                   '', '', '', '', '', '', '', f'asin="{neg}"', '', '', '', '', '', '', '', '']
                rows.append(row_neg_product)
            
            for negb in neg_brand:
                row_neg_brand = [product, '否定商品定向', operation, campaign_name, campaign_name, '', '', campaign_name, '', status, 
                                 '', '', '', '', '', '', '', f'brand="{negb}"', '', '', '', '', '', '', '', '']
                rows.append(row_neg_brand)
    
    # Create DF and save
    df_header = pd.DataFrame(rows, columns=output_columns); df_header = df_header.fillna('')
    try:
        df_header.to_excel(output_file, index=False, engine='openpyxl')
        print(f"生成完成！输出文件：{output_file}，总行数：{len(rows)}")
        return output_file
    except Exception as e:
        print(f"保存出错：{e}")
        return None

# =============================================================================
# SBV 详情页生成函数 (from test SBV 详情页 集合版.py)
# =============================================================================
def generate_header_sbv_detail(brand_survey_file, output_file, sheet_name=0):
    try:
        # 读取整个文件，使用第一行作为列名（用于关键词和否定列）
        df_survey = pd.read_excel(brand_survey_file, sheet_name=sheet_name)
        print(f"成功读取文件：{brand_survey_file}，数据形状：{df_survey.shape}")
        print(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        print(f"错误：未找到文件 {brand_survey_file}。请确保文件在同一目录下。")
        return None
    except Exception as e:
        print(f"读取文件时出错：{e}")
        return None

    # 新加：动态区域检测函数
    def find_region_start_end(df, target_theme):
        """扫描A列找到主题行，返回 (header_row, end_row) (0-based索引)"""
        theme_row = None
        next_theme_row = None
        for idx, val in enumerate(df.iloc[:, 0]):  # A列 (index 0)
            if pd.notna(val) and target_theme in str(val).strip():
                theme_row = idx
                break
        if theme_row is None:
            print(f"错误：未找到主题 '{target_theme}' 在A列")
            return None, None
        
        # 找下一个主题
        next_themes = ["SBV落地页：品牌旗舰店"]  # 详情页的下一个主题
        for idx in range(theme_row + 1, len(df)):
            val = str(df.iloc[idx, 0]).strip()
            if any(nt in val for nt in next_themes):
                next_theme_row = idx
                break
        end_row = next_theme_row - 1 if next_theme_row else len(df) - 1  # 到文件末尾
        header_row = theme_row + 1  # header在主题行下一行
        print(f"找到 '{target_theme}' 区域: 主题行 {theme_row+1}, header行 {header_row+1}, 数据到行 {end_row+1}")
        return header_row, end_row

    # 先找主题行，用于限全局设置范围
    temp_result = find_region_start_end(df_survey, 'SBV落地页：商品详情页')  # target_theme
    if temp_result[0] is None:
        return None
    global_limit = temp_result[0]  # 用header_row限全局设置提取范围
    
    # 提取全局设置：从row2-20的A:B列（标签在A，值在B）
    global_settings = {}
    for i in range(0, min(20, global_limit)):  # 从 iloc[0] (文档 row2) 开始
        if i >= len(df_survey):
            break
        label = str(df_survey.iloc[i, 0]).strip() if pd.notna(df_survey.iloc[i, 0]) else ''
        value = str(df_survey.iloc[i, 1]).strip() if pd.notna(df_survey.iloc[i, 1]) and len(df_survey.columns) > 1 else ''
        print(f"Row {i+1}: label='{label}', value='{value}'")  # 调试打印：检查提取
        
        # 更robust匹配：使用 in 或 startswith，避免空格或变体问题
        if '品牌实体编号' in label:
            global_settings['brand_entity_id'] = value
            print(f"匹配品牌实体编号: {value}")  # 确认匹配
        elif '品牌名称' in label:
            global_settings['brand_name'] = value
        elif '竞价优化' in label:
            global_settings['bidding_optimization'] = value if value else '手动'
        elif '预算类型' in label:
            global_settings['budget_type'] = value if value else '每日'
        elif 'SB广告格式' in label:
            global_settings['ad_format'] = value if value else '商品集'
        elif 'SBV广告格式' in label:  # 添加SBV匹配（日志中有）
            global_settings['ad_format'] = value if value else '视频'
        elif '创意素材标题' in label:
            global_settings['creative_title'] = value
        elif '落地页 URL' in label:
            global_settings['landing_url'] = value
    
    print(f"全局设置: {global_settings}")
    
    # 动态读取活动区域
    header_row, end_row = temp_result  # 复用
    if header_row is None:
        return None

    # 读取header行作为列名
    header_data = pd.read_excel(brand_survey_file, sheet_name=sheet_name, skiprows=header_row, nrows=1)
    col_names = header_data.iloc[0].tolist()  # 获取列名

    # 读取数据行 (从header下一行到end_row)
    activity_df = pd.DataFrame()
    if end_row > header_row:
        activity_df = pd.read_excel(brand_survey_file, sheet_name=sheet_name, skiprows=header_row + 1, nrows=end_row - header_row)
        activity_df.columns = col_names  # 设置列名
        print(f"活动数据形状: {activity_df.shape}")
        print(f"活动列名: {list(activity_df.columns)}")
    else:
        print("无活动数据行")
        return None
    
    # 先收集 ASIN 数据，使用原始 activity_df（在清理前）
    activity_to_asins = {}
    for idx, row in activity_df.iterrows():
        campaign_name = str(row.iloc[1]).strip() if len(row) > 1 else ''  # B列 '广告活动名称'
        if pd.isna(campaign_name) or not campaign_name or campaign_name.lower() == 'nan':
            continue
        # 收集创意素材 ASIN：从D列（index 3）开始，到F列（index 5）结束，收集非空值，并过滤非 ASIN（如数字）
        creative_asins = []
        start_idx = 3  # D列 (0-based index 3)
        end_idx = 6   # F列结束 (exclusive, so up to index 5)
        for j in range(start_idx, end_idx):
            if j < len(row):
                val = row.iloc[j]
                if pd.notna(val) and str(val).strip():
                    val_str = str(val).strip()
                    # 过滤：假设 ASIN 是字符串，非纯数字/浮点
                    if not val_str.replace('.', '').replace('-', '').isdigit():
                        creative_asins.append(val_str)
        activity_to_asins[campaign_name] = creative_asins
        print(f"活动 {campaign_name} ASIN: {creative_asins}")
    
    # 现在清理Unnamed列：先转str处理NaN列名
    activity_df.columns = activity_df.columns.astype(str)
    activity_df = activity_df.loc[:, ~activity_df.columns.str.contains('^Unnamed')]
    
    unique_campaigns = [name for name in activity_df.iloc[:, 1].dropna() if str(name).strip() and str(name).lower() != 'nan']  # B列
    unique_campaigns = list(dict.fromkeys(unique_campaigns))  # 去重
    print(f"独特活动名称数量: {len(unique_campaigns)}: {unique_campaigns}")
    
    # 活动到值的映射（CPC, 预算, 自定义竞价调整百分比, 自定义图片素材编号, 创意素材 ASIN）
    required_cols = ['CPC', '预算', '自定义竞价调整百分比', '自定义图片素材编号']
    activity_to_values = {}
    for idx, row in activity_df.iterrows():
        campaign_name = str(row.iloc[1]).strip() if len(row) > 1 else ''  # B列
        if pd.isna(campaign_name) or not campaign_name:
            continue
        vals = {}
        # 动态列索引匹配（因为列名可能Unnamed，转str后用iloc或find）
        cpc_idx = None
        budget_idx = None
        bid_adjust_idx = None
        image_idx = None
        for j, col in enumerate(activity_df.columns):
            col_str = str(col).strip()
            if 'CPC' in col_str:
                cpc_idx = j
            elif '预算' in col_str:
                budget_idx = j
            elif '自定义竞价调整' in col_str:
                bid_adjust_idx = j
            elif '自定义图片素材编号' in col_str:
                image_idx = j
        vals['CPC'] = row.iloc[cpc_idx] if cpc_idx is not None else ''
        vals['预算'] = row.iloc[budget_idx] if budget_idx is not None else 10
        vals['自定义竞价调整百分比'] = row.iloc[bid_adjust_idx] if bid_adjust_idx is not None else 0
        vals['自定义图片素材编号'] = row.iloc[image_idx] if image_idx is not None else ''
        
        # 使用预收集的 ASIN
        vals['creative_asin_list'] = activity_to_asins.get(campaign_name, [])
        
        # 新增：提取每个活动的 '品牌徽标素材编号' 从 J 列（假设列名为 '品牌徽标素材编号' 或位置 index 9）
        logo_col_name = '品牌徽标素材编号'  # 如果列名不同，调整这里
        logo_value = ''
        if logo_col_name in activity_df.columns:
            logo_value = str(row[logo_col_name]).strip() if pd.notna(row[logo_col_name]) else ''
        else:
            # 备选：用固定位置 J 列 (index 9, 清理后列数可能变，用iloc[8] if 0-based)
            logo_idx = 8  # 假设J是第9列 (0-based 8)
            if len(row) > logo_idx:
                logo_value = str(row.iloc[logo_idx]).strip() if pd.notna(row.iloc[logo_idx]) else ''
        vals['logo_asset_id'] = logo_value
        print(f"活动 {campaign_name} 的品牌徽标素材编号: {logo_value}")  # 调试打印
        
        # 新增：提取每个活动的 '视频媒体编号'
        video_col_name = '视频媒体编号'
        video_value = ''
        if video_col_name in activity_df.columns:
            video_value = str(row[video_col_name]).strip() if pd.notna(row[video_col_name]) else ''
        else:
            # 备选：假设I列 (index 8)
            video_idx = 7  # 调整根据列
            if len(row) > video_idx:
                video_value = str(row.iloc[video_idx]).strip() if pd.notna(row.iloc[video_idx]) else ''
        vals['video_media_id'] = video_value
        print(f"活动 {campaign_name} 的视频媒体编号: {video_value}")  # 新增调试打印
        
        activity_to_values[campaign_name] = vals
    
    print(f"生成的活动字典（有 {len(activity_to_values)} 个活动）: {list(activity_to_values.keys())}")
    
    # 关键词列：从df_survey的列7:17（与script-C.py一致）
    keyword_columns = df_survey.columns[7:17]
    # 过滤掉Unnamed列，避免活动数据列被误认为是关键词列
    keyword_columns = [col for col in keyword_columns if not str(col).startswith('Unnamed:')]
    print(f"关键词列: {list(keyword_columns)}")
    
    # 检查关键词重复
    duplicates_found = False
    print("\n=== 检查关键词重复 ===")
    for col in keyword_columns:
        if col in df_survey.columns:
            kw_list = [kw for kw in df_survey[col].dropna() if str(kw).strip()]
            if len(kw_list) > len(set(kw_list)):
                duplicates_found = True
                print(f"警告：{col} 列有重复关键词")
    
    # 否定关键词列：从df_survey找（仅用于重复检查）
    neg_exact = []
    neg_phrase = []
    suzhu_extra_neg_exact = []
    suzhu_extra_neg_phrase = []
    neg_asin = []
    neg_brand = []  # 针对“否品牌”列，生成 brand="XXX"
    
    neg_cols = {
        '否定精准': neg_exact,
        '否定词组': neg_phrase,
        '宿主额外否精准': suzhu_extra_neg_exact,
        '宿主额外否词组': suzhu_extra_neg_phrase,
        '否定ASIN': neg_asin,
        '否品牌': neg_brand  # 提取“否品牌”列
    }
    
    for col_name, lst in neg_cols.items():
        col_idx = None
        for idx, col in enumerate(df_survey.columns):
            if col_name in str(col):
                col_idx = idx
                break
        if col_idx is not None:
            if col_name == '否品牌':
                col_data = [str(int(kw)).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
            else:
                col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
            lst.extend(col_data)
            lst[:] = list(dict.fromkeys(lst))  # 去重
            if len(col_data) > len(set(col_data)):
                duplicates_found = True
                print(f"警告：'{col_name}' 列有重复关键词")
    
    # 分别处理否定ASIN和否品牌（不合并）
    print(f"否定ASIN列表: {neg_asin}")
    print(f"否品牌列表: {neg_brand}")
    
    if duplicates_found:
        print("\n提示：由于检测到关键词重复，本次不生成表格。请清理重复后重试。")
        return None
    
    print("关键词无重复，继续生成...")
    
    # 识别关键词类别（与script-C.py一致）
    keyword_categories = set()
    for col in keyword_columns:
        col_lower = str(col).lower()
        if '/' in col_lower:
            parts = col_lower.split('/')
            if len(parts) > 0 and parts[0]:
                keyword_categories.add(parts[0].strip())
            if len(parts) > 1 and parts[1]:
                chinese_part = parts[1].split('-')[0].strip() if '-' in parts[1] else parts[1].strip()
                keyword_categories.add(chinese_part)
        else:
            for suffix in ['精准词', '广泛词', '精准', '广泛']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    if prefix:
                        keyword_categories.add(prefix)
                        break
    
    keyword_categories.update(['suzhu', '宿主', 'host', 'case', '包', '对手', 'tape'])  # 与script-C.py一致，移除'xxx'
    print(f"识别到的关键词类别: {keyword_categories}")
    
    # Header列定义（32列）
    columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告活动草稿编号', '广告组合编号', '广告组编号', '关键词编号', '商品投放 ID',
        '广告活动名称', '开始日期', '结束日期', '状态', '预算类型', '预算', '竞价优化', '自定义竞价调整百分比', '竞价',
        '关键词文本', '匹配类型', '拓展商品投放编号', '广告格式', '落地页 URL', '落地页 ASIN', '品牌实体编号',
        '品牌名称', '品牌徽标素材编号', '自定义图片素材编号', '创意素材标题', '创意素材 ASIN', '视频媒体编号', '创意素材类型'
    ]
    
    product = '品牌推广'
    operation = 'Create'
    status = '已启用'
    rows = []
    
    for campaign_name, campaign_values in activity_to_values.items():
        cpc = float(campaign_values.get('CPC', 0)) if campaign_values.get('CPC') else 0
        budget = float(campaign_values.get('预算', 10))
        bid_adjust = float(campaign_values.get('自定义竞价调整百分比', 0))
        image_asset = campaign_values.get('自定义图片素材编号', '')
        creative_asins = campaign_values.get('creative_asin_list', [])
        creative_asin = ', '.join(creative_asins) if creative_asins else ''
        
        # 新增：每个活动的 logo_asset_id 从 campaign_values 获取
        logo_asset_id = campaign_values.get('logo_asset_id', '')
        
        # 新增：每个活动的 video_media_id 从 campaign_values 获取
        video_media_id = campaign_values.get('video_media_id', '')
        
        print(f"处理活动: {campaign_name}")
        print(f"  创意 ASIN: {creative_asin}")
        print(f"  品牌实体编号: {global_settings.get('brand_entity_id', '未提取')}")
        print(f"  品牌徽标素材编号: {logo_asset_id}")  # 新增调试打印
        print(f"  视频媒体编号: {video_media_id}")  # 新增调试打印
        
        campaign_name_normalized = str(campaign_name).lower()
        
        # 检测是否 ASIN 活动
        is_asin = 'asin' in campaign_name_normalized
        
        # 检测匹配类别
        matched_category = None
        for cat in keyword_categories:
            if cat in campaign_name_normalized:
                matched_category = cat
                break
        
        # 检测匹配类型（精准/广泛）
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        match_type = '精准' if is_exact else '广泛' if is_broad else '精准'  # 默认 exact
        
        # 生成广告活动行
        row_campaign = [
            product, '广告活动', operation, campaign_name, '', '', '', '', '',
            campaign_name, '', '', status, global_settings.get('budget_type', '每日'), budget,
            global_settings.get('bidding_optimization', '手动'), bid_adjust, '', '', '',
            '', '视频', '',  # 修改：广告格式为'视频'，落地页 URL 为空
            '', global_settings.get('brand_entity_id', ''), global_settings.get('brand_name', ''),
            '',  # 修改：用每个活动的 logo_asset_id，而不是全局
            image_asset, '',
            creative_asin, video_media_id, '视频'  # 修改：视频媒体编号和创意素材类型
        ]
        rows.append(row_campaign)
        
        # 关键词：从df_survey的关键词列拉取（与script-C.py一致的收集逻辑）
        keywords = []
        matched_columns = []
        if matched_category and (is_exact or is_broad):
            for col in keyword_columns:
                if col in df_survey.columns:
                    col_lower = str(col).lower()
                    if matched_category in col_lower:
                        if (is_exact and any(x in col_lower for x in ['精准', 'exact'])) or \
                           (is_broad and any(x in col_lower for x in ['广泛', 'broad'])):
                            matched_columns.append(col)
                            keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
            keywords = list(dict.fromkeys(keywords))
            print(f"  匹配的列: {matched_columns}")
            print(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
        else:
            print("  无匹配的关键词列，关键词为空")
        
        if keywords:
            for kw in keywords:
                row_keyword = [
                    product, '关键词', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', cpc, kw, match_type,
                    '', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_keyword)
        
        # 否定关键词（主要针对广泛，且非ASIN组）
        if not is_asin and matched_category:
            # 新规则：固定列索引（使用实际列名映射）
            col_indices = {
                'Y': df_survey.columns.get_loc('case精准-否精准') if 'case精准-否精准' in df_survey.columns else None,
                'Z': df_survey.columns.get_loc('case精准-否词组') if 'case精准-否词组' in df_survey.columns else None,
                'AC': df_survey.columns.get_loc('case广泛-否精准') if 'case广泛-否精准' in df_survey.columns else None,
                'AD': df_survey.columns.get_loc('case广泛-否词组') if 'case广泛-否词组' in df_survey.columns else None,
                'W': df_survey.columns.get_loc('宿主精准-否精准') if '宿主精准-否精准' in df_survey.columns else None,
                'X': df_survey.columns.get_loc('宿主精准-否词组') if '宿主精准-否词组' in df_survey.columns else None,
                'AA': df_survey.columns.get_loc('宿主广泛-否精准') if '宿主广泛-否精准' in df_survey.columns else None,
                'AB': df_survey.columns.get_loc('宿主广泛-否词组') if '宿主广泛-否词组' in df_survey.columns else None,
            }
            
            # 列名映射，用于日志（使用实际列名）
            col_names = {
                'Y': 'case精准-否精准',
                'Z': 'case精准-否词组',
                'AC': 'case广泛-否精准',
                'AD': 'case广泛-否词组',
                'W': '宿主精准-否精准',
                'X': '宿主精准-否词组',
                'AA': '宿主广泛-否精准',
                'AB': '宿主广泛-否词组'
            }
            
            # 选择列
            selected_cols = []
            if matched_category in ['case', '包']:
                if is_exact:
                    selected_cols = ['Y', 'Z']
                elif is_broad:
                    selected_cols = ['AC', 'AD']
            elif matched_category in ['suzhu', '宿主', 'host']:
                if is_exact:
                    selected_cols = ['W', 'X']
                elif is_broad:
                    selected_cols = ['AA', 'AB']
            
            # 收集数据，按类型分组，使用 defaultdict(list) 跟踪来源
            neg_data_sources = {
                '否定精准匹配': defaultdict(list),  # kw -> [col_keys]
                '否定词组': defaultdict(list)
            }
            for col_key in selected_cols:
                if col_indices.get(col_key) is not None:
                    col_idx = col_indices[col_key]
                    col_data = [str(kw).strip() for kw in df_survey.iloc[:, col_idx].dropna() if str(kw).strip()]
                    col_data = list(dict.fromkeys(col_data))  # 列内去重
                    m_type = '否定精准匹配' if col_key in ['Y', 'AC', 'W', 'AA'] else '否定词组'
                    for kw in col_data:
                        neg_data_sources[m_type][kw].append(col_key)
            
            # 检查重复：如果 kw 有多个来源
            duplicates_detected = False
            for m_type, kw_sources in neg_data_sources.items():
                for kw, sources in kw_sources.items():
                    if len(sources) > 1:
                        duplicates_detected = True
                        source_names = [col_names.get(s, s) for s in sources]
                        print(f"\n=== 检测到重复否定关键词 ===")
                        print(f"活动: {campaign_name}")
                        print(f"类型: {m_type}")
                        print(f"重复关键词: '{kw}'")
                        print(f"来源列: {', '.join(source_names)}")
                        print(f"原因: 该关键词在多个否定列中出现，导致生成重复行。请检查 survey 文件的这些列并清理重复值。")
                        print("暂停生成 header 表。")
                        return None  # 暂停生成
            
            print("\n=== 重复检测完成（无重复）===")
            
            # 生成行：使用去重后的列表
            for m_type, kw_sources in neg_data_sources.items():
                kws = list(kw_sources.keys())  # 已去重
                if kws:
                    print(f"  {m_type} 否定关键词数量: {len(kws)}")
                for kw in kws:
                    row_neg = [
                        product, '否定关键词', operation, campaign_name, '', '', campaign_name, '', '',
                        campaign_name, '', '', status, '', '', '', '', '', kw, m_type,
                        '', '', '', '', '', '', '', '', '', '', ''
                    ]
                    rows.append(row_neg)
        
        # 商品定向（ASIN）：从df_survey的ASIN列（列名匹配活动名称）
        asin_targets = []
        if is_asin:
            # 精确匹配列名
            for col in df_survey.columns:
                if str(col).strip() == str(campaign_name):
                    col_idx = df_survey.columns.get_loc(col)
                    if col_idx is not None:
                        asin_targets = [str(asin).strip() for asin in df_survey.iloc[:, col_idx].dropna() if str(asin).strip()]
                        asin_targets = list(dict.fromkeys(asin_targets))
                        print(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
                        break
        
        if asin_targets:
            for asin in asin_targets:
                row_asin = [
                    product, '商品定向', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', cpc, '', '',
                    f'asin="{asin}"', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_asin)
        
        # 否定商品定向（全局，针对所有ASIN活动）
        if is_asin:
            # 否定ASIN：使用 asin="XXX"
            for asin in neg_asin:
                row_neg_asin = [
                    product, '否定商品定向', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', '', '', '',
                    f'asin="{asin}"', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_neg_asin)
            
            # 否品牌：使用 brand="XXX"
            for brand in neg_brand:
                row_neg_brand = [
                    product, '否定商品定向', operation, campaign_name, '', '', campaign_name, '', '',
                    campaign_name, '', '', status, '', '', '', '', '', '', '',
                    f'brand="{brand}"', '', '', '', '', '', '', '', '', '', ''
                ]
                rows.append(row_neg_brand)
    
    df_header = pd.DataFrame(rows, columns=columns)
    try:
        df_header.to_excel(output_file, index=False, engine='openpyxl')
        print(f"生成完成！输出文件：{output_file}，总行数：{len(rows)}")
        return output_file
    except Exception as e:
        print(f"保存出错：{e}")
        return None

# =============================================================================
# Streamlit App
# =============================================================================
st.title("Header 生成工具")

uploaded_file = st.file_uploader("上传 Survey Excel 文件", type=['xlsx'])

if uploaded_file is not None:
    # 保存上传的文件到临时目录
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        survey_path = tmp_file.name

    st.success(f"文件上传成功: {uploaded_file.name}")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("生成 SB 商品集 Header"):
            with st.spinner("生成中..."):
                output_path = tempfile.mktemp(suffix='-SB商品集.xlsx')
                result = generate_header_from_brand_survey(survey_path, output_path, sheet_name=0)
                if result:
                    with open(result, 'rb') as f:
                        st.download_button(
                            label="下载 header-SB商品集.xlsx",
                            data=f.read(),
                            file_name="header-SB商品集.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("生成失败，请检查控制台日志。")

    with col2:
        if st.button("生成 SBV 品牌旗舰店 Header"):
            with st.spinner("生成中..."):
                output_path = tempfile.mktemp(suffix='-SBV品牌旗舰店.xlsx')
                result = generate_header_for_sbv_brand_store(survey_path, output_path, sheet_name='品牌广告')
                if result:
                    with open(result, 'rb') as f:
                        st.download_button(
                            label="下载 header-SBV品牌旗舰店.xlsx",
                            data=f.read(),
                            file_name="header-SBV品牌旗舰店.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("生成失败，请检查控制台日志。")

    with col3:
        if st.button("生成 SBV 商品详情页 Header"):
            with st.spinner("生成中..."):
                output_path = tempfile.mktemp(suffix='-SBV商品详情页.xlsx')
                result = generate_header_sbv_detail(survey_path, output_path, sheet_name=0)
                if result:
                    with open(result, 'rb') as f:
                        st.download_button(
                            label="下载 header-SBV商品详情页.xlsx",
                            data=f.read(),
                            file_name="header-SBV商品详情页.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("生成失败，请检查控制台日志。")

    # 清理临时文件
    os.unlink(survey_path)
else:
    st.info("请上传 Survey Excel 文件以开始生成。")